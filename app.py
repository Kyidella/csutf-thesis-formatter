"""论文格式工具 —— 图形界面。

命令行能干的事这里都能干：拖入 .docx → 检查（只读）或格式化（另存）→ 看结果。

几条刻意的设计，都是有原因的：

  * **没有取消按钮。** load() 内部没有中断点，QThread.terminate() 会直接杀线程，
    而写文件不是原子的——杀在半路会留下一个残缺的 .docx。慢一点没关系，
    留个坏文件就麻烦了。运行期间禁用按钮，用进度条 + 计时器证明还活着。
  * **不解析 Markdown 报告。** 界面直接吃 run_check 返回的结构化数据，
    省一次解析，也不会因为报告格式改了而界面跟着坏。
  * **结果分三页。** 一次检查能出 90 条，平铺会淹；重点问题单列，
    完整清单按分类折起来，再加一个搜索框。

用法:
    python app.py
"""

import html
import os
import sys
from pathlib import Path

# --windowed 打包后没有控制台，stdout/stderr 是 None，
# 任何一句残留的 print 都会炸成 AttributeError。先兜住。
if sys.stdout is None:
    sys.stdout = open(os.devnull, "w", encoding="utf-8")
if sys.stderr is None:
    sys.stderr = open(os.devnull, "w", encoding="utf-8")

from PySide6.QtCore import Qt, QTimer                            # noqa: E402
from PySide6.QtGui import QAction, QColor, QFont                 # noqa: E402
from PySide6.QtWidgets import (QAbstractItemView, QApplication,  # noqa: E402
                               QCheckBox, QComboBox, QFileDialog, QHBoxLayout,
                               QHeaderView, QLabel, QLineEdit, QMainWindow,
                               QMessageBox, QProgressBar, QPushButton,
                               QTabWidget, QTextBrowser, QTreeWidget,
                               QTreeWidgetItem, QVBoxLayout, QWidget)

from check_docx import (FIXABLE_LABEL, MANUAL_LABEL,  # noqa: E402
                        FIX_KIND, fix_kind)
from gui_worker import (CheckPayload, JobThread, default_rules_path,  # noqa: E402
                        load_rules, setup_logging, suggest_output_path)

LEVEL_STYLE = {
    "error": ("错误", "#c0392b"),
    "warn": ("警告", "#b9770e"),
    "info": ("提示", "#2471a3"),
}


def human_size(n):
    for unit in ("B", "KB", "MB", "GB"):
        if n < 1024 or unit == "GB":
            return f"{n:.0f} {unit}" if unit == "B" else f"{n:.1f} {unit}"
        n /= 1024


# --------------------------------------------------------------------------
# 拖拽区
# --------------------------------------------------------------------------
class DropZone(QLabel):
    """把 .docx 拖进来。点一下也能选文件。"""

    def __init__(self, on_pick, on_file):
        super().__init__("把论文 .docx 拖到这里\n（或者点一下选文件）")
        self.on_pick = on_pick
        self.on_file = on_file
        self.setAlignment(Qt.AlignCenter)
        self.setAcceptDrops(True)
        self.setMinimumHeight(96)
        self.setCursor(Qt.PointingHandCursor)
        self._idle()

    def _idle(self):
        self.setStyleSheet(
            "QLabel { border: 2px dashed #9aa4ad; border-radius: 8px;"
            " color: #5b6770; padding: 10px; }")

    def _hot(self):
        self.setStyleSheet(
            "QLabel { border: 2px dashed #2e86c1; border-radius: 8px;"
            " background: #eaf4fb; color: #1b4f72; padding: 10px; }")

    def mousePressEvent(self, event):
        self.on_pick()

    def dragEnterEvent(self, event):
        if not self.isEnabled():
            event.ignore()
            return
        urls = event.mimeData().urls()
        if urls and urls[0].toLocalFile().lower().endswith(".docx"):
            event.acceptProposedAction()
            self._hot()
        else:
            event.ignore()

    def dragLeaveEvent(self, event):
        self._idle()

    def dropEvent(self, event):
        self._idle()
        for url in event.mimeData().urls():
            path = url.toLocalFile()
            if path.lower().endswith(".docx"):
                self.on_file(path)
                event.acceptProposedAction()
                return
        event.ignore()


# --------------------------------------------------------------------------
# 主窗口
# --------------------------------------------------------------------------
class MainWindow(QMainWindow):
    def __init__(self, rules_path=None):
        super().__init__()
        self.setWindowTitle("论文格式工具 — 中南林业科技大学硕士学位论文")
        self.resize(1060, 760)

        self.input_path = None
        self.rules_path = rules_path or default_rules_path()
        self.rules = None
        self.job = None
        self._last_check = None

        self._build_ui()
        self._load_rules_silently()
        self._timer = QTimer(self)
        self._timer.setInterval(1000)
        self._timer.timeout.connect(self._tick)
        self._elapsed = 0

    # ---- 界面 ----
    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(14, 12, 14, 12)
        root.setSpacing(10)

        self.drop = DropZone(self.pick_input, self.set_input)
        root.addWidget(self.drop)

        info = QHBoxLayout()
        self.file_label = QLabel("尚未选择文件")
        self.file_label.setStyleSheet("color:#5b6770;")
        info.addWidget(self.file_label, 1)
        self.rule_label = QLabel()
        self.rule_label.setStyleSheet("color:#5b6770;")
        info.addWidget(self.rule_label)
        switch = QPushButton("换规则文件…")
        switch.clicked.connect(self.pick_rules)
        info.addWidget(switch)
        root.addLayout(info)

        actions = QHBoxLayout()
        self.check_btn = QPushButton("开始检查")
        self.check_btn.setMinimumHeight(38)
        self.check_btn.clicked.connect(self.start_check)
        actions.addWidget(self.check_btn)

        self.format_btn = QPushButton("自动排版并另存…")
        self.format_btn.setMinimumHeight(38)
        self.format_btn.clicked.connect(self.start_format)
        actions.addWidget(self.format_btn)

        self.dry_box = QCheckBox("仅预览（不写文件）")
        actions.addWidget(self.dry_box)

        actions.addStretch(1)
        self.save_report_btn = QPushButton("保存检查报告…")
        self.save_report_btn.setEnabled(False)
        self.save_report_btn.clicked.connect(self.save_report)
        actions.addWidget(self.save_report_btn)
        root.addLayout(actions)

        self.tabs = QTabWidget()
        root.addWidget(self.tabs, 1)

        self.pri_tree = self._make_tree(["级别", "分类", "说明", "位置", "能否自动排版"])
        self.tabs.addTab(self.pri_tree, "重点问题")

        full = QWidget()
        fl = QVBoxLayout(full)
        fl.setContentsMargins(0, 6, 0, 0)
        frow = QHBoxLayout()
        self.filter_box = QLineEdit()
        self.filter_box.setPlaceholderText("搜关键字，如：字体、缩进、页边距、图2.1 …")
        self.filter_box.textChanged.connect(self._apply_filter)
        frow.addWidget(self.filter_box, 1)
        self.filter_mode = QComboBox()
        self.filter_mode.addItems(["显示：全部", FIXABLE_LABEL, MANUAL_LABEL])
        self.filter_mode.setToolTip("按「这条我需不需要自己动手」筛")
        self.filter_mode.currentIndexChanged.connect(self._apply_filter)
        frow.addWidget(self.filter_mode)
        fl.addLayout(frow)
        self.full_tree = self._make_tree(["级别", "分类", "说明", "位置", "能否自动排版"])
        fl.addWidget(self.full_tree, 1)
        self.tabs.addTab(full, "完整清单")

        self.change_view = QTextBrowser()
        self.change_view.setOpenExternalLinks(False)
        self.tabs.addTab(self.change_view, "改动清单")

        self.overview_view = QTextBrowser()
        self.tabs.addTab(self.overview_view, "文档概况")

        bottom = QHBoxLayout()
        self.bar = QProgressBar()
        self.bar.setTextVisible(False)
        self.bar.setMaximumHeight(10)
        bottom.addWidget(self.bar, 1)
        self.status = QLabel("就绪")
        self.status.setMinimumWidth(320)
        bottom.addWidget(self.status)
        root.addLayout(bottom)

        pick = QAction("打开…", self)
        pick.setShortcut("Ctrl+O")
        pick.triggered.connect(self.pick_input)
        self.addAction(pick)

    # 各列初始宽度。别用 QHeaderView.Stretch：只要有一列是 Stretch，
    # 跟它相邻的那条边界就拖不动了（用户实测："位置"前面那条拖不了，
    # 拖"能否自动排版"前面那条却会带着"位置"一起动）。
    COL_WIDTHS = (60, 90, 520, 300, 110)

    def _make_tree(self, headers):
        t = QTreeWidget()
        t.setColumnCount(len(headers))
        t.setHeaderLabels(headers)
        t.setAlternatingRowColors(True)
        t.setSelectionBehavior(QAbstractItemView.SelectRows)
        t.setUniformRowHeights(True)
        h = t.header()
        h.setSectionResizeMode(QHeaderView.Interactive)   # 每列都能拖
        h.setStretchLastSection(False)
        for i, w in enumerate(self.COL_WIDTHS[:len(headers)]):
            t.setColumnWidth(i, w)
        return t

    # ---- 规则 ----
    def _load_rules_silently(self):
        try:
            self.rules = load_rules(self.rules_path)
            self.rule_label.setText(f"规则：{Path(self.rules_path).name}")
        except Exception as exc:                       # noqa: BLE001
            self.rules = None
            self.rule_label.setText("规则：加载失败")
            self._warn("规则文件读不了", str(exc))

    def pick_rules(self):
        path, _ = QFileDialog.getOpenFileName(self, "选择规则文件", "",
                                              "规则文件 (*.yaml *.yml)")
        if path:
            self.rules_path = path
            self._load_rules_silently()

    # ---- 选文件 ----
    def pick_input(self):
        path, _ = QFileDialog.getOpenFileName(self, "选择论文", "", "Word 文档 (*.docx)")
        if path:
            self.set_input(path)

    def set_input(self, path):
        self.input_path = path
        size = human_size(Path(path).stat().st_size)
        self.file_label.setText(f"{Path(path).name}（{size}）")
        self.save_report_btn.setEnabled(False)
        self._idle_status()

    # ---- 运行 ----
    def _guard(self):
        if self.job is not None and self.job.isRunning():
            return False
        if not self.input_path:
            self._warn("还没选文件", "先把论文 .docx 拖进来，或者点上面的框选一个。")
            return False
        if not self.rules:
            self._warn("规则文件没加载上", "换一个规则文件再试。")
            return False
        return True

    def _set_busy(self, busy, note=""):
        self.check_btn.setEnabled(not busy)
        self.format_btn.setEnabled(not busy)
        self.dry_box.setEnabled(not busy)
        self.save_report_btn.setEnabled(
            not busy and self._last_check is not None)
        self.drop.setEnabled(not busy)
        if busy:
            self._elapsed = 0
            self.bar.setRange(0, 100)
            self.bar.setValue(0)
            self._timer.start()
            self.status.setText(note or "开始…")
        else:
            self._timer.stop()
            if not note:
                self.status.setText("就绪")

    def _tick(self):
        self._elapsed += 1
        m, s = divmod(self._elapsed, 60)
        base = self.status.text().split("（已用")[0]
        self.status.setText(f"{base}（已用 {m:02d}:{s:02d}）")

    def start_check(self):
        if not self._guard():
            return
        self._last_check = None
        self.job = JobThread("check", self.input_path, self.rules,
                             rules_path=self.rules_path, parent=self)
        self._wire(self.job)
        self._set_busy(True, "正在检查…（大文件要一两分钟）")
        self.job.start()

    def start_format(self):
        if not self._guard():
            return
        dry = self.dry_box.isChecked()
        output = None
        if not dry:
            default = suggest_output_path(self.input_path)
            output, _ = QFileDialog.getSaveFileName(self, "另存为", default,
                                                    "Word 文档 (*.docx)")
            if not output:
                return
            if Path(output).resolve() == Path(self.input_path).resolve():
                self._warn("不能原地修改", "输出文件不能和原文件是同一个，换个名字。")
                return
        self.job = JobThread("format", self.input_path, self.rules,
                             output_path=output, dry_run=dry, parent=self)
        self._wire(self.job)
        self._set_busy(True, "正在自动排版…（大文件要一两分钟）")
        self.job.start()

    def _wire(self, job):
        job.stage.connect(self._on_stage)
        job.finished.connect(self._on_finished)
        job.failed.connect(self._on_failed)

    def _on_stage(self, msg, done, total):
        pct = int(done / total * 100) if total else 0
        self.bar.setValue(max(0, min(100, pct)))
        self.status.setText(msg)

    def _on_failed(self, name, msg):
        self._set_busy(False, "出错了")
        QMessageBox.critical(self, "出错了",
                             f"{msg}\n\n（技术细节：{name}，完整堆栈在日志里）")

    def _on_finished(self, payload):
        self._set_busy(False)
        if isinstance(payload, CheckPayload):
            self.show_check(payload)
        else:
            self.show_format(payload)

    # ---- 渲染：检查 ----
    def show_check(self, p):
        self.pri_tree.clear()
        self.full_tree.clear()
        self._fill_tree(self.pri_tree, p.rep.priority())
        self._fill_full(p.rep.items)

        n = {lv: p.rep.count(lv) for lv in ("error", "warn", "info")}
        self.status.setText(
            f"检查完成：错误 {n['error']} · 警告 {n['warn']} · 提示 {n['info']}")
        self.tabs.setCurrentIndex(0)

        self.overview_view.setHtml(self._overview_html(p))
        self.change_view.setHtml(
            "<p style='color:#5b6770'>这次是检查（只读），没有改动。"
            "要看会改什么，点「自动排版并另存…」并勾上「仅预览」。</p>")
        self.save_report_btn.setEnabled(True)
        self._last_check = p

    def _fill_tree(self, tree, items):
        for it in items:
            tree.addTopLevelItem(self._row(it))

    def _fill_full(self, items):
        """按分类折起来。90 条平铺会淹，分成十几组就不会。"""
        self.full_tree.clear()
        groups = {}
        for it in items:
            groups.setdefault(it["category"], []).append(it)
        for cat in sorted(groups, key=lambda c: (-len(groups[c]), c)):
            parent = QTreeWidgetItem([f"{cat}（{len(groups[cat])}）", "", "", "", ""])
            f = parent.font(0)
            f.setBold(True)
            parent.setFont(0, f)
            for it in groups[cat]:
                parent.addChild(self._row(it))
            self.full_tree.addTopLevelItem(parent)
        self.full_tree.expandAll()

    def _row(self, it):
        label, color = LEVEL_STYLE.get(it["level"], ("", "#000000"))
        kind = fix_kind(it["category"], it.get("key"))
        hint = FIXABLE_LABEL if kind else MANUAL_LABEL
        row = QTreeWidgetItem([label, it["category"], it["msg"], it["loc"], hint])
        row.setForeground(0, QColor(color))
        row.setForeground(4, QColor("#1e8449" if kind else "#7f8c8d"))
        row.setToolTip(2, it["msg"])
        row.setToolTip(3, it["loc"])
        row.setToolTip(4, f"工具会自动处理（{kind}）" if kind
                           else "工具不会自动改，需要你自己处理")
        # 能不能自动修要看规则键，从这列的文字（分类）反推不出来，存下来给筛选用。
        row.setData(4, Qt.UserRole, bool(kind))
        return row

    def _apply_filter(self, *_):
        """搜索框（搜所有列）+ 下拉（要不要我自己动手），两个条件是"与"。

        搜索范围覆盖全部五列——位置那列也有内容可搜（段落开头几个字），
        写死只搜分类和说明会让人以为搜不到。
        """
        text = self.filter_box.text().strip()
        mode = self.filter_mode.currentIndex()      # 0 全部 / 1 支持 / 2 建议手动
        want = None if mode == 0 else (mode == 1)

        for i in range(self.full_tree.topLevelItemCount()):
            parent = self.full_tree.topLevelItem(i)
            shown = 0
            for j in range(parent.childCount()):
                child = parent.child(j)
                hit = (not text) or any(
                    text in child.text(c) for c in range(5))
                if hit and want is not None:
                    hit = bool(child.data(4, Qt.UserRole)) == want
                child.setHidden(not hit)
                shown += hit
            parent.setHidden(shown == 0)

    def _overview_html(self, p):
        s = p.doc.stats
        rows = [("段落总数", s["paragraphs_total"]),
                ("正文流", s["paragraphs_body"]),
                ("表格内", s["paragraphs_in_table"]),
                ("内容控件内", s["paragraphs_in_sdt"]),
                ("节", s["sections"]),
                ("含引用域的段落", s["paragraphs_with_fields"])]
        body = "".join(
            f"<tr><td style='padding:3px 16px 3px 0;color:#5b6770'>{k}</td>"
            f"<td><b>{v}</b></td></tr>" for k, v in rows)
        head = ("<h3 style='margin:0 0 8px'>文档概况</h3>"
                f"<table>{body}</table>")

        nk = p.unimplemented or []
        if nk:
            items = "".join(
                f"<li><code>{html.escape(str(e.get('key','')))}</code> — "
                f"{html.escape(str(e.get('desc','')))}<br>"
                f"<span style='color:#7f8c8d'>{html.escape(str(e.get('why','')))}"
                f"</span></li>" for e in nk)
            head += (f"<h3 style='margin:18px 0 6px'>规范要求、本工具尚未实现"
                     f"（{len(nk)} 项）</h3>"
                     "<p style='color:#5b6770;margin:4px 0'>这些条目写在规则文件里，"
                     "但没有对应的检查或修正代码。列出来是为了不让人以为它们已经在查了。</p>"
                     f"<ul style='line-height:1.5'>{items}</ul>")
        return head

    # ---- 渲染：格式化 ----
    def show_format(self, p):
        groups = [("节级格式", []), ("样式", []), ("文本层", [])]
        for c in p.changes:
            cls = type(c).__name__
            groups[0 if cls == "Change" else 1 if cls == "StyleChange" else 2][1].append(c)

        if p.dry_run:
            banner = ("<div style='background:#fef9e7;border:1px solid #f7dc6f;"
                      "padding:8px;border-radius:4px'>"
                      "<b>预览：没有写文件。</b>下面是正式跑会改的东西。</div>")
        else:
            banner = (f"<div style='background:#eafaf1;border:1px solid #abebc6;"
                      f"padding:8px;border-radius:4px'><b>已写入</b> "
                      f"<code>{html.escape(p.output_path or '')}</code></div>")

        miss = (p.stat or {}).get("missing_sections") or []
        if miss:
            banner += ("<div style='background:#fdecea;border:1px solid #f5b7b1;"
                       "padding:8px;border-radius:4px;margin-top:6px'>"
                       "<b>这份文档不完整：</b>没有检测到 "
                       f"{html.escape('、'.join(miss))}。"
                       "自动排版只整理已有的内容，缺的部分不会凭空出现；"
                       "而且章节定位依赖这些内容，缺得越多、排版结果越不可靠。"
                       "</div>")
        parts = [banner]
        for title, items in groups:
            if not items:
                continue
            lis = "".join(f"<li>{html.escape(c.render())}</li>" for c in items)
            parts.append(f"<h3 style='margin:14px 0 4px'>{title}（{len(items)} 条）</h3>"
                         f"<ul style='line-height:1.6'>{lis}</ul>")

        if p.notes:
            lis = "".join(f"<li>{html.escape(str(n))}</li>" for n in p.notes)
            parts.append("<h3 style='margin:14px 0 4px;color:#7f8c8d'>没动的部分"
                         "（有意跳过）</h3>"
                         f"<ul style='color:#7f8c8d;line-height:1.6'>{lis}</ul>")

        if not p.changes:
            parts.append("<p>这份文档已经合规，没有要改的地方。</p>")

        self.change_view.setHtml("".join(parts))
        self.tabs.setCurrentIndex(2)

        total = len(p.changes)
        if p.dry_run:
            self.status.setText(f"预览完成：会改 {total} 处（没写文件）")
        elif total:
            self.status.setText(f"自动排版完成：改了 {total} 处 → {Path(p.output_path).name}")
        else:
            self.status.setText("自动排版完成：没有要改的地方")

    # ---- 报告 ----
    def save_report(self):
        p = self._last_check
        if p is None:
            return
        default = str(Path(self.input_path).with_name(
            Path(self.input_path).stem + "_检查报告.md"))
        path, _ = QFileDialog.getSaveFileName(self, "保存报告", default,
                                              "Markdown (*.md)")
        if not path:
            return
        try:
            from check_docx import write_report
            write_report(p.rep, p.doc, p.rows, p.input_path, self.rules,
                         self.rules_path, path)
            self.status.setText(f"报告已保存：{Path(path).name}")
        except Exception as exc:                       # noqa: BLE001
            self._warn("报告没保存成功", str(exc))

    def _idle_status(self):
        self.status.setText("就绪。选「开始检查」看问题，选「自动排版并另存」做修正。")
        self.bar.setValue(0)

    def _warn(self, title, text):
        QMessageBox.warning(self, title, text)

    def closeEvent(self, event):
        if self.job is not None and self.job.isRunning():
            QMessageBox.information(
                self, "还在跑",
                "任务正在运行，等它结束再关。\n"
                "中途强杀会留下不完整的文件，所以这里不做强制退出。")
            event.ignore()
            return
        event.accept()


def main():
    log_path = setup_logging()
    app = QApplication(sys.argv)
    app.setApplicationName("论文格式工具")
    app.setFont(QFont("Microsoft YaHei", 9))
    win = MainWindow()
    win.show()
    print(f"日志：{log_path}")
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
