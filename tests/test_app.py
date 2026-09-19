"""app.py 的冒烟测试。

用 Qt 的 offscreen 平台跑，不需要人盯着屏幕。测不到"好不好看"，
但能测到"点下去会不会崩"——控件名拼错、信号忘了接、渲染时取到 None，
这类错误等到面试演示时才暴露就太晚了。

跑法（和别的测试一样）：python -m pytest tests/test_app.py
"""

import os

# 必须在任何 QtWidgets 被 import 之前设好，否则会去找显示设备
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

import pytest                      # noqa: E402
import yaml                        # noqa: E402

from conftest import RULES, build_docx   # noqa: E402

pytest.importorskip("PySide6", reason="没装 PySide6（GUI 依赖），跳过界面冒烟")

from PySide6.QtCore import QEventLoop, QTimer          # noqa: E402
from PySide6.QtWidgets import QApplication, QMessageBox  # noqa: E402

import app as app_module                               # noqa: E402
from check_docx import FIXABLE_LABEL, MANUAL_LABEL     # noqa: E402

pytestmark = pytest.mark.skipif(not RULES.exists(), reason="规则文件缺失")


@pytest.fixture(scope="module")
def qapp():
    return QApplication.instance() or QApplication([])


@pytest.fixture(scope="module")
def rules():
    return yaml.safe_load(open(RULES, encoding="utf-8"))


@pytest.fixture
def window(qapp, monkeypatch):
    """建一个真窗口，并把所有弹窗换成记录（模态对话框会把测试挂住）。"""
    seen = {"warn": [], "critical": [], "info": []}
    for name, key in (("warning", "warn"), ("critical", "critical"),
                      ("information", "info")):
        monkeypatch.setattr(
            QMessageBox, name,
            staticmethod(lambda *a, k=key, **kw: seen[k].append(a[1:])))
    win = app_module.MainWindow(rules_path=str(RULES))
    win._popups = seen
    yield win
    win.close()
    win.deleteLater()


def wait_for(job, timeout_ms=60000):
    """等后台任务结束（信号是排队的，得让事件循环转起来）。"""
    loop = QEventLoop()
    job.finished.connect(loop.quit)
    job.failed.connect(loop.quit)
    QTimer.singleShot(timeout_ms, loop.quit)
    loop.exec()
    job.wait(5000)
    return job.isFinished()


def sample(tmp_path, name="a.docx", **kw):
    return build_docx(tmp_path / name, **kw)


# --------------------------------------------------------------------------
def test_窗口能搭起来(window):
    assert window.windowTitle()
    assert window.check_btn.text() == "开始检查"
    assert window.tabs.count() == 4, "四个结果页"
    assert not window.rules_path is None


def test_规则文件被读进来(window, rules):
    assert window.rules is not None, "内置规则没加载上"
    assert window.rules["page"]["margins_cm"]["top"] == 2.5
    assert "规则" in window.rule_label.text()


def test_没选文件就点检查会给提示而不是崩(window):
    window.input_path = None
    window.start_check()
    assert window._popups["warn"], "应该弹一个提示"
    assert window.job is None, "不该启动任务"


def test_检查跑完把结果填进三个页签(window, tmp_path):
    src = sample(tmp_path, margins_cm=(2.54, 2.54, 3.17, 3.17),
                 paragraphs=[("摘  要", None, "center", "黑体", None, 18),
                             ("油茶果实滞育研究。", None, None, "宋体", None, 12),
                             ("图2.1油茶果实", None, "center", "宋体", None, 10.5)])
    window.set_input(str(src))
    assert "a.docx" in window.file_label.text()

    window.start_check()
    assert window.job is not None
    assert wait_for(window.job), "任务没跑完"
    assert not window._popups["critical"], f"不该报错：{window._popups['critical']}"

    # 重点问题
    assert window.pri_tree.topLevelItemCount() > 0, "重点问题页是空的"
    # 完整清单：按分类折起来，应该有若干顶层分组
    assert window.full_tree.topLevelItemCount() > 1
    # 表格里"能否自动修"那一列要有内容
    first = window.pri_tree.topLevelItem(0)
    assert first.text(4) in (FIXABLE_LABEL, MANUAL_LABEL)
    # 文档概况里应带上"尚未实现"
    assert "尚未实现" in window.overview_view.toPlainText()
    # 状态栏给出计数
    assert "检查完成" in window.status.text()
    # 检查完之后才能保存报告
    assert window.save_report_btn.isEnabled()


def test_完整清单的筛选框能用(window, tmp_path):
    src = sample(tmp_path, paragraphs=[("随便正文。",)])
    window.set_input(str(src))
    window.start_check()
    assert wait_for(window.job)

    total = window.full_tree.topLevelItemCount()
    assert total > 1
    def visible_groups():
        return [window.full_tree.topLevelItem(i)
                for i in range(window.full_tree.topLevelItemCount())
                if not window.full_tree.topLevelItem(i).isHidden()]

    window.filter_box.setText("缺")
    assert len(visible_groups()) < total, "搜字之后该少一些"
    window.filter_box.setText("")
    assert len(visible_groups()) == total, "清空搜索该复原"

    # 下拉：只看"建议手动修改"的
    window.filter_mode.setCurrentIndex(2)
    kids = [g.child(j) for g in visible_groups()
            for j in range(g.childCount()) if not g.child(j).isHidden()]
    assert kids, "该还剩一些手动项"
    assert all(k.text(4) == MANUAL_LABEL for k in kids)

    window.filter_mode.setCurrentIndex(1)
    kids2 = [g.child(j) for g in visible_groups()
             for j in range(g.childCount()) if not g.child(j).isHidden()]
    assert kids2 and all(k.text(4) == FIXABLE_LABEL for k in kids2)

    window.filter_mode.setCurrentIndex(0)
    assert len(visible_groups()) == total


def test_格式化预览不写文件(window, tmp_path):
    src = sample(tmp_path, paragraphs=[
        ("图2.1油茶果实", None, "center", "宋体", None, 10.5)])
    window.set_input(str(src))
    window.dry_box.setChecked(True)
    window.start_format()
    assert wait_for(window.job)
    assert not window._popups["critical"], window._popups["critical"]

    # 改动清单页应该有内容，而且明确说是预览
    text = window.change_view.toPlainText()
    assert "改动" in text or "预览" in text
    assert "预览" in window.status.text() or "预览" in text
    # 目录里不该多出文件
    assert sorted(p.name for p in tmp_path.glob("*.docx")) == ["a.docx"]


def test_任务跑着的时候按钮是禁用的(window, tmp_path):
    src = sample(tmp_path, paragraphs=[("正文。",)])
    window.set_input(str(src))
    window.check_btn.setEnabled(False)
    window._set_busy(True, "正在跑")
    assert not window.check_btn.isEnabled()
    assert not window.format_btn.isEnabled()
    window._set_busy(False)
    assert window.check_btn.isEnabled()
    assert window.format_btn.isEnabled()


def test_出错时给中文提示不崩(window, tmp_path):
    bad = tmp_path / "假的.docx"
    bad.write_text("我不是 docx", encoding="utf-8")
    window.set_input(str(bad))
    window.start_check()
    assert wait_for(window.job)
    assert window._popups["critical"], "应该弹错误框"
    msg = window._popups["critical"][0][1]
    assert "docx" in msg
    # 出错后按钮要恢复，不能把人卡住
    assert window.check_btn.isEnabled()
