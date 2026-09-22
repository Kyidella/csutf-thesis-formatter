"""GUI 的后台线程与错误翻译。

这里**不 import 任何 Qt 控件**（只 import QtCore 的 QThread/Signal），
所以它能被无头测试，逻辑也不会和界面缠在一起。

为什么要独立线程：一篇 200MB 的论文 load + classify 要一两分钟，
放主线程界面会整个卡死（"未响应"）。
线程间只通过 Signal 传数据——`emit` 是线程安全的，会自动排队回 GUI 线程。
核心模块（docmodel / check_docx / formatter）完全不知道 Qt 的存在。

**没有取消功能，这是有意的。** load() 内部没有中断点，而 QThread.terminate()
会直接杀线程——formatter._write 不是原子写，杀在半路会留下一个残缺的 .docx，
比慢一点糟糕得多。所以运行期间禁用按钮，用进度条和计时器证明程序还活着。
"""

import logging
import os
import sys
import tempfile
import traceback
from dataclasses import dataclass, field
from pathlib import Path

from PySide6.QtCore import QThread, Signal

import yaml

from check_docx import (OutputPathError, friendly_error,
                        not_implemented_items, run_check)
from formatter import apply


# --------------------------------------------------------------------------
# 结果载荷
# --------------------------------------------------------------------------
@dataclass
class CheckPayload:
    """检查的结果。界面直接从这些结构化数据渲染，不去解析 Markdown 报告。"""

    input_path: str
    rep: object
    doc: object
    rows: list = field(default_factory=list)
    report_path: str | None = None
    unimplemented: list = field(default_factory=list)


@dataclass
class FormatPayload:
    """格式化的结果。"""

    input_path: str
    output_path: str | None
    changes: list = field(default_factory=list)
    notes: list = field(default_factory=list)
    stat: dict = field(default_factory=dict)
    dry_run: bool = False


# --------------------------------------------------------------------------
# 后台线程
# --------------------------------------------------------------------------
class JobThread(QThread):
    """跑一次检查或一次格式化。同一时刻只应有一个在跑。"""

    stage = Signal(str, int, int)      # 说明, 已完成, 总数
    finished = Signal(object)          # CheckPayload 或 FormatPayload
    failed = Signal(str, str)          # 异常类名, 中文提示

    def __init__(self, kind, input_path, rules, output_path=None,
                 dry_run=False, report_path=None, rules_path=None, parent=None):
        super().__init__(parent)
        self.kind = kind               # "check" 或 "format"
        self.input_path = str(input_path)
        self.rules = rules
        self.rules_path = rules_path
        self.output_path = str(output_path) if output_path else None
        self.dry_run = dry_run
        self.report_path = report_path
        self.log = logging.getLogger("thesis_formatter")

    def _emit_stage(self, msg, done, total):
        self.stage.emit(msg, done, total)

    def run(self):
        try:
            if self.kind == "check":
                self.finished.emit(self._do_check())
            else:
                self.finished.emit(self._do_format())
        except Exception as exc:                      # noqa: BLE001 - 兜底，界面不能崩
            name, msg = friendly_error(exc)
            self.log.error("任务失败：%s\n%s", name, traceback.format_exc())
            self.failed.emit(name, msg)

    def _do_check(self):
        rep, doc, rows = run_check(self.input_path, self.rules,
                                   report_path=self.report_path,
                                   rules_path=self.rules_path,
                                   on_stage=self._emit_stage)
        return CheckPayload(input_path=self.input_path, rep=rep, doc=doc,
                            rows=rows, report_path=self.report_path,
                            unimplemented=not_implemented_items(self.rules))

    def _do_format(self):
        """先写到同目录的临时文件，成功了再原子替换过去。

        formatter._write 不是原子写，中途失败会留下半截 docx——
        对一个论文文件来说，宁可没有输出，也不能留个坏文件在那。
        """
        target = Path(self.output_path) if self.output_path else None
        if target is None:
            changes, notes, _, stat = apply(self.input_path, self.rules,
                                            None, True, on_stage=self._emit_stage)
            return FormatPayload(self.input_path, None, changes, notes, stat, True)

        # 提前说清楚，别等到写入时才抛一个含糊的 FileNotFoundError
        if not target.parent.exists():
            raise OutputPathError(f"输出目录不存在：{target.parent}")
        if Path(self.input_path).resolve() == target.resolve():
            raise OutputPathError("输出文件不能和原文件是同一个，换个名字。")

        tmp = target.with_name(target.name + ".part")
        try:
            changes, notes, _, stat = apply(self.input_path, self.rules,
                                            str(tmp), self.dry_run,
                                            on_stage=self._emit_stage)
            if self.dry_run or not changes:
                return FormatPayload(self.input_path, None, changes, notes,
                                     stat, self.dry_run)
            os.replace(tmp, target)          # 同盘符下是原子的
            return FormatPayload(self.input_path, str(target), changes, notes,
                                 stat, False)
        finally:
            if tmp.exists():
                try:
                    tmp.unlink()
                except OSError:
                    self.log.warning("临时文件没能删掉：%s", tmp)


# --------------------------------------------------------------------------
# 打包后的路径与日志
# --------------------------------------------------------------------------
def app_dir():
    """程序所在目录。打包成单文件 exe 后是解包到临时的 _MEIPASS。"""
    if getattr(sys, "frozen", False):
        return Path(sys._MEIPASS)
    return Path(__file__).resolve().parent


def default_rules_path():
    """找内置规则：规则目录里第一个 .yaml。

    刻意不写死那个中文文件名——中文路径进 PyInstaller 的 --add-data 最容易翻车，
    所以打包时复制成 ASCII 名（见 build_exe.py），这里按扩展名找。
    """
    d = app_dir() / "rules"
    if d.is_dir():
        for p in sorted(d.glob("*.yaml")):
            return p
    # 源码运行时兜底：仓库里的规则目录
    d = Path(__file__).resolve().parent / "rules"
    if d.is_dir():
        for p in sorted(d.glob("*.yaml")):
            return p
    return None


def load_rules(path):
    with open(path, encoding="utf-8") as f:
        return yaml.safe_load(f)


def setup_logging():
    """日志写到 %LOCALAPPDATA%\\ThesisFormatter\\log.txt。

    --windowed 的 exe 没有控制台，出错时这是唯一的排查入口。
    """
    base = Path(os.environ.get("LOCALAPPDATA") or tempfile.gettempdir())
    d = base / "ThesisFormatter"
    d.mkdir(parents=True, exist_ok=True)
    path = d / "log.txt"
    logging.basicConfig(
        filename=str(path), filemode="a", encoding="utf-8",
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(name)s %(message)s")
    return path


def suggest_output_path(input_path):
    """默认输出名。必须与原文件不同名——_write 会拒绝原地修改。"""
    p = Path(input_path)
    return str(p.with_name(f"{p.stem}_已排版{p.suffix}"))
