"""gui_worker 的测试。

刻意不测界面：worker 里只有线程调度和错误翻译，这两样都能无头测。
界面那层（app.py）是真机手动看的，测不到的就别假装测了。
每条测试都在一个真的 QThread 上跑，走的是界面同一条路。
"""

import pathlib
import zipfile

import pytest
import yaml

from conftest import RULES, build_docx

pytest.importorskip("PySide6", reason="没装 PySide6（GUI 依赖），跳过 worker 测试")

pytestmark = pytest.mark.skipif(not RULES.exists(), reason="规则文件缺失")

from PySide6.QtCore import QCoreApplication, QEventLoop   # noqa: E402

from gui_worker import (JobThread, default_rules_path,       # noqa: E402
                        friendly_error, suggest_output_path)


@pytest.fixture(scope="module")
def rules():
    return yaml.safe_load(open(RULES, encoding="utf-8"))


@pytest.fixture(scope="module")
def app():
    """QThread 需要一个 QCoreApplication 才转得起来（不需要窗口）。"""
    a = QCoreApplication.instance() or QCoreApplication([])
    yield a


def run_job(app, job):
    """启动线程并等它结束，收集三个信号发出来的东西。"""
    out = {"stage": [], "finished": None, "failed": None}
    loop = QEventLoop()
    job.stage.connect(lambda m, d, t: out["stage"].append((m, d, t)))
    job.finished.connect(lambda p: (out.update(finished=p), loop.quit()))
    job.failed.connect(lambda n, m: (out.update(failed=(n, m)), loop.quit()))
    job.start()
    loop.exec()
    job.wait(5000)
    return out


# ---- 检查 ----
def test_检查跑通并推阶段(app, rules, tmp_path):
    src = build_docx(tmp_path / "a.docx",
                     paragraphs=[("摘  要", None, "center", "黑体", None, 18),
                                 ("正文内容。", None, None, "宋体", None, 12)])
    out = run_job(app, JobThread("check", src, rules))
    assert out["failed"] is None, out["failed"]
    assert out["stage"], "一个阶段都没推"
    assert out["stage"][0][1] == 0 and out["stage"][-1][1] == out["stage"][-1][2]

    p = out["finished"]
    assert p.rep.items, "该报出问题"
    assert p.doc.stats["paragraphs_total"] >= 2
    assert p.unimplemented, "未实现清单应该带出来"


def test_检查结果能按分类取用(app, rules, tmp_path):
    """界面要按分类分组渲染，这里先确认数据形状对得上。"""
    src = build_docx(tmp_path / "a.docx", paragraphs=[("随便正文。",)])
    p = run_job(app, JobThread("check", src, rules))["finished"]
    cats = {i["category"] for i in p.rep.items}
    assert "缺失章节" in cats
    for i in p.rep.items:
        assert {"level", "category", "msg", "loc"} <= set(i)


# ---- 格式化 ----
def test_格式化写出文件且不原地修改(app, rules, tmp_path):
    src = build_docx(tmp_path / "a.docx",
                     paragraphs=[("图2.1油茶果实", None, "center", "宋体", None, 10.5)])
    before = src.read_bytes()
    dst = tmp_path / "b.docx"

    out = run_job(app, JobThread("format", src, rules, output_path=dst))
    assert out["failed"] is None, out["failed"]
    p = out["finished"]
    assert p.output_path == str(dst) and dst.exists()
    assert p.changes, "应该报出改动"
    assert src.read_bytes() == before, "原文件一个字节都不该动"
    assert not list(tmp_path.glob("*.part")), "临时文件要清干净"


def test_预演不写文件(app, rules, tmp_path):
    src = build_docx(tmp_path / "a.docx",
                     paragraphs=[("图2.1油茶果实", None, "center", "宋体", None, 10.5)])
    dst = tmp_path / "b.docx"
    out = run_job(app, JobThread("format", src, rules, output_path=dst, dry_run=True))
    p = out["finished"]
    assert p.dry_run and not dst.exists()
    assert p.changes, "预演也要把会改什么列出来"


def test_输出路径与原文件相同会被拦下(app, rules, tmp_path):
    src = build_docx(tmp_path / "a.docx",
                     paragraphs=[("图2.1油茶果实", None, "center", "宋体", None, 10.5)])
    out = run_job(app, JobThread("format", src, rules, output_path=src))
    assert out["failed"] is not None, "原地修改必须被拒绝"
    assert out["finished"] is None
    assert src.exists(), "原文件还在"


def test_失败不留半截文件(app, rules, tmp_path):
    """写一半失败时，目标位置不能出现一个残缺的 docx。"""
    src = build_docx(tmp_path / "a.docx",
                     paragraphs=[("图2.1油茶果实", None, "center", "宋体", None, 10.5)])
    dst = tmp_path / "sub" / "b.docx"      # 目录不存在 -> 写入必失败
    out = run_job(app, JobThread("format", src, rules, output_path=dst))
    assert out["failed"] is not None
    assert not dst.exists()
    assert not list(tmp_path.glob("*.part"))


# ---- 错误翻译 ----
def test_非docx文件给出人话(app, rules, tmp_path):
    bad = tmp_path / "not.docx"
    bad.write_text("我不是 docx", encoding="utf-8")
    out = run_job(app, JobThread("check", bad, rules))
    assert out["failed"] is not None
    name, msg = out["failed"]
    assert name == "BadZipFile"
    assert "docx" in msg


def test_文件不存在给出人话(app, rules, tmp_path):
    out = run_job(app, JobThread("check", tmp_path / "没有这个.docx", rules))
    assert out["failed"][1].startswith("找不到文件")


def test_异常翻译表覆盖常见错误():
    assert "docx" in friendly_error(zipfile.BadZipFile())[1]
    assert "占用" in friendly_error(PermissionError())[1]
    assert "缺少必要" in friendly_error(KeyError("word/document.xml"))[1]
    assert "输出路径" in friendly_error(ValueError("拒绝原地修改：输出路径与输入相同"))[1]
    name, msg = friendly_error(RuntimeError("没见过这种"))
    assert name == "RuntimeError" and "没见过这种" in msg


# ---- 打包相关的小工具 ----
def test_默认输出名与原文件不同():
    assert suggest_output_path(r"C:\论文\我的论文.docx") != r"C:\论文\我的论文.docx"
    assert suggest_output_path(r"C:\论文\我的论文.docx").endswith("_已排版.docx")


def test_能找到内置规则():
    p = default_rules_path()
    assert p is not None and p.exists(), "内置规则没找到，打包后会起不来"
    assert p.suffix == ".yaml"


def test_输出目录不存在时给出明确的提示(app, rules, tmp_path):
    """别让人以为是"文件被移动了"——那是输入侧的说法。"""
    src = build_docx(tmp_path / "a.docx",
                     paragraphs=[("图2.1油茶果实", None, "center", "宋体", None, 10.5)])
    dst = tmp_path / "没有这个目录" / "b.docx"
    out = run_job(app, JobThread("format", src, rules, output_path=dst))
    name, msg = out["failed"]
    assert name == "OutputPathError"
    assert "输出目录不存在" in msg


def test_输出路径与原文件相同报的是路径错(app, rules, tmp_path):
    src = build_docx(tmp_path / "a.docx",
                     paragraphs=[("图2.1油茶果实", None, "center", "宋体", None, 10.5)])
    out = run_job(app, JobThread("format", src, rules, output_path=src))
    assert out["failed"][0] == "OutputPathError"
    assert "不能和原文件是同一个" in out["failed"][1]
