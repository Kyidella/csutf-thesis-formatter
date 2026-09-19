"""端到端回归测试：拿真实论文跑，锁住已知结论。

样本是学生的真实论文（含姓名等），不适合入版本库，所以按路径引用、
文件不在就跳过。这里断言的是"人工确认过的真问题"，不是"当前输出的快照"——
以后改代码若让这些结论变了，就是回归。
"""

import pytest
import yaml

from check_docx import run_check
from classifier import classify
from conftest import RULES, SAMPLES
from docmodel import load

pytestmark = pytest.mark.skipif(not RULES.exists(), reason="规则文件缺失")


@pytest.fixture(scope="module")
def rules():
    return yaml.safe_load(open(RULES, encoding="utf-8"))


def _sample(name):
    """取样本路径；文件不在就跳过。

    真实论文含姓名与未发表内容，不入版本库（见 conftest 的说明）。
    """
    p = SAMPLES[name]
    if not p.exists():
        pytest.skip(f"样本不在本地，跳过：{p}")
    return p


def _load(name):
    return load(_sample(name))


def hit(rep, category, needle):
    """某个类别里是否有包含关键字的条目。"""
    return any(i["category"] == category and needle in i["msg"]
               for i in rep.items)


# --------------------------------------------------------------------------
# 模型层：这两组数字是"漏 57% 段落"那个 bug 的锚点
# --------------------------------------------------------------------------
def test_终版段落覆盖完整():
    d = _load("终版")
    assert d.stats["paragraphs_total"] == 1623
    assert d.stats["paragraphs_in_table"] == 815
    assert d.stats["paragraphs_in_sdt"] == 115
    # 三部分相加应等于总数（不能有段落既不算正文流也不算表格/控件）
    assert (d.stats["paragraphs_body"] + d.stats["paragraphs_in_table"]
            + d.stats["paragraphs_in_sdt"]) == d.stats["paragraphs_total"]


def test_终版节口径一致():
    d = _load("终版")
    assert d.stats["sections"] == len(d.sections) == 14


# --------------------------------------------------------------------------
# 分类器：位置感知的效果
# --------------------------------------------------------------------------
@pytest.mark.parametrize("name,expect", [("终版", 6), ("送审版", 6)])
def test_章标题数量(name, expect):
    """第一~六章。多一个少一个都说明标题判定出了问题。"""
    d = _load(name)
    rows = classify(d)
    h1 = [r for r in rows if r.key == "heading1"]
    assert len(h1) == expect
    assert all("第" in r.para.text and "章" in r.para.text for r in h1)


def test_摘要正文被正确归类():
    """纯文本特征分不出摘要正文；这是位置感知分类器的核心价值。"""
    d = _load("终版")
    rows = classify(d)
    counts = {}
    for r in rows:
        counts[r.key] = counts.get(r.key, 0) + 1
    assert counts["abstract_body"] > 0, "摘要正文不能为空"
    assert counts["ack_body"] > 0, "致谢正文不能为空"
    assert counts["table_content"] > 800, "表格段落应归 table_content"


# --------------------------------------------------------------------------
# 检查器：人工确认过的真问题
# --------------------------------------------------------------------------
def test_终版页边距违规(rules):
    """13/14 个节用了 Word 出厂默认 2.54/3.17，规范要求 2.5/3.0。"""
    rep, _, _ = run_check(_sample("终版"), rules)
    assert hit(rep, "页边距", "13/14")
    assert hit(rep, "页边距", "2.54")


def test_终版二级标题西文字体错(rules):
    """用户肉眼发现的问题：二级标题的数字显示成宋体。"""
    rep, _, _ = run_check(_sample("终版"), rules)
    assert hit(rep, "西文字体", "二级标题")
    assert hit(rep, "西文字体", "宋体×27")


def test_终版参考文献标题未居中(rules):
    """用户肉眼发现的第二个问题。"""
    rep, _, _ = run_check(_sample("终版"), rules)
    assert hit(rep, "对齐", "参考文献标题")
    assert hit(rep, "对齐", "center")


def test_终版中英题注错位(rules):
    """表4.1 的中文与英文题注互相错位。"""
    rep, _, _ = run_check(_sample("终版"), rules)
    assert hit(rep, "中英题注不符", "表4.1")
    assert hit(rep, "中英题注不符", "表4.2")


def test_两版错误数稳定(rules):
    """错误数是最稳的指标——两版的错误都是同样那四类。"""
    for name in ("终版", "送审版"):
        rep, _, _ = run_check(_sample(name), rules)
        assert rep.count("error") == 4, f"{name} 错误数变了：{rep.count('error')}"


def test_报告优先区能筛出标题类问题(rules):
    """用户反馈'重点问题埋在 70 条警告里'，优先区必须把它们捞出来。"""
    rep, doc, rows = run_check(_sample("终版"), rules)
    pri = rep.priority()
    assert len(pri) < len(rep.items) / 2, "优先区应显著短于完整清单"
    assert all(i["level"] == "error" or i.get("ok") == 0
               or i.get("key") in
               {"heading1", "heading2", "heading3", "heading4",
                "abstract_title", "abstract_en_title", "reference_title",
                "ack_title", "appendix_title", "toc_title",
                "abstract_keywords", "abstract_en_keywords"}
               for i in pri)
    # 参考文献标题未居中是标题类问题，必须在优先区里
    assert any("参考文献标题" in i["msg"] for i in pri)


def test_报告优先区也捞非标题类的整体性错误(rules):
    """筛选规则有三条：错误 / 整体不符规范(0 段符合) / 章节标题类。
    这条专门盯住中间那条——只测标题类的话，"整体不符"这一支被删掉也测不出来。
    """
    rep, doc, rows = run_check(_sample("终版"), rules)
    title_keys = {"heading1", "heading2", "heading3", "heading4",
                  "abstract_title", "abstract_en_title", "reference_title",
                  "ack_title", "appendix_title", "toc_title",
                  "abstract_keywords", "abstract_en_keywords"}
    from check_docx import TITLE_KEYS
    non_title_total = [i for i in rep.items
                       if i["level"] == "warn" and i.get("ok") == 0
                       and i.get("key") not in TITLE_KEYS]
    assert non_title_total, "样本里应存在非标题类的整体性不符，否则这条测试失去意义"
    pri_ids = {id(i) for i in rep.priority()}
    assert any(id(i) in pri_ids for i in non_title_total), \
        "'整体不符规范'的条目必须进优先区"
