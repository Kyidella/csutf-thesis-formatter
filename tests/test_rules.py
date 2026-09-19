"""规则文件自身的测试。

管的是"规则里写了、代码里没读"这件事——它是**静默**的：配置摆在那儿，
看规则的人以为在生效，跑起来却什么都不发生。脚注编号就是这么发现的
（检查器报 warn，格式化器里根本没有写 footnotePr 的代码）。

做法不是去追每个键，而是立一条硬规矩：**每个叶子键，要么代码里读得到，
要么登记在 not_implemented 里**。新加一条规则忘了接线，测试就红。

局限（写清楚，免得误信）：判"读得到"的办法是看键名有没有作为字符串/属性
出现在核心模块里。像 `align`、`size_pt`、`font_cn` 这种通用名字到处都是，
必然"读得到"，所以**通用名字的漏接线这条测试抓不到**。
它专抓 `new_page`、`gap_after_number`、`number_gap` 这类有特征的名字。
"""

import fnmatch
import pathlib
import re

import pytest
import yaml

from conftest import RULES

pytestmark = pytest.mark.skipif(not RULES.exists(), reason="规则文件缺失")

ROOT = pathlib.Path(__file__).resolve().parent.parent
CORE = ["docmodel.py", "classifier.py", "check_docx.py", "formatter.py",
        "textfix.py", "wordstyles.py"]

# 只用来写文档、不承载行为的键
DOC_ONLY = {
    "desc", "note", "why", "issue", "decision", "observed", "examples",
    "policy", "exempt", "low_confidence", "source", "size", "auto_fix",
    "name",                     # characters/synonyms 里给报错用的名字
    "label",                    # "关键词：" 这类标签文本
    "template_instance", "sample_instance",     # fallback 段里的样本记录
}
# 整段都是记录的（规范出处、回退依据），不参与接线检查
DOC_SECTIONS = ("meta", "fallback")


@pytest.fixture(scope="module")
def rules():
    return yaml.safe_load(open(RULES, encoding="utf-8"))


@pytest.fixture(scope="module")
def code():
    return "\n".join((ROOT / f).read_text(encoding="utf-8") for f in CORE)


def leaf_paths(node, path=""):
    """展开所有叶子键的完整路径。列表只取第一个元素当样本（规则里列表结构一致）。"""
    out = []
    if isinstance(node, dict):
        for k, v in node.items():
            out += leaf_paths(v, f"{path}.{k}" if path else str(k))
    elif isinstance(node, list) and node:
        out += leaf_paths(node[0], path)
    else:
        out.append(path)
    return out


def referenced(path, code):
    """键名有没有作为字符串/属性名出现在核心模块里。"""
    name = path.split(".")[-1]
    return bool(re.search(rf"[\"']{re.escape(name)}[\"']|\.{re.escape(name)}\b",
                          code))


def test_规则里每个叶子键要么被读要么登记在案(rules, code):
    declared = [e["key"] for e in rules["not_implemented"]["keys"]]

    def covered(path):
        # 条目可以带通配符，按"一类规则"记：styles.*.new_page
        return any(fnmatch.fnmatchcase(path, d) or path.startswith(d + ".")
                   for d in declared)

    unlisted = [p for p in leaf_paths(rules)
                if not p.startswith(DOC_SECTIONS)
                and p.split(".")[-1] not in DOC_ONLY
                and not referenced(p, code)
                and not covered(p)]

    assert not unlisted, (
        "这些规则键代码里读不到，也没登记进 not_implemented：\n  "
        + "\n  ".join(sorted(unlisted))
        + "\n要么接上代码，要么登记，别让它当摆设。")


def test_每条未实现都写了理由(rules):
    for e in rules["not_implemented"]["keys"]:
        assert e.get("key"), f"缺 key：{e}"
        assert e.get("desc"), f"{e['key']} 缺 desc（规范原文）"
        assert e.get("why"), f"{e['key']} 缺 why（为什么还没做）"


# 注：不写"反过来查列表有没有腐烂"的测试。原因是这个检查按**键名**判"读得到"，
# 而 font_en / size_pt 这类通用名字在代码里到处都是，反过来查必然误报
# （pagination.size_pt、reference_item.gap_after_number 都会被冤枉）。
# 所以只有正向这一条硬约束：新加规则忘了接线，测试会红；
# 至于"实现了却忘了从列表里删"，只能靠 YAML 里那句注释和自觉。


def test_规则文件仍是合法YAML且关键段齐全(rules):
    for key in ("page", "styles", "word_styles", "characters", "punctuation",
                "required_sections", "word_count", "not_implemented"):
        assert key in rules, f"缺少顶层段 {key}"
    assert rules["styles"]["body"]["size_pt"] == 12.0
