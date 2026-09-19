"""分类器的单元测试。

重点验证"位置感知"这一层——纯文本特征分不出摘要正文和正文，
必须靠章节上下文。
"""

from classifier import classify, heading_level
from conftest import make_doc, make_para


def keys(paras):
    return [r.key for r in classify(make_doc(paras))]


def test_摘要之后的普通段落归为摘要正文():
    """这是分组器存在的理由：摘要正文和正文长得一模一样。"""
    ks = keys([
        make_para(0, "摘  要"),
        make_para(1, "本文研究了油茶果实的滞育机制，发现在低温条件下……"),
        make_para(2, "ABSTRACT"),
        make_para(3, "This study investigates the diapause of Camellia oleifera."),
    ])
    assert ks[0] == "abstract_title"
    assert ks[1] == "abstract_body", "摘要下面的普通文字必须归摘要正文，不是正文"
    assert ks[2] == "abstract_en_title"
    assert ks[3] == "abstract_en_body"


def test_参考文献之后的条目归为参考文献():
    ks = keys([
        make_para(0, "参考文献"),
        make_para(1, "[1] 曹凌．中国佛教疑伪经综录[M]．上海：上海古籍出版社，2011：19."),
        make_para(2, "陈永忠. 我国油茶科技进展[J]. 中南林业科技大学学报, 2023."),
        make_para(3, "致  谢"),
        make_para(4, "感谢导师的悉心指导，感谢同门在实验中的帮助。"),
    ])
    assert ks[0] == "reference_title"
    assert ks[1] == "reference_item"
    assert ks[2] == "reference_item", "无编号的条目也应落在参考文献区"
    assert ks[3] == "ack_title"
    assert ks[4] == "ack_body"


def test_表格内段落不判标题():
    """单元格里的 "1.5"、"1.2.3" 会命中编号模式，误判成标题。"""
    ks = keys([
        make_para(0, "1.2 标题", in_table=True),
        make_para(1, "1.5", in_table=True),
        make_para(2, "2.5", in_table=True),
    ])
    assert ks == ["table_content"] * 3


def test_封面区不判标题():
    """封面段落常被手动设大纲级别，不挡住就会误判成一级标题。"""
    ks = keys([
        make_para(0, "油茶果实前期滞育研究", outline_level=0),
        make_para(1, "by", outline_level=0),
        make_para(2, "Associate Professor LI Ning", outline_level=0),
        make_para(3, "摘  要"),
        make_para(4, "正文内容。"),
    ])
    assert ks[:3] == ["cover"] * 3
    assert ks[3] == "abstract_title"
    assert ks[4] == "abstract_body"


def test_日期不是标题():
    assert heading_level(make_para(0, "2025年6月"))[0] is None
    assert heading_level(make_para(0, "2025.6"))[0] is None


def test_列表项不是标题():
    assert heading_level(make_para(0, "1）参加项目课题"))[0] is None
    assert heading_level(make_para(0, "2、发表论文"))[0] is None


def test_长条文不是标题():
    long_text = "1.5 (μL):1 (μL)比例混匀并预热。按顺序加样，注意避免气泡产生。"
    assert heading_level(make_para(0, long_text))[0] is None


def test_封面地址不误判为一级标题():
    """'498 Shaoshan South Road' 以数字开头，但绝不该是章标题。"""
    assert heading_level(make_para(0, "498 Shaoshan South Road，Tianxin District"))[0] is None
    assert heading_level(make_para(0, "1  绪论"))[0] == 1


def test_章标题各种编号形式():
    cases = {
        "第一章 绪论": 1,
        "1  绪论": 1,
        "1.1 研究背景": 2,
        "1.2.3 细胞壁": 3,
        "2.4.1.1 ROS含量": 4,
    }
    for text, want in cases.items():
        got, why = heading_level(make_para(0, text))
        assert got == want, f"{text!r} 应为 {want} 级，实际 {got}（{why}）"


def test_大纲级别优先于编号模式():
    got, why = heading_level(make_para(0, "随便一句话", outline_level=2))
    assert got == 3 and why == "大纲级别"


def test_目录区终止于章标题():
    """目录里的 '附录A …… 100' 不能当章节锚点，否则会提前打断目录区。"""
    ks = keys([
        make_para(0, "目  录"),
        make_para(1, "\t\t摘  要\tI", style_name="toc 1"),
        make_para(2, "\t\t1  绪论\t1", style_name="toc 1"),
        make_para(3, "第一章 绪论", outline_level=0),
        make_para(4, "正文开始。"),
    ])
    assert ks[0] == "toc_title"
    assert ks[1].startswith("toc_entry")
    assert ks[2].startswith("toc_entry")
    assert ks[3] == "heading1", "章标题出现即目录区结束"
    assert ks[4] == "body"


def test_二级标题不重置章节上下文():
    """只有一级标题代表新章节；小节标题不该把摘要正文的上下文清掉。"""
    ks = keys([
        make_para(0, "摘  要"),
        make_para(1, "摘要第一段。"),
        make_para(2, "1.1 小节", outline_level=1),
        make_para(3, "摘要里的小节之后。"),
    ])
    assert ks[3] == "abstract_body"


def test_图题表题跨区段仍能识别():
    ks = keys([
        make_para(0, "摘  要"),
        make_para(1, "图1.1 技术路线"),
        make_para(2, "表2.1 洗脱梯度"),
        make_para(3, "续表2.1 表名"),
    ])
    assert ks[1] == "figure_caption"
    assert ks[2] == "table_caption"
    assert ks[3] == "table_continued"
