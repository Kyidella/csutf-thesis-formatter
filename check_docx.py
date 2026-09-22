"""按 YAML 规则检查一篇 .docx，输出 Markdown 报告。

解析全部交给 docmodel（一次解析、统一口径），段落类型判定交给 classifier
（位置感知）。本模块只负责「比对规则」和「写报告」。

用法:
    python check_docx.py input.docx --rules rules/xxx.yaml --report report.md
"""

import argparse
import re
import sys
import zipfile
from collections import Counter

import yaml
from lxml import etree

from classifier import NUM_PATTERNS, classify
from docmodel import has_drawing, load, stage
from textfix import caption_gap_pos, punct_fixes


# --------------------------------------------------------------------------
# 异常 → 人能看懂的中文
# --------------------------------------------------------------------------
# 界面和两个命令行入口共用这一份。放在这里而不是 gui_worker，是因为
# gui_worker 依赖 Qt（QThread），核心层不能反过来引它——而命令行也要说人话。
# 翻译逻辑只有一份，就不会出现"界面说得好好的、命令行甩 traceback"。
#
# 顺序重要：先匹配子类/更具体的。KeyError 放在最后面之前，
# 因为 zipfile 读不到部件时抛的就是 KeyError。
ERROR_MESSAGES = [
    (zipfile.BadZipFile, "这不是有效的 .docx 文件（可能是 .doc，或者文件已损坏）"),
    (FileNotFoundError, "找不到文件，可能已被移动或删除"),
    (PermissionError, "文件被 WPS / Word 占用，请先关掉再试"),
    (etree.XMLSyntaxError, "文档内部的 XML 损坏，解析不了"),
    (yaml.YAMLError, "规则文件格式有误"),
    (KeyError, "文档里缺少必要的部件（如 word/document.xml），可能不是 Word 生成的文档"),
    (ValueError, "参数不对——最常见的是输出路径和原文件相同"),
    (TypeError, "规则文件内容类型不符（比如该是数字的地方写了文字）"),
    (MemoryError, "文件太大，内存不够"),
]


class OutputPathError(Exception):
    """输出路径本身有问题（目录不存在、没权限之类）。

    单独一类是因为默认的 ValueError 提示语是"输出路径和原文件相同"，
    套在别的路径问题上会把人带偏。
    """


def friendly_error(exc):
    """把异常翻成中文。返回 (异常类名, 提示语)。"""
    if isinstance(exc, OutputPathError):
        return type(exc).__name__, str(exc)
    for cls, msg in ERROR_MESSAGES:
        if isinstance(exc, cls):
            return type(exc).__name__, msg
    return type(exc).__name__, f"出错了：{exc}"


def known_error(exc):
    """这个异常认得吗？

    命令行靠它决定要不要把 traceback 亮出来：认得的翻成中文，不认得的原样抛
    ——那才是真出了 bug，把 traceback 藏掉反而没法排查。
    """
    return isinstance(exc, OutputPathError) or any(
        isinstance(exc, cls) for cls, _ in ERROR_MESSAGES)


def cli_guard(fn):
    """命令行入口的收尾。用法：`cli_guard(main)`。"""
    try:
        return fn()
    except Exception as exc:                       # noqa: BLE001
        if not known_error(exc):
            raise
        sys.exit(f"错误：{friendly_error(exc)[1]}")


# --------------------------------------------------------------------------
# 报告
# --------------------------------------------------------------------------


class Report:
    def __init__(self):
        self.items = []

    def add(self, level, category, msg, loc="", key=None, ok=None, total=None):
        self.items.append({"level": level, "category": category, "msg": msg,
                           "loc": loc, "key": key, "ok": ok, "total": total})

    def count(self, level):
        return sum(1 for i in self.items if i["level"] == level)

    def priority(self):
        """挑出最值得先看的条目。

        判断依据不是"规则多重要"，而是"人眼看不看得见"：
          - 错误：一律置顶
          - 整体不符规范（0 段符合）：属于系统性错误
          - 章节标题类：在页面上最显眼
        其余（如某类型里 2 段字号异常）留在完整清单里，不往上顶。
        """
        out = []
        for i in self.items:
            if i["level"] == "error":
                out.append(i)
            elif i["level"] == "warn" and i.get("ok") == 0:
                out.append(i)
            elif i["level"] == "warn" and i.get("key") in TITLE_KEYS:
                out.append(i)
        return out


# 章节标题类：在页面上最显眼，出问题优先报
TITLE_KEYS = {"heading1", "heading2", "heading3", "heading4",
              "abstract_title", "abstract_en_title", "abstract_keywords",
              "abstract_en_keywords", "reference_title", "ack_title",
              "appendix_title", "toc_title"}


# --------------------------------------------------------------------------
# 哪些问题格式化器能自己修、修到哪一层
# --------------------------------------------------------------------------
# 写成显式映射，而不是让人从报告里猜。报告和界面共用这一份。
# 没登记的按"需人工修改"处理——宁可说保守，也别让人以为会自动修。
#
# 这个映射是被一次审查逼出来的：脚注编号检查器报 warn，格式化器里却根本没有
# 写 footnotePr 的代码（"脚注" 这个词只在 docmodel 的读取侧出现）。报告里看不出来，
# 看的人会以为它会自己修好。
FIX_KIND = {
    # 改 w:sectPr
    "页边距": "文档级",
    "页面尺寸": "文档级",
    "页码": "文档级",
    # 改 word/styles.xml 与段落的 pStyle
    "字号": "样式级",
    "中文字体": "样式级",
    "西文字体": "样式级",
    "行距": "样式级",
    "段前": "样式级",
    "段后": "样式级",
    "首行缩进": "样式级",
    "左缩进": "样式级",
    "对齐": "样式级",
    # 改 w:t 里的字
    "图题格式": "文本级",
    "表题格式": "文本级",
    "标点": "文本级",
    "异常字符": "文本级",
    "字符混用": "文本级",
    # 内容问题，改了等于替作者做决定，一律不自动修
    "脚注": None,
    "标题与正文同段": None,
    "缺失章节": None,
    "中英题注不符": None,
    "编号顺序": None,
    "关键词": None,
    "字数": None,
}


FIXABLE_LABEL = "支持自动排版"
MANUAL_LABEL = "建议手动修改"

# 规则键级别的例外：分类上"能自动修"，落在这些章节上也不行。
# 目录条目由 TOC 域生成，改 pStyle 会在刷新目录时被 Word 覆盖，
# 所以格式化器只报不改（见 formatter 里跳过它们的说明）。
#
# 为什么不能只按分类判：分类的粒度是「字号」「对齐」，同是「字号」，
# 正文标题修得了、目录条目修不了。只看分类会标成"支持自动排版"，
# 用户点了自动排版却发现目录纹丝不动——报告等于在承诺工具不会做的事。
# 格式化器直接引用这个集合，两边不可能各说各话。
SKIP_FIX_KEYS = {"toc_entry_front", "toc_entry_1", "toc_entry_2", "toc_entry_3"}


def fix_kind(category, key=None):
    """内部层次（文档级/样式级/文本级），给悬停提示用；不自动修返回 None。

    key 是该条问题对应的规则键（如 toc_entry_2）。同一个分类落在不同章节上，
    能不能自动修未必一样，所以要看键，不能只看分类。
    """
    if key in SKIP_FIX_KEYS:
        return None
    return FIX_KIND.get(category)


def fix_hint(category, key=None):
    """报告里那条问题后面跟的标记。

    给用户看的只有两种说法：工具会不会自己处理。
    FIX_KIND 里那个"文档级/样式级/文本级"是内部实现分层，对用户没意义，
    所以只放进界面的悬停提示（见 fix_kind），不印在报告上。
    """
    return f"　`{FIXABLE_LABEL}`" if fix_kind(category, key) else f"　`{MANUAL_LABEL}`"


def para_loc(p, n=28):
    """问题的位置：先给段落开头几个字，段号缩进括号里。

    段号是给开发者定位用的，对用户毫无意义；段落开头那几个字才是可操作的
    ——在 WPS 里 Ctrl+F 一搜就到。段号留在后面，我调试时还看得到。
    """
    t = " ".join(p.text.split())
    if not t:
        return f"段{p.index}"
    return f"{t[:n]}{'…' if len(t) > n else ''}（段{p.index}）"


def _show(kind, val):
    return {"lines": f"{val}行", "pt": f"{val}磅"}.get(kind, str(val))


def check_dist(rep, tag, desc, dist, want, fmt, loc, level="warn", key=None):
    """按『分布』而非单个代表值比对：符合多少段、不符多少段。"""
    if not dist:
        return
    bad = {v: n for v, n in dist.items() if v != want}
    if not bad:
        return
    ok = sum(n for v, n in dist.items() if v == want)
    detail = "，".join(f"{fmt(v)}×{n}"
                      for v, n in sorted(bad.items(), key=lambda x: -x[1])[:4])
    rep.add(level, tag, f"{desc}：规范要求 {fmt(want)}；符合 {ok} 段，"
                        f"不符 {sum(bad.values())} 段（{detail}）", loc,
            key=key, ok=ok, total=sum(dist.values()))


def expect_line_spacing(rule):
    t = rule.get("type")
    if t == "single":
        return ("multiple", 1.0)
    if t == "multiple":
        return ("multiple", float(rule["value"]))
    return (t, float(rule.get("value", 0)))


# --------------------------------------------------------------------------
# 页面与节
# --------------------------------------------------------------------------
def check_page(doc, rules, rep):
    want = rules["page"]["margins_cm"]
    bad_sections, ok_count = [], 0
    for sec in doc.sections:
        got = sec.margins_cm
        if any(abs((got.get(k) or -1) - want[k]) > 0.01 for k in want):
            bad_sections.append((sec.index, got))
        else:
            ok_count += 1
    if bad_sections:
        i, got = bad_sections[0]
        rep.add("error", "页边距",
                f"{len(bad_sections)}/{len(doc.sections)} 个节的页边距不符"
                f"（合规 {ok_count} 个）。首个偏差：节 {i} 上{got['top']} 下{got['bottom']} "
                f"左{got['left']} 右{got['right']}；规范要求 上{want['top']} 下{want['bottom']} "
                f"左{want['left']} 右{want['right']}",
                loc=f"节 {[b[0] for b in bad_sections]}")

    w, h = rules["page"]["width_mm"], rules["page"]["height_mm"]
    off = [s.index for s in doc.sections
           if s.page_w_cm and abs(s.page_w_cm * 10 - w) > 1]
    if off:
        rep.add("error", "页面尺寸", f"节 {off} 非 A4（{w}×{h}mm）", loc=str(off))


NUMFMT_CN = {
    "decimal": "阿拉伯数字", "upperRoman": "大写罗马数字",
    "lowerRoman": "小写罗马数字", "decimalEnclosedCircle": "①②③",
}


def check_page_numbering(doc, rules, rep):
    groups = {}
    for sec in doc.sections:
        groups.setdefault((sec.pg_num_fmt, sec.pg_num_start), []).append(sec.index)
    for (fmt, start), secs in sorted(groups.items(), key=lambda x: min(x[1])):
        rep.add("info", "页码",
                f"节 {sorted(secs)}：{NUMFMT_CN.get(fmt, fmt or '默认')}"
                + (f"，起始 {start}" if start else ""))
    kinds = {f for f, _ in groups}
    if "upperRoman" not in kinds:
        rep.add("warn", "页码", "未检测到大写罗马数字页码"
                "（规范：摘要到目录用 I,II,III）")

    fn = rules["styles"]["footnote"].get("marker_style")
    any_sec = doc.sections[0] if doc.sections else None
    fmt = next((s.footnote_fmt for s in doc.sections if s.footnote_fmt), None)
    rst = next((s.footnote_restart for s in doc.sections if s.footnote_restart), None)
    if fn == "circled":
        if fmt != "decimalEnclosedCircle":
            rep.add("warn", "脚注", f"规范要求脚注用 ①②③ 编号，实际 numFmt={fmt or '未设置'}")
        if rst != "eachPage":
            rep.add("warn", "脚注", f"规范要求脚注每页重新编号，实际 numRestart={rst or '未设置'}")


# --------------------------------------------------------------------------
# 结构
# --------------------------------------------------------------------------
def norm(s):
    return re.sub(r"\s+", "", s)


def missing_sections(doc, rules):
    """没检测到的必备章节名。检查器和排版器共用这一份判据。

    排版器也要它：在缺摘要、缺参考文献的稿子上排版，改动清单看起来一切正常，
    用户会以为文档没问题。得明说"缺这些"，而且缺得越多、定位越不准。
    """
    found = set()
    for t in doc.all_text(include_tables=True).split("\n"):
        t = norm(t)
        if not t or len(t) > 40:
            continue
        for s in rules["required_sections"]:
            if norm(s["label"]) in t:
                found.add(s["label"])
    return [s["label"] for s in rules["required_sections"]
            if s["key"] != "body_chapters" and s["label"] not in found]


def check_sections(doc, rules, rep):
    for label in missing_sections(doc, rules):
        rep.add("warn", "缺失章节", f"未检测到“{label}”", loc=label)


PAT_FIG = re.compile(r"^\s*图\s*(\d+)[.\-–](\d+)")
PAT_TAB = re.compile(r"^\s*表\s*(\d+)[.\-–](\d+)")
PAT_FIG_EN = re.compile(r"^\s*Fig(?:ure)?\.?\s*(\d+)[.\-–](\d+)", re.I)
PAT_TAB_EN = re.compile(r"^\s*Tab(?:le)?\.?\s*(\d+)[.\-–](\d+)", re.I)


NEWLINE = chr(10)      # 软回车（w:br）在段落文本里就是一个换行符


def _looks_like_numbered_heading(line):
    """这一行像不像带编号的标题（如 "3.2.3 xxx"）。

    复用分类器的编号模式，不另写一套——两处判据分叉过一次就够了。
    """
    return any(pat.match(line) for pat, lvl in NUM_PATTERNS if lvl >= 2)


def check_merged_headings(rows, rep):
    """标题和正文挤在同一个段落里（中间是软回车，不是回车）。

    这种段落整段超过标题长度上限（classifier.HEADING_MAX_LEN），分类器只能
    当正文处理——于是"标题的字体字号套不上"，用户会以为工具漏了。
    工具不该替作者拆段落（那是改结构），但必须说出来，否则就是默默不干活。
    """
    bad = []
    for r in rows:
        p = r.para
        if NEWLINE not in p.text or r.key.startswith("heading"):
            continue
        first = p.text.split(NEWLINE, 1)[0].strip()
        if len(first) > 45 or not _looks_like_numbered_heading(first):
            continue
        bad.append((p.index, first))
    if not bad:
        return
    where = "、".join("段" + str(i) for i, _ in bad[:4])
    rep.add("warn", "标题与正文同段",
            str(len(bad)) + " 处的第 1 行是标题，正文紧跟在同一个段落里（中间是软回车）。"
            "这种段落整段只能按正文排，标题的字体字号套不上——"
            "在 WPS 里把光标放到标题末尾按回车就能分开。"
            "例：" + repr(bad[0][1][:26]),
            loc=where + (" 等" if len(bad) > 4 else ""))


def check_numbering(rows, rep):
    figs, tabs = [], []
    seq = [r.para for r in rows]

    def bilingual(m, en_pat, zh_label, en_label, para):
        """中/英文题注必须同号。英文题注若不存在则不判。"""
        idx = para.index
        if idx + 1 >= len(seq):
            return
        e = en_pat.match(seq[idx + 1].text.strip())
        if e is None:
            return
        zh = f"{m.group(1)}.{m.group(2)}"
        en = f"{e.group(1)}.{e.group(2)}"
        if zh != en:
            rep.add("error", "中英题注不符",
                    f"中文“{zh_label}{zh}”对应英文“{en_label}{en}”",
                    loc=para_loc(para))

    for r in rows:
        p = r.para
        if r.key not in ("figure_caption", "table_caption"):
            continue
        t = p.text.strip()
        is_fig = r.key == "figure_caption"
        m = (PAT_FIG if is_fig else PAT_TAB).match(t)
        if not m:
            continue
        num = f"{m.group(1)}.{m.group(2)}"
        (figs if is_fig else tabs).append((num, p.index, t))
        label = "图" if is_fig else "表"
        if caption_gap_pos(t, label) is not None:
            rep.add("warn", f"{label}题格式",
                    f"{label}号后缺空格（规范：{label}序与{label}名之间空 1 个空格）",
                    loc=para_loc(p))
        bilingual(m, PAT_FIG_EN if is_fig else PAT_TAB_EN,
                  label, "Fig " if is_fig else "Table ", p)

    caption_text = {idx: t for _, idx, t in figs + tabs}

    def text_of(idx, n=28):
        t = caption_text.get(idx, "")
        return t[:n] + ("…" if len(t) > n else "")

    for name, lst in (("图", figs), ("表", tabs)):
        last = {}
        for num, idx, _ in lst:
            ch, sub = (int(x) for x in num.split("."))
            if ch in last and sub <= last[ch][0]:
                rep.add("error", "编号顺序",
                        f"{name}{num} 出现在 {name}{last[ch][1]} 之后，编号未递增",
                        loc=f"{text_of(idx)}（段{idx}）")
            last[ch] = (sub, num)


# --------------------------------------------------------------------------
# 样式
# --------------------------------------------------------------------------
def metrics_of(p):
    return {
        "line_spacing": p.pformat.get("line_spacing"),
        "space_before": p.pformat.get("space_before"),
        "space_after": p.pformat.get("space_after"),
        "first_line_chars": p.size_chars(None),
        "left_chars": p.left_chars(),
        "align": p.align or "left",
    }


def check_fmt_match(rule, dist, rep, desc, loc, key=None):
    if "line_spacing" in rule and dist.get("line_spacing"):
        check_dist(rep, "行距", desc, dist["line_spacing"],
                   expect_line_spacing(rule["line_spacing"]),
                   lambda v: f"{v[0]} {v[1]}", loc, key=key)

    for fld, tag in (("space_before", "段前"), ("space_after", "段后")):
        if fld not in rule or not dist.get(fld):
            continue
        w = rule[fld]
        want = ("lines", float(w["lines"])) if "lines" in w else ("pt", float(w["pt"]))
        got = dist[fld]
        if {v[0] for v in got} != {want[0]}:
            rep.add("info", tag, f"{desc}：规范要求 {_show(*want)}，实际以 "
                    f"{_show(*max(got, key=got.get))} 为主（单位不同，请人工确认）", loc)
        else:
            check_dist(rep, tag, desc, got, want, lambda v: _show(*v), loc, key=key)

    if "indent" in rule:
        w = rule["indent"]
        if w.get("first_line_chars") is not None:
            check_dist(rep, "首行缩进", desc, dist.get("first_line_chars", {}),
                       float(w["first_line_chars"]), lambda v: f"{v}字符", loc, key=key)
        if w.get("left_chars") is not None:
            check_dist(rep, "左缩进", desc, dist.get("left_chars", {}),
                       float(w["left_chars"]), lambda v: f"{v}字符", loc, key=key)

    if rule.get("align"):
        check_dist(rep, "对齐", desc, dist.get("align", {}), rule["align"],
                   lambda v: str(v), loc, key=key)


SKIP_KEYS = {"cover", "table_content", "toc_title"}


def check_styles(doc, rows, rules, rep):
    want = rules["styles"]
    seen = {}
    for r in rows:
        key = r.key
        if key in SKIP_KEYS or key not in want or not key.startswith(
                ("heading", "body", "abstract", "ack", "reference", "appendix",
                 "figure", "table", "toc_entry", "abstract_en")):
            continue
        p = r.para
        if not p.text.strip():
            continue          # 空段是空行/占位，不承载格式意图，不计入统计
        if has_drawing(p.element):
            # 图片段落排版时有意跳过（套正文样式会把嵌入型图片裁掉），
            # 那就不能用样式规范去要求它——报了也没法改，只是噪声。
            continue
        rec = seen.setdefault(key, {"n": 0, "sizes": {}, "fonts": {}, "fonts_en": {},
                                    "metrics": {}, "first": p.index,
                                    "first_para": p})
        rec["n"] += 1
        if p.size_pt:
            rec["sizes"][p.size_pt] = rec["sizes"].get(p.size_pt, 0) + 1
        if p.font_cn:
            rec["fonts"][p.font_cn] = rec["fonts"].get(p.font_cn, 0) + 1
        if p.font_ascii:
            rec["fonts_en"][p.font_ascii] = rec["fonts_en"].get(p.font_ascii, 0) + 1
        for k, v in metrics_of(p).items():
            if v is not None:
                rec["metrics"].setdefault(k, {}).setdefault(v, 0)
                rec["metrics"][k][v] += 1

    for key, info in sorted(seen.items()):
        rule = want[key]
        desc = rule.get("desc", key)
        loc = f"首个：{para_loc(info['first_para'])}"
        if rule.get("size_pt") and info["sizes"]:
            check_dist(rep, "字号", desc, info["sizes"], float(rule["size_pt"]),
                       lambda v: f"{v}pt", loc, key=key)
        if rule.get("font_cn") and info["fonts"]:
            check_dist(rep, "中文字体", desc, info["fonts"], rule["font_cn"],
                       lambda v: str(v), loc, key=key)
        # 西文字体单独查：中文正文里的阿拉伯数字、"1.2" 这类编号走的是西文字体，
        # 只查中文字体会漏掉「标题数字显示成宋体」这种很扎眼的问题。
        if rule.get("font_en") and info["fonts_en"]:
            check_dist(rep, "西文字体", desc, info["fonts_en"], rule["font_en"],
                       lambda v: str(v), loc, key=key)
        check_fmt_match(rule, info["metrics"], rep, desc, loc, key=key)
    return seen


# --------------------------------------------------------------------------
# 内容
# --------------------------------------------------------------------------
def check_keywords(doc, rules, rep):
    cfg = rules["styles"]["abstract_keywords"]
    for p in doc.body_paragraphs:
        t = p.text.strip()
        if not t.startswith("关键词"):
            continue
        body = t.split("：", 1)[1] if "：" in t else t
        if body.endswith(("；", ";")):
            rep.add("warn", "关键词", "最后一个关键词后不应有标点符号", loc=t[:40])
        n = len([x for x in re.split(r"[；;]", body) if x.strip()])
        lo, hi = cfg["count_range"]
        if not (lo <= n <= hi):
            rep.add("warn", "关键词", f"关键词个数 {n}，规范要求 {lo}～{hi} 个",
                    loc=t[:40])


def check_wordcount(doc, rules, rep):
    total = sum(len(p.text) for p in doc.body_paragraphs)
    lo, hi = rules["word_count"]["thesis_total"]
    tag = "（偏少）" if total < lo else "（偏多）" if total > hi else "（在范围内）"
    rep.add("info", "字数", f"正文约 {total} 字，规范要求 {lo}～{hi} 字{tag}")


def classify_punct(text):
    """半角标点分三类：高置信度应改 / 需人工 / 引用标注豁免。

    判据本体在 textfix.punct_fixes（格式化器改的就是同一份判据），
    这里只把下标翻译成人能看的片段，保证"报的"和"改的"永远是一回事。
    """
    hi, lo, ex = punct_fixes(text)

    def snip(i):
        return text[max(0, i - 12): i + 13]

    return [snip(g[0]) for g in hi], [snip(i) for i in lo], ex


def check_punctuation(doc, rules, rep):
    if not rules.get("punctuation", {}).get("enabled"):
        return
    hi, lo, ex = [], [], 0
    for p in doc.paragraphs:
        if not p.text.strip():
            continue
        a, b, c = classify_punct(p.text)
        hi += [(p, s) for s in a]
        lo += [(p, s) for s in b]
        ex += c
    if hi:
        rep.add("warn", "标点", f"{len(hi)} 处半角标点两侧都是中文，应为全角。"
                f"例：{hi[0][1]!r}", loc=f"{len(hi)} 处 · 首处：{para_loc(hi[0][0])}")
    if lo:
        rep.add("info", "标点", f"{len(lo)} 处需人工判断（括号内含拉丁字母，"
                f"或比例式冒号）。例：{lo[0][1]!r}", loc=f"{len(lo)} 处 · 首处：{para_loc(lo[0][0])}")
    if ex:
        rep.add("info", "标点", f"{ex} 处引用标注 [数字] 按规范保持半角，已豁免")


def _bad_char_name(cp, ranges):
    for item in ranges:
        lo, hi = item["range"]
        if lo <= cp <= hi:
            return item["name"], item.get("note", "")
    return None, None


def check_characters(doc, rules, rep):
    """查不换行空格、零宽字符、私有区字符，以及同义字符混用。

    用黑名单而不是白名单：白名单式的"报一切非常见字符"会淹在制表符、
    ℃、外国人名重音字母、罗马数字、勾选框里——那些都是正常用法。
    """
    cfg = rules.get("characters", {})
    if not cfg.get("enabled"):
        return

    ranges = cfg.get("bad", [])
    hits = {}
    for p in doc.paragraphs:
        for ch in p.text:
            name, note = _bad_char_name(ord(ch), ranges)
            if not name:
                continue
            rec = hits.setdefault(name, {"n": 0, "first": p.index,
                                         "note": note, "snippet": None,
                                         "para": p})
            rec["n"] += 1
            if rec["snippet"] is None:
                i = p.text.index(ch)
                rec["snippet"] = p.text[max(0, i - 18): i + 18]

    for name, rec in sorted(hits.items(), key=lambda x: -x[1]["n"]):
        rep.add("warn", "异常字符",
                f"{name} × {rec['n']} 个：{rec['note']}。例：{rec['snippet']!r}",
                loc=para_loc(rec["para"]))

    for syn in cfg.get("synonyms", []):
        found, where = {}, None
        for p in doc.paragraphs:
            for ch in p.text:
                if ch in syn["group"]:
                    found[ch] = found.get(ch, 0) + 1
                    if where is None:
                        i = p.text.index(ch)
                        where = (p.index, p.text[max(0, i - 18): i + 18])
        if len(found) > 1:
            detail = "，".join(f"U+{ord(c):04X} ×{n}"
                              for c, n in sorted(found.items()))
            rep.add("warn", "字符混用",
                    f"{syn['name']} 在同一篇里混用：{detail}；"
                    f"建议统一为 {syn['prefer']}"
                    + (f"。例：{where[1]!r}" if where else ""),
                    loc=f"段{where[0]}" if where else "")


def not_implemented_items(rules):
    """规则里写了、但没有对应代码的条目。报告和界面共用这一份。"""
    return (rules.get("not_implemented") or {}).get("keys") or []


def run_check(input_path, rules, report_path=None, rules_path=None, on_stage=None):
    """跑完整检查。返回 (Report, Doc, rows)，供 CLI 与测试共用。

    rules_path 只用来在报告里注明规则出处——之前的报告把整份规则字典
    打进了"规则："那一行，33KB 一行的垃圾，报告基本没法读。

    on_stage(说明, 已完成, 总数) 是给 GUI 的进度回调，不传就是纯命令行行为。
    load() 占了大头时间，所以它单独算一步，进度条不会一直卡在 0。
    """
    checks = [
        ("页边距与页面尺寸", lambda: check_page(doc, rules, rep)),
        ("页码与脚注", lambda: check_page_numbering(doc, rules, rep)),
        ("章节结构", lambda: check_sections(doc, rules, rep)),
        ("图表题注与编号", lambda: check_numbering(rows, rep)),
        ("标题与正文同段", lambda: check_merged_headings(rows, rep)),
        ("字体字号行距缩进", lambda: check_styles(doc, rows, rules, rep)),
        ("关键词", lambda: check_keywords(doc, rules, rep)),
        ("字数", lambda: check_wordcount(doc, rules, rep)),
        ("标点全角半角", lambda: check_punctuation(doc, rules, rep)),
        ("异常字符", lambda: check_characters(doc, rules, rep)),
    ]
    total = 2 + len(checks)

    stage(on_stage, "正在读取文档…", 0, total)
    doc = load(input_path)
    stage(on_stage, "正在分类段落…", 1, total)
    rows = classify(doc)

    rep = Report()
    for k, (name, fn) in enumerate(checks):
        stage(on_stage, f"正在检查：{name}", 2 + k, total)
        fn()
    stage(on_stage, "正在整理报告…", total, total)

    order = {"error": 0, "warn": 1, "info": 2}
    rep.items.sort(key=lambda i: (order.get(i["level"], 3), i["category"]))
    if report_path:
        write_report(rep, doc, rows, input_path, rules, rules_path, report_path)
    return rep, doc, rows


def write_report(rep, doc, rows, input_path, rules, rules_path, report_path):
    tag = {"error": "错误", "warn": "警告", "info": "提示"}
    lines = [
        "# 格式检查报告", "",
        f"- 源文件：`{input_path}`",
        f"- 规则：`{rules_path or '（未提供路径）'}`",
        f"- 错误 {rep.count('error')} · 警告 {rep.count('warn')} · 提示 {rep.count('info')}",
        "",
    ]

    pri = rep.priority()
    lines += [f"## 重点问题（{len(pri)} 条）", ""]
    lines.append("> 筛选标准：全部错误 + 整体不符规范的类型 + 章节标题类。"
                 "其余明细见下方完整清单。")
    lines.append("")
    if pri:
        for it in pri:
            lines.append(f"- **[{tag[it['level']]}] {it['category']}** — {it['msg']}"
                         + (f"  `{it['loc']}`" if it["loc"] else "")
                         + fix_hint(it["category"], it.get("key")))
    else:
        lines.append("- （无）")

    lines += ["", "## 完整清单", ""]
    for it in rep.items:
        lines.append(f"- **[{tag[it['level']]}] {it['category']}** — {it['msg']}"
                     + (f"  `{it['loc']}`" if it["loc"] else "")
                     + (fix_hint(it["category"], it.get("key"))
                        if it["level"] != "info" else ""))

    counts = Counter(r.key for r in rows)
    lines += ["", "## 文档概况", ""]
    lines.append(f"- 段落总数 {doc.stats['paragraphs_total']}"
                 f"（正文流 {doc.stats['paragraphs_body']} · "
                 f"表格内 {doc.stats['paragraphs_in_table']} · "
                 f"内容控件内 {doc.stats['paragraphs_in_sdt']}）")
    lines.append(f"- 节 {doc.stats['sections']} · 含引用域的段落 {doc.stats['paragraphs_with_fields']}")
    lines += ["", "## 段落分类统计", ""]
    for k, n in counts.most_common():
        lines.append(f"- {k}: {n}")

    nk = not_implemented_items(rules)
    if nk:
        lines += ["", f"## 规范要求、本工具尚未实现（{len(nk)} 项）", "",
                  "> 这些条目写在规则文件里，但没有对应的检查或修正代码。"
                  "单列出来，是为了不让人以为它们已经在查了。"]
        lines.append("")
        for e in nk:
            lines.append(f"- `{e['key']}` — {e['desc']}  \n  _{e['why']}_")

    with open(report_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")


# --------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("input")
    ap.add_argument("--rules", required=True)
    ap.add_argument("--report", default="report.md")
    args = ap.parse_args()

    rules = yaml.safe_load(open(args.rules, encoding="utf-8"))
    rep, doc, rows = run_check(args.input, rules, args.report,
                               rules_path=args.rules)
    tag = {"error": "错误", "warn": "警告", "info": "提示"}

    print(f"# 格式检查报告\n")
    print(f"- 错误 {rep.count('error')} · 警告 {rep.count('warn')} · 提示 {rep.count('info')}\n")
    print(f"## 重点问题（{len(rep.priority())} 条）\n")
    for it in rep.priority():
        print(f"- **[{tag[it['level']]}] {it['category']}** — {it['msg']}"
              + (f"  `{it['loc']}`" if it["loc"] else ""))
    print(f"\n已写入 {args.report}")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    cli_guard(main)
