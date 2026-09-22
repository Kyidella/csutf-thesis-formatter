"""格式化器。

目前覆盖三层，改动面从小到大：
  1. 文档级 —— 页边距 / 页面尺寸 / 页眉页脚边距 / 页码格式（改 w:sectPr）
  2. 样式级 —— 把规则写进 word/styles.xml 的样式定义（见 wordstyles.py）
  3. 文本级 —— 题注缺空格 / 异常字符 / 半角标点（见 textfix.py）

前两层**不碰正文 run**，第三层会改字但**只动 w:t**，对 Zotero 引用域零风险
（见 field_safety_test.py 与 textfix.py 开头的三条纪律）。

为什么样式级先做而不是逐段写直接格式：
  某样本论文的二级标题的 ascii 字体是"宋体"（"1.2" 显示成宋体），27 段全错，
  根因在样式定义里。改 1 处样式就能全好，比动 27 段的几百个 run 面小得多。
  改动越少，出错面越小。

两条硬约束：
  1. 只改 word/document.xml 与 word/styles.xml，其余 zip 部件逐字节照抄
     —— 由此可验证"没动别的"
  2. 永远另存新文件，绝不原地修改

用法:
    python formatter.py input.docx --rules rules/xxx.yaml --output fixed.docx [--dry-run]
"""

import argparse
import sys
import zipfile
from dataclasses import dataclass
from pathlib import Path

from lxml import etree

from classifier import classify
from check_docx import SKIP_FIX_KEYS, cli_guard, missing_sections
from docmodel import (TWIPS_PER_CM, has_drawing, has_inline_drawing,
                      load, stage)
from textfix import fix_text
from wordstyles import STYLES, StyleChange, ensure_styles

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
DOCUMENT = "word/document.xml"


@dataclass
class Change:
    section: int | None
    what: str
    old: str
    new: str

    def render(self):
        loc = f"节 {self.section}" if self.section is not None else "全文"
        return f"{loc}：{self.what} {self.old} → {self.new}"


def cm_to_twips(v):
    return str(int(round(v * TWIPS_PER_CM)))


def _get_margins_twips(sect):
    pgmar = sect.find(f"{W}pgMar")
    if pgmar is None:
        pgmar = sect.makeelement(f"{W}pgMar", {})
        sect.append(pgmar)
    return pgmar


def _fmt_cm(tw):
    return f"{round(int(tw) / TWIPS_PER_CM, 2)}cm" if tw else "未设置"


def _same_twips(cur, target, tol=2):
    """按数值比较，容差 2 twips（约 0.004cm）。

    直接比字符串会把 850 与 851 判成不同，于是报出 "1.5cm → 1.5cm" 这种假改动。
    """
    if cur is None:
        return False
    try:
        return abs(int(cur) - int(target)) <= tol
    except (TypeError, ValueError):
        return False


# --------------------------------------------------------------------------
# 改页边距
# --------------------------------------------------------------------------
def fix_margins(doc, rules, changes, notes):
    want = rules["page"]["margins_cm"]
    hdr = rules["page"].get("header_margin_cm")
    ftr = rules["page"].get("footer_margin_cm")

    for sec in doc.sections:
        pgmar = _get_margins_twips(sec.element)
        for tag in ("top", "bottom", "left", "right"):
            cur = pgmar.get(f"{W}{tag}")
            target = cm_to_twips(want[tag])
            if _same_twips(cur, target):
                continue
            changes.append(Change(sec.index, f"{tag}边距",
                                  _fmt_cm(cur), _fmt_cm(target)))
            pgmar.set(f"{W}{tag}", target)
        if hdr:
            cur = pgmar.get(f"{W}header")
            target = cm_to_twips(hdr)
            if not _same_twips(cur, target):
                changes.append(Change(sec.index, "页眉边距",
                                      _fmt_cm(cur), _fmt_cm(target)))
                pgmar.set(f"{W}header", target)
        if ftr:
            cur = pgmar.get(f"{W}footer")
            target = cm_to_twips(ftr)
            if not _same_twips(cur, target):
                changes.append(Change(sec.index, "页脚边距",
                                      _fmt_cm(cur), _fmt_cm(target)))
                pgmar.set(f"{W}footer", target)


# --------------------------------------------------------------------------
# 改页面尺寸
# --------------------------------------------------------------------------
def fix_page_size(doc, rules, changes, notes):
    w_tw = cm_to_twips(rules["page"]["width_mm"] / 10)
    h_tw = cm_to_twips(rules["page"]["height_mm"] / 10)

    for sec in doc.sections:
        pgsz = sec.element.find(f"{W}pgSz")
        if pgsz is None:
            pgsz = sec.element.makeelement(f"{W}pgSz", {})
            sec.element.insert(0, pgsz)
        cw, ch = pgsz.get(f"{W}w"), pgsz.get(f"{W}h")
        # 横向页不动——旋转页在论文里通常是刻意的（大幅插图）
        if pgsz.get(f"{W}orient") == "landscape":
            notes.append(f"节 {sec.index}：横向页，跳过页面尺寸")
            continue
        if not (_same_twips(cw, w_tw) and _same_twips(ch, h_tw)):
            changes.append(Change(sec.index, "页面尺寸",
                                  f"{_fmt_cm(cw)}×{_fmt_cm(ch)}",
                                  f"{_fmt_cm(w_tw)}×{_fmt_cm(h_tw)}"))
            pgsz.set(f"{W}w", w_tw)
            pgsz.set(f"{W}h", h_tw)


# --------------------------------------------------------------------------
# 改页码格式
# --------------------------------------------------------------------------
def plan_numbering(doc, rows, rules, notes):
    """规划每节该用哪种页码格式。

    规范：前置部分（摘要到目录）大写罗马数字；正文（绪论起）阿拉伯数字重新从 1 开始。
    封面与声明页不在规范里，保持原样不动。
    用哪种格式、从几开始，取 rules["pagination"]——以前这里是写死的
    "upperRoman"/"decimal"，规则改了也不生效（审查发现的）。
    """
    pg = rules.get("pagination", {})
    front = pg.get("front_matter", {}) or {}
    body_cfg = pg.get("body", {}) or {}

    sec_of = {r.key: r.para.section for r in rows if r.key in (
        "abstract_title", "toc_title", "heading1")}

    fm = sec_of.get("abstract_title")
    if fm is None:
        notes.append("未找到「摘要」，页码格式按原样保留")
        return {}

    body = None
    for r in rows:
        if r.key == "heading1" and r.para.section >= fm:
            body = r.para.section
            break
    if body is None:
        notes.append("摘要之后未找到一级标题，正文起页无法确定，页码按原样保留")
        return {}

    def fmt_of(d, default):
        return str(d.get("num_fmt", default))

    def start_of(d):
        return str(d.get("restart_at", 1))

    plan = {}
    for sec in doc.sections:
        if fm <= sec.index < body:
            plan[sec.index] = (fmt_of(front, "upperRoman"), start_of(front))
        elif sec.index >= body:
            plan[sec.index] = (fmt_of(body_cfg, "decimal"), start_of(body_cfg))
    return plan


NUMFMT_CN = {"decimal": "阿拉伯数字", "upperRoman": "大写罗马数字",
             "lowerRoman": "小写罗马数字"}


def fix_numbering(doc, rows, rules, changes, notes):
    plan = plan_numbering(doc, rows, rules, notes)
    for sec in doc.sections:
        if sec.index not in plan:
            continue
        fmt, start = plan[sec.index]
        pn = sec.element.find(f"{W}pgNumType")
        if pn is None:
            pn = sec.element.makeelement(f"{W}pgNumType", {})
            sect_order = ("pgSz", "pgMar", "cols", "docGrid")
            idx = len(sec.element)
            for i, child in enumerate(sec.element):
                if etree.QName(child).localname in sect_order:
                    idx = i + 1
            sec.element.insert(idx, pn)

        old_fmt, old_start = pn.get(f"{W}fmt"), pn.get(f"{W}start")
        if old_fmt == fmt and old_start == start:
            continue

        def desc(f, s):
            return (NUMFMT_CN.get(f, f or "默认")
                    + (f"，起始 {s}" if s else "，未设起始"))

        changes.append(Change(sec.index, "页码格式",
                              desc(old_fmt, old_start), desc(fmt, start)))
        pn.set(f"{W}fmt", fmt)
        pn.set(f"{W}start", start)


# --------------------------------------------------------------------------
# 段落映射到样式
# --------------------------------------------------------------------------
# 目录条目由 TOC 域生成，样式由 Word 自己套（toc 1/2/3）；
# 我们改 pStyle 会在刷新目录时被覆盖，所以跳过。名单在 check_docx 里
# ——检查报告要按同一个名单说"建议手动修改"，两处共用一份，免得各说各话。

# 与目标样式冲突的直接格式：清掉它们，样式才能生效。
# 只清我们管的这几项——加粗、斜体、上下标等有语义的格式一律不动。
PPR_STRIP = ("spacing", "ind", "jc", "outlineLvl")
RPR_STRIP = ("sz", "szCs", "rFonts")


def _strip_conflicting(ppr, pel):
    """清掉段落级与 run 级上会盖住样式的直接格式。"""
    n = 0
    if ppr is not None:
        for tag in PPR_STRIP:
            node = ppr.find(f"{W}{tag}")
            if node is not None:
                ppr.remove(node)
                n += 1
    for run in pel.findall(f"{W}r"):
        rpr = run.find(f"{W}rPr")
        if rpr is None:
            continue
        for tag in RPR_STRIP:
            node = rpr.find(f"{W}{tag}")
            if node is not None:
                rpr.remove(node)
                n += 1
        if len(rpr) == 0:
            run.remove(rpr)
    return n


# --------------------------------------------------------------------------
# 图片段落：不套样式，还要把固定行距拿掉
# --------------------------------------------------------------------------
def _set_single_spacing(el):
    """给段落写一条"单倍行距"的直接格式。

    为什么要动图片段落：**固定行距（lineRule=exact）会把嵌入型图片裁掉**
    ——这是 Word/WPS 的老行为，图越高裁得越狠，看起来就是"图片位置很奇怪"。
    环绕型图片浮在文字层上，不受行高影响，所以只有嵌入型出问题。

    写 line=240 + lineRule=auto 就是单倍（240/240），比只写 lineRule 明确。
    """
    ppr = el.find(f"{W}pPr")
    if ppr is None:
        ppr = el.makeelement(f"{W}pPr", {})
        el.insert(0, ppr)
    sp = ppr.find(f"{W}spacing")
    if sp is None:
        sp = ppr.makeelement(f"{W}spacing", {})
        ppr.append(sp)
    changed = not (sp.get(f"{W}lineRule") == "auto" and sp.get(f"{W}line") == "240")
    sp.set(f"{W}lineRule", "auto")
    sp.set(f"{W}line", "240")
    return changed


def _style_uses_fixed_spacing(rules, key):
    """这个类型要套的样式给的是不是固定行距（exact / atLeast）。"""
    ls = ((rules.get("styles") or {}).get(key) or {}).get("line_spacing") or {}
    return ls.get("type") == "exact"      # atLeast 不会裁（行可以变高）


def apply_paragraph_styles(doc, rows, rules, style_map, changes, notes):
    """按分类结果把段落指到对应的 Word 样式，并清掉冲突的直接格式。

    跳过的两类：
      - 含引用域的段落：格式改动会被 Zotero 刷新覆盖，且会触发它的
        "引注已被修改"提示，对用户是纯干扰（见 zotero_refresh_test.py）
      - 目录条目：由 TOC 域生成，改了会被覆盖
    """
    stat = {"styled": 0, "stripped": 0, "skip_field": 0, "skip_toc": 0,
            "same": 0, "skip_image": 0, "unclipped": 0}

    for r in rows:
        key = r.key
        # 表格内容与封面不在 word_styles 里，会在这一步被过滤掉——
        # 分类器保证表格段落只会是 table_content，所以无需再判 in_table。
        if key not in style_map:
            continue
        p = r.para
        if key in SKIP_FIX_KEYS:
            stat["skip_toc"] += 1
            continue
        if p.has_field:
            stat["skip_field"] += 1
            continue

        # 图片段落分两种处理（一开始一刀切跳过，结果把"3.3"这种
        # 图片锚在标题里的段落一起跳过了，标题就没排上）：
        #   - 图片独占一段（段里没字）：整段跳过。本来就没东西可排，
        #     套正文样式只会给图片加上首行缩进。
        #   - 段里有字（标题/题注和图片挤在一起）：照常排版，但见下面的行距护栏。
        if has_drawing(p.element) and not p.text.strip():
            stat["skip_image"] += 1
            # 跳过归跳过，但它自己带着的固定行距得救回来——否则图还是被裁的。
            ls = p.pformat.get("line_spacing")
            if ls and ls[0] == "exact" and has_inline_drawing(p.element):
                if _set_single_spacing(p.element):
                    stat["unclipped"] += 1
            continue

        el = p.element
        ppr = el.find(f"{W}pPr")
        if ppr is None:
            ppr = el.makeelement(f"{W}pPr", {})
            el.insert(0, ppr)

        sid = style_map[key]
        cur = ppr.find(f"{W}pStyle")
        if cur is not None and cur.get(f"{W}val") == sid:
            stat["same"] += 1
        else:
            if cur is None:
                cur = ppr.makeelement(f"{W}pStyle", {})
                ppr.insert(0, cur)
            cur.set(f"{W}val", sid)
            stat["styled"] += 1

        removed = _strip_conflicting(ppr, el)
        if removed:
            stat["stripped"] += 1

        # 行距护栏：这一段有嵌入型图片，而它要套的样式给的是固定行距——
        # 固定行距会把图裁掉，所以在这一段上用直接格式钉成单倍。
        # 只在这种情况下动手，别的图片段落不多写一个字（免得幂等性破功）。
        if has_inline_drawing(el) and _style_uses_fixed_spacing(rules, key):
            if _set_single_spacing(el):
                stat["unclipped"] += 1

    if stat["styled"]:
        changes.append(StyleChange(
            "（段落映射）", "指到规范样式",
            "—", f"{stat['styled']} 段改动，{stat['same']} 段已正确"))

    if stat["skip_field"]:
        notes.append(
            f"跳过（含引用域，格式会被 Zotero 刷新覆盖）：{stat['skip_field']} 段")
    if stat["skip_image"]:
        notes.append(f"跳过（图片独占一段，没有文字可排）：{stat['skip_image']} 段")
    if stat["unclipped"]:
        notes.append(f"{stat['unclipped']} 段里有嵌入型图片，行距已钉成单倍"
                     "（固定行距会把嵌入型图片裁掉）")
    notes.append(f"跳过（目录条目，由 TOC 域生成）：{stat['skip_toc']} 段")
    return stat


# --------------------------------------------------------------------------
def apply(input_path, rules, output_path=None, dry_run=False, on_stage=None):
    """返回 (changes, notes, style_map, stat)。changes 里既有节级的 Change 也有 StyleChange。

    dry_run=True 时只报告不写文件。两种 change 都有 .what / .render()。

    on_stage(说明, 已完成, 总数) 是给 GUI 的进度回调，不传就是纯命令行行为。
    """
    total = 6
    stage(on_stage, "正在读取文档…", 0, total)
    doc = load(input_path)
    stage(on_stage, "正在分类段落…", 1, total)
    rows = classify(doc)
    changes, notes = [], []

    stage(on_stage, "正在改文档级格式（页边距/页面/页码）…", 2, total)
    fix_margins(doc, rules, changes, notes)
    fix_page_size(doc, rules, changes, notes)
    fix_numbering(doc, rows, rules, changes, notes)

    stage(on_stage, "正在修文本（题注空格/异常字符/标点）…", 3, total)
    fix_text(doc, rows, rules, changes, notes)

    stage(on_stage, "正在写样式定义…", 4, total)
    styles_root = _load_styles_root(input_path)
    style_map = ensure_styles(styles_root, rules, changes, notes)

    stage(on_stage, "正在把段落指到样式…", 5, total)
    stat = apply_paragraph_styles(doc, rows, rules, style_map,
                                  changes, notes)

    # 缺章节目录在缺摘要、缺参考文献的稿子上排版时，改动清单看着一切正常，
    # 用户会以为文档没问题。结构化放进 stat，界面才能单独醒目地提示。
    miss = missing_sections(doc, rules)
    stat["missing_sections"] = miss
    if miss:
        notes.append(f"这份文档没有检测到：{'、'.join(miss)}"
                     "——自动排版只整理已有的内容，缺的部分不会凭空出现；"
                     "而且分类定位依赖这些章节，缺得越多、结果越不可靠")

    if changes and output_path and not dry_run:
        stage(on_stage, "正在写入文件…", 6, total)
        _write(input_path, doc.root, styles_root, output_path)
    return changes, notes, style_map, stat


def _load_styles_root(path):
    with zipfile.ZipFile(path) as z:
        return etree.fromstring(z.read(STYLES))


def _write(input_path, doc_root, styles_root, output_path):
    """只替换 document.xml 与 styles.xml，其余部件逐字节照抄。

    由此可验证"没动到别的东西"：比对其余部件的哈希即可。
    """
    src = Path(input_path)
    dst = Path(output_path)
    if src.resolve() == dst.resolve():
        raise ValueError("拒绝原地修改：输出路径与输入相同")

    def dump(root):
        return etree.tostring(root, xml_declaration=True,
                              encoding="UTF-8", standalone=True)

    new = {DOCUMENT: dump(doc_root), STYLES: dump(styles_root)}

    with zipfile.ZipFile(src) as zin, \
            zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = new.get(item.filename) or zin.read(item.filename)
            zout.writestr(item, data)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("input")
    ap.add_argument("--rules", required=True)
    ap.add_argument("--output", required=True)
    ap.add_argument("--dry-run", action="store_true")
    args = ap.parse_args()

    import yaml
    rules = yaml.safe_load(open(args.rules, encoding="utf-8"))
    changes, notes, _, _ = apply(args.input, rules, args.output, args.dry_run)

    print(f"# 格式{'预演' if args.dry_run else '修正'}")
    print(f"\n- 源文件：`{args.input}`")
    if not args.dry_run:
        print(f"- 输出：`{args.output}`")
    print(f"- 改动 {len(changes)} 处\n")
    for c in changes:
        print(f"- {c.render()}")
    if notes:
        print("\n## 跳过")
        for n in notes:
            print(f"- {n}")
    if args.dry_run:
        print("\n（预演模式，未写出文件）")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    cli_guard(main)
