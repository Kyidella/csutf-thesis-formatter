"""把 .docx 解析成结构 JSON，用于验证"格式规则能否被结构化提取"。

用法:
    python dump_docx.py input.docx -o structure.json
"""

import argparse
import json
import sys
from collections import Counter

from docx import Document
from docx.shared import Pt, Emu

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

EMU_PER_CM = 360000.0


def _rfonts(el):
    """从 rPr 元素里取出 w:rFonts 的 eastAsia / ascii 字体名。"""
    if el is None:
        return None, None
    rpr = el.find(f"{W}rPr")
    if rpr is None:
        return None, None
    rf = rpr.find(f"{W}rFonts")
    if rf is None:
        return None, None
    return rf.get(f"{W}eastAsia"), rf.get(f"{W}ascii")


def _pt(size):
    if size is None:
        return None
    return round(size.pt, 2)


def _cm(length):
    if length is None:
        return None
    return round(length / EMU_PER_CM, 2)


def _outline_lvl_of(ppr):
    if ppr is None:
        return None
    ol = ppr.find(f"{W}outlineLvl")
    if ol is None:
        return None
    v = ol.get(f"{W}val")
    return int(v) if v is not None else None


def _outline_level(para):
    """返回段落的大纲级别（0 起），段落级优先，其次样式级。"""
    lvl = _outline_lvl_of(para._p.pPr)
    if lvl is not None:
        return lvl
    try:
        return _outline_lvl_of(para.style.element.pPr)
    except Exception:
        return None


def _list_level(para):
    ppr = para._p.pPr
    if ppr is None:
        return None
    npr = ppr.find(f"{W}numPr")
    if npr is None:
        return None
    ilvl = npr.find(f"{W}ilvl")
    numid = npr.find(f"{W}numId")
    return {
        "ilvl": int(ilvl.get(f"{W}val")) if ilvl is not None else None,
        "numId": int(numid.get(f"{W}val")) if numid is not None else None,
    }


def describe_run(run):
    """单个 run 的字体信息；未直接设置的字段为 None。"""
    el = run._element
    ea, ascii_ = _rfonts(el)
    return {
        "text": run.text,
        "bold": run.bold,
        "italic": run.italic,
        "font_eastasia": ea,
        "font_ascii": ascii_,
        "size_pt": _pt(run.font.size),
    }


ALIGN_MAP = {"start": "left", "end": "right"}


def _align(para):
    """直读 w:jc，绕开 python-docx 对 start/end 等新值的枚举限制。"""
    ppr = para._p.pPr
    if ppr is None:
        return None
    jc = ppr.find(f"{W}jc")
    if jc is None:
        return None
    v = jc.get(f"{W}val")
    return ALIGN_MAP.get(v, v)


def describe_paragraph(para, index):
    pf = para.paragraph_format
    runs = [describe_run(r) for r in para.runs]
    return {
        "index": index,
        "text": para.text,
        "text_head": para.text[:30],
        "style": para.style.name if para.style is not None else None,
        "outline_level": _outline_level(para),
        "num_pr": _list_level(para),
        "align": _align(para),
        "first_line_indent_cm": _cm(pf.first_line_indent),
        "left_indent_cm": _cm(pf.left_indent),
        "line_spacing": pf.line_spacing
        if not hasattr(pf.line_spacing, "pt")
        else round(pf.line_spacing.pt, 2),
        "line_spacing_rule": str(pf.line_spacing_rule)
        if pf.line_spacing_rule is not None
        else None,
        "space_before_pt": _pt(pf.space_before),
        "space_after_pt": _pt(pf.space_after),
        "runs": runs,
        # 该段是否完全没有任何 run 级直接格式（用来判断"是否靠样式排版"）
        "has_direct_run_fmt": any(
            r[1] is not None for r in
            ((k, v) for r in runs for k, v in r.items()
             if k in ("bold", "italic", "font_eastasia", "font_ascii", "size_pt"))
        ),
    }


def dump_styles(doc):
    out = []
    for s in doc.styles:
        ea, ascii_ = _rfonts(s.element)
        try:
            base = s.base_style.name if s.base_style is not None else None
        except Exception:
            base = None
        font = getattr(s, "font", None)
        ppr = s.element.pPr
        jc = ppr.find(f"{W}jc") if ppr is not None else None
        align = ALIGN_MAP.get(jc.get(f"{W}val"), jc.get(f"{W}val")) if jc is not None else None
        out.append({
            "name": s.name,
            "type": str(s.type),
            "builtin": s.builtin,
            "base_style": base,
            "font_eastasia": ea,
            "font_ascii": ascii_,
            "size_pt": _pt(font.size) if font is not None else None,
            "bold": font.bold if font is not None else None,
            "align": align,
            "outline_level": _outline_lvl_of(ppr),
        })
    return out


def dump_sections(doc):
    out = []
    for i, sec in enumerate(doc.sections):
        out.append({
            "index": i,
            "page_w_cm": _cm(sec.page_width),
            "page_h_cm": _cm(sec.page_height),
            "margins_cm": {
                "top": _cm(sec.top_margin),
                "bottom": _cm(sec.bottom_margin),
                "left": _cm(sec.left_margin),
                "right": _cm(sec.right_margin),
            },
            "orientation": str(sec.orientation),
            "different_first_page": sec.different_first_page_header_footer,
            "header_text": " | ".join(
                p.text for p in sec.header.paragraphs if p.text.strip()
            ),
            "footer_text": " | ".join(
                p.text for p in sec.footer.paragraphs if p.text.strip()
            ),
        })
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("input")
    ap.add_argument("-o", "--output", default="structure.json")
    args = ap.parse_args()

    doc = Document(args.input)

    paragraphs = [describe_paragraph(p, i) for i, p in enumerate(doc.paragraphs)]
    result = {
        "source": args.input,
        "counts": {
            "paragraphs": len(doc.paragraphs),
            "tables": len(doc.tables),
            "sections": len(doc.sections),
            "inline_shapes": len(doc.inline_shapes),
        },
        "sections": dump_sections(doc),
        "styles": dump_styles(doc),
        "paragraphs": paragraphs,
    }

    with open(args.output, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=1)

    # ---- 控制台摘要 ----
    print(f"段落 {len(doc.paragraphs)} | 表格 {len(doc.tables)} | 节 {len(doc.sections)}")
    print("\n[样式使用统计]")
    for name, n in Counter(p["style"] for p in paragraphs).most_common():
        sub = [p for p in paragraphs if p["style"] == name]
        direct = sum(1 for p in sub if p["has_direct_run_fmt"])
        print(f"  {n:5d}  {name:<24} 其中有直接格式 {direct}")
    print(f"\n已写入 {args.output}")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
