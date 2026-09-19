"""文档模型层：一次解析 .docx，产出统一结构，供检查器与格式化器共用。

为什么需要这一层：
  - python-docx 的 doc.paragraphs 只遍历 body 的直接子级，会漏掉内容控件(SDT)
    内的段落和所有表格内的段落（实测某样本 1623 段漏掉 930 段，含整份目录）。
  - 各模块若各自解析一遍，节索引口径、段落口径就会不一致（已经踩过：
    页边距检查用 doc.sections=13，页码检查用 body.iter(sectPr)=14）。
  - 格式化器需要在解析时就知道「这段在不在引用域里」，才能避让。

有效格式（字体/字号/对齐/行距/缩进）在这里一次算清：
  run 级直接格式（按字符数加权）→ 样式链 → 文档默认。
不要再让各模块各写一份，那正是口径不一致的来源。

用法:
    from docmodel import load
    doc = load("thesis.docx")
    doc.paragraphs        # 全部段落，含表格与 SDT 内的
    doc.body_paragraphs   # 仅 body 直接子级
    doc.sections          # 节，携带 sectPr 元素
"""

from collections import Counter
from dataclasses import dataclass, field as dc_field
from functools import cached_property
import zipfile

from lxml import etree

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
EMU_PER_CM = 360000.0
# 注意两套单位：pgMar / pgSz 用 twips（1/20 磅），行高缩进等也用 twips，
# 但 python-docx 的 Length 是 EMU。1440 twips = 1 英寸 = 2.54cm。
TWIPS_PER_CM = 1440 / 2.54
ALIGN_MAP = {"start": "left", "end": "right"}
FIELD_MARKERS = ("ZOTERO_ITEM", "ZOTERO_BIBL", "ADDIN EN", "Mendeley")


def _tag(el):
    return etree.QName(el).localname


DRAWING_TAGS = ("drawing", "pict", "blip", "object")


def has_drawing(el):
    """段落里有没有图形（嵌入型 wp:inline 或环绕型 wp:anchor）。"""
    return any(node.tag.split("}")[-1] in DRAWING_TAGS for node in el.iter())


def has_inline_drawing(el):
    """段落里有没有**嵌入型**图片。

    只有嵌入型怕固定行距——它坐在行里，行高被锁死就超出部分被裁掉；
    环绕型浮在文字层之上，不受行高影响。所以"要不要把行距钉成单倍"
    只看这一种。
    """
    return any(node.tag.split("}")[-1] == "inline" for node in el.iter())


def stage(on_stage, msg, done, total):
    """进度回调的统一入口：没传回调就什么都不做。

    GUI 靠它把"正在读取文档…"这类阶段推给进度条。核心模块只在阶段边界调一下，
    不碰任何 Qt 的东西——所以命令行跑起来和以前完全一样。
    """
    if on_stage is not None:
        on_stage(msg, done, total)


def _text_of(el):
    parts = []
    for node in el.iter():
        t = _tag(node)
        if t == "t":
            parts.append(node.text or "")
        elif t == "tab":
            parts.append("\t")
        elif t in ("br", "cr"):
            parts.append("\n")
    return "".join(parts)


def _rfonts(rpr_owner):
    """从"带 rPr 的元素"（如 run）取字体。"""
    rf = rpr_owner.find(f"{W}rPr/{W}rFonts") if rpr_owner is not None else None
    if rf is None:
        return None, None
    return rf.get(f"{W}eastAsia"), rf.get(f"{W}ascii")


def _rfonts_of_rpr(rpr):
    """从 rPr 元素本身取字体（docDefaults 那一层是裸的 rPr）。"""
    rf = rpr.find(f"{W}rFonts") if rpr is not None else None
    if rf is None:
        return None, None
    return rf.get(f"{W}eastAsia"), rf.get(f"{W}ascii")


def _sz_of_rpr(rpr):
    if rpr is None:
        return None
    node = rpr.find(f"{W}sz")
    return round(int(node.get(f"{W}val")) / 2.0, 2) if node is not None else None


def _sz(rpr_owner):
    if rpr_owner is None:
        return None
    return _sz_of_rpr(rpr_owner.find(f"{W}rPr"))


def _cm(v):
    """EMU → cm（python-docx 的 Length 口径）。"""
    return None if v is None else round(v / EMU_PER_CM, 2)


def _tw_cm(v):
    """twips → cm（OOXML 里 pgMar / pgSz 的口径）。"""
    return None if v is None else round(v / TWIPS_PER_CM, 2)


# --------------------------------------------------------------------------
# 样式表
# --------------------------------------------------------------------------
def load_styles(z):
    """返回 (styleId → {name, basedOn, element}, 默认段落样式id)。

    默认样式必须单独取出来：没有 pStyle 的段落走的就是它（通常是 Normal），
    不补这一步，字号和行距会算不出来。
    """
    if "word/styles.xml" not in z.namelist():
        return {}, None
    root = etree.fromstring(z.read("word/styles.xml"))
    out, default_id = {}, None
    for s in root.findall(f"{W}style"):
        sid = s.get(f"{W}styleId")
        if not sid:
            continue
        nm = s.find(f"{W}name")
        base = s.find(f"{W}basedOn")
        typ = s.get(f"{W}type")
        out[sid] = {
            "name": nm.get(f"{W}val") if nm is not None else None,
            "basedOn": base.get(f"{W}val") if base is not None else None,
            "element": s,
        }
        if s.get(f"{W}default") == "1" and typ == "paragraph":
            default_id = sid
    return out, default_id


def load_theme(z):
    """读主题字体表。字体常以 w:asciiTheme / w:eastAsiaTheme 引用，
    不解析出来就永远算作"未设置"（python-docx 自带模板就是这样）。
    """
    if "word/theme/theme1.xml" not in z.namelist():
        return {}
    A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
    root = etree.fromstring(z.read("word/theme/theme1.xml"))
    sch = root.find(f".//{A}fontScheme")
    if sch is None:
        return {}
    out = {}
    for kind in ("majorFont", "minorFont"):
        f = sch.find(f"{A}{kind}")
        if f is None:
            continue
        for which in ("latin", "ea"):
            node = f.find(f"{A}{which}")
            tf = node.get("typeface") if node is not None else None
            if tf:
                out[f"{kind}_{which}"] = tf
    return out


def _theme_font(theme, rf):
    """把 rFonts 上的主题引用翻译成实际字体名。"""
    if rf is None or not theme:
        return None, None

    def look(attr):
        if not attr:
            return None
        kind = "majorFont" if attr.startswith("major") else "minorFont"
        which = "ea" if "EastAsia" in attr else "latin"
        return theme.get(f"{kind}_{which}")

    return (look(rf.get(f"{W}eastAsiaTheme")),
            look(rf.get(f"{W}asciiTheme")))


def load_defaults(z):
    """读 w:docDefaults —— 样式链走到头之后还有这一层。

    很多文档的字体/字号只写在文档级默认里（Normal 样式本身不设），
    不读这层就会算成 None。
    """
    if "word/styles.xml" not in z.namelist():
        return {}
    root = etree.fromstring(z.read("word/styles.xml"))
    rpr = root.find(f"{W}docDefaults/{W}rPrDefault/{W}rPr")
    if rpr is None:
        return {}
    ea, asc = _rfonts_of_rpr(rpr)
    return {"font_cn": ea, "font_ascii": asc, "size_pt": _sz_of_rpr(rpr)}


def _chain(styles, sid):
    seen, out = set(), []
    while sid and sid in styles and sid not in seen:
        seen.add(sid)
        out.append(styles[sid])
        sid = styles[sid]["basedOn"]
    return out


def _from_styles(styles, sid, getter):
    for info in _chain(styles, sid):
        v = getter(info["element"])
        if v is not None:
            return v
    return None


def _style_outline(styles, sid):
    def g(el):
        node = el.find(f"{W}pPr/{W}outlineLvl")
        return int(node.get(f"{W}val")) if node is not None else None
    return _from_styles(styles, sid, g)


def _fonts_chain(styles, sid):
    """中文字体与西文字体必须各自独立沿样式链回退。

    不能合成一个 getter：某个样式可能只定义了 eastAsia 没定义 ascii，
    这时 ascii 要继续往上找（否则 "1.2" 这类数字的字体就判不出来）。
    """
    ea = asc = None
    for info in _chain(styles, sid):
        e, a = _rfonts(info["element"])
        if ea is None:
            ea = e
        if asc is None:
            asc = a
        if ea and asc:
            break
    return ea, asc


# --------------------------------------------------------------------------
# 段落级有效格式
# --------------------------------------------------------------------------
def _dominant_run_format(el):
    """run 级直接格式，按字符数加权取众数。

    取「第一个 run」是错的：中英混排段落里，第一个 run 常常不是主体。
    """
    fonts, sizes = Counter(), Counter()
    for r in el.findall(f"{W}r"):
        n = len("".join(x.text or "" for x in r.iter(f"{W}t"))) or 1
        ea, asc = _rfonts(r)
        if ea or asc:
            fonts[(ea, asc)] += n
        sz = _sz(r)
        if sz is not None:
            sizes[sz] += n
    font = fonts.most_common(1)[0][0] if fonts else (None, None)
    size = sizes.most_common(1)[0][0] if sizes else None
    return font[0], font[1], size


def _para_metrics(ppr):
    """从单个 pPr 取出行距 / 段前段后 / 缩进 / 对齐。"""
    m = {}
    sp = ppr.find(f"{W}spacing") if ppr is not None else None
    ind = ppr.find(f"{W}ind") if ppr is not None else None
    jc = ppr.find(f"{W}jc") if ppr is not None else None

    if sp is not None:
        line, rule = sp.get(f"{W}line"), sp.get(f"{W}lineRule")
        if line:
            if rule in ("exact", "atLeast"):
                m["line_spacing"] = (rule, round(int(line) / 20, 2))
            else:
                m["line_spacing"] = ("multiple", round(int(line) / 240, 3))
        bl, b = sp.get(f"{W}beforeLines"), sp.get(f"{W}before")
        if bl is not None:
            m["space_before"] = ("lines", int(bl) / 100)
        elif b is not None:
            m["space_before"] = ("pt", int(b) / 20)
        al, a = sp.get(f"{W}afterLines"), sp.get(f"{W}after")
        if al is not None:
            m["space_after"] = ("lines", int(al) / 100)
        elif a is not None:
            m["space_after"] = ("pt", int(a) / 20)

    if ind is not None:
        flc, fl = ind.get(f"{W}firstLineChars"), ind.get(f"{W}firstLine")
        if flc is not None:
            m["first_line_chars"] = int(flc) / 100
        elif fl is not None:
            m["first_line_pt"] = round(int(fl) / 20, 2)
        lc, l = ind.get(f"{W}leftChars"), ind.get(f"{W}left")
        if lc is not None:
            m["left_chars"] = int(lc) / 100
        elif l is not None:
            m["left_pt"] = round(int(l) / 20, 2)

    if jc is not None and jc.get(f"{W}val"):
        v = jc.get(f"{W}val")
        m["align"] = ALIGN_MAP.get(v, v)
    return m


def _merge_metrics(styles, sid, ppr):
    """段落级格式：段落 pPr 优先，缺的走样式链。"""
    out = dict(_para_metrics(ppr))
    keys = ("line_spacing", "space_before", "space_after",
            "first_line_chars", "first_line_pt", "left_chars", "left_pt", "align")

    def from_style(el, k):
        return _para_metrics(el.find(f"{W}pPr")).get(k)

    for key in keys:
        if key in out:
            continue
        v = _from_styles(styles, sid, lambda el, k=key: from_style(el, k))
        if v is not None:
            out[key] = v
    return out


# --------------------------------------------------------------------------
# 数据结构
# --------------------------------------------------------------------------
@dataclass
class Para:
    index: int
    element: object
    text: str
    style_id: str | None
    style_name: str | None
    outline_level: int | None
    in_table: bool
    in_sdt: bool
    section: int
    has_field: bool
    pformat: dict = dc_field(default_factory=dict)

    @property
    def is_empty(self):
        return not self.text.strip()

    @property
    def align(self):
        return self.pformat.get("align")

    @cached_property
    def size_pt(self):
        return self.pformat.get("size_pt")

    @cached_property
    def font_cn(self):
        return self.pformat.get("font_cn")

    @cached_property
    def font_ascii(self):
        return self.pformat.get("font_ascii")

    def size_chars(self, chars):
        """首行缩进折算成字符数；未设置视为 0。"""
        v = self.pformat.get("first_line_chars")
        if v is None and "first_line_pt" in self.pformat and self.size_pt:
            v = round(self.pformat["first_line_pt"] / self.size_pt, 2)
        return 0.0 if v is None else v

    def left_chars(self):
        v = self.pformat.get("left_chars")
        if v is None and "left_pt" in self.pformat and self.size_pt:
            v = round(self.pformat["left_pt"] / self.size_pt, 2)
        return 0.0 if v is None else v


@dataclass
class Section:
    index: int
    element: object
    page_w_cm: float | None
    page_h_cm: float | None
    margins_cm: dict
    n_header_refs: int
    n_footer_refs: int
    pg_num_fmt: str | None
    pg_num_start: str | None
    footnote_fmt: str | None
    footnote_restart: str | None
    starts_at: int


@dataclass
class Doc:
    path: str
    paragraphs: list
    body_paragraphs: list
    sections: list
    fields: list
    stats: dict
    styles: dict
    root: object = None       # document.xml 的根元素，供格式化器回写

    @cached_property
    def by_style(self):
        out = {}
        for p in self.paragraphs:
            out.setdefault(p.style_name, []).append(p)
        return out

    def all_text(self, include_tables=False):
        src = self.paragraphs if include_tables else self.body_paragraphs
        return "\n".join(p.text for p in src)


def _walk(el, out, state):
    """按文档顺序递归产出段落，记录是否位于表格 / 内容控件内。"""
    for child in el:
        tag = _tag(child)
        if tag == "p":
            in_sect = child.find(f"{W}pPr/{W}sectPr") is not None
            out.append((child, state["in_table"], state["in_sdt"], in_sect))
        elif tag == "tbl":
            state["in_table"] = True
            _walk(child, out, state)
            state["in_table"] = False
        elif tag == "sdt":
            state["in_sdt"] = True
            content = child.find(f"{W}sdtContent")
            if content is not None:
                _walk(content, out, state)
            state["in_sdt"] = False
        elif tag in ("tr", "tc", "sdtContent"):
            _walk(child, out, state)


def load(path):
    z = zipfile.ZipFile(path)
    root = etree.fromstring(z.read("word/document.xml"))
    body = root.find(f"{W}body")
    styles, default_style = load_styles(z)
    defaults = load_defaults(z)
    theme = load_theme(z)

    raw = []
    _walk(body, raw, {"in_table": False, "in_sdt": False})

    paragraphs, fields = [], []
    section_idx, section_first = 0, [0]
    body_paragraphs = []

    for i, (el, in_table, in_sdt, in_sect) in enumerate(raw):
        ppr = el.find(f"{W}pPr")
        style_el = ppr.find(f"{W}pStyle") if ppr is not None else None
        style_id = style_el.get(f"{W}val") if style_el is not None else None
        style_name = styles.get(style_id, {}).get("name") if style_id else None
        chain_id = style_id or default_style   # 无 pStyle 时走默认样式

        ol = None
        if ppr is not None:
            node = ppr.find(f"{W}outlineLvl")
            if node is not None:
                ol = int(node.get(f"{W}val"))
        if ol is None:
            ol = _style_outline(styles, chain_id)

        text = _text_of(el)
        instr = " ".join(t.text or "" for t in el.iter(f"{W}instrText"))
        has_field = any(m in instr for m in FIELD_MARKERS)
        if has_field:
            for m in FIELD_MARKERS:
                if m in instr:
                    fields.append((i, m))
                    break

        # 有效字体/字号：run 级（加权众数）→ 样式链。
        # eastAsia 与 ascii 分别回退——混在一起会漏判西文字体。
        ea, asc, sz = _dominant_run_format(el)
        if ea is None or asc is None:
            cea, casc = _fonts_chain(styles, chain_id)
            ea = ea if ea is not None else cea
            asc = asc if asc is not None else casc
        if sz is None:
            sz = _from_styles(styles, chain_id, _sz)
        # 主题字体引用（w:asciiTheme="minorHAnsi" 这类）。
        # 注意：rFonts 在 run 里，不在 w:p 的直接子级（w:p 下面只有 pPr）。
        if ea is None or asc is None:
            rf = None
            for r in el.findall(f"{W}r"):
                rf = r.find(f"{W}rPr/{W}rFonts")
                if rf is not None:
                    break
            if rf is None:
                for info in _chain(styles, chain_id):
                    rf = info["element"].find(f"{W}rPr/{W}rFonts")
                    if rf is not None:
                        break
            tea, tasc = _theme_font(theme, rf)
            ea = ea if ea is not None else tea
            asc = asc if asc is not None else tasc
        # 最后一层：文档级默认
        ea = ea if ea is not None else defaults.get("font_cn")
        asc = asc if asc is not None else defaults.get("font_ascii")
        sz = sz if sz is not None else defaults.get("size_pt")

        pformat = _merge_metrics(styles, chain_id, ppr)
        pformat["font_cn"], pformat["font_ascii"], pformat["size_pt"] = ea, asc, sz

        para = Para(index=i, element=el, text=text, style_id=style_id,
                    style_name=style_name, outline_level=ol, in_table=in_table,
                    in_sdt=in_sdt, section=section_idx, has_field=has_field,
                    pformat=pformat)
        paragraphs.append(para)
        if not in_table and not in_sdt:
            body_paragraphs.append(para)

        if in_sect:
            section_idx += 1
            section_first.append(i + 1)

    # ---- 节 ----
    sect_els = list(body.iter(f"{W}sectPr"))
    sections = []
    for k, s in enumerate(sect_els):
        # 注意：pgMar / pgSz 是元素上的属性，不是子元素
        pgmar = s.find(f"{W}pgMar")
        m = {}
        for tag in ("top", "bottom", "left", "right"):
            v = pgmar.get(f"{W}{tag}") if pgmar is not None else None
            m[tag] = _tw_cm(int(v)) if v else None
        pgsz = s.find(f"{W}pgSz")
        pw = pgsz.get(f"{W}w") if pgsz is not None else None
        ph = pgsz.get(f"{W}h") if pgsz is not None else None
        pn = s.find(f"{W}pgNumType")
        fp = s.find(f"{W}footnotePr")
        nf = fp.find(f"{W}numFmt") if fp is not None else None
        nr = fp.find(f"{W}numRestart") if fp is not None else None
        sections.append(Section(
            index=k, element=s,
            page_w_cm=_tw_cm(int(pw)) if pw else None,
            page_h_cm=_tw_cm(int(ph)) if ph else None,
            margins_cm=m,
            n_header_refs=len(s.findall(f"{W}headerReference")),
            n_footer_refs=len(s.findall(f"{W}footerReference")),
            pg_num_fmt=pn.get(f"{W}fmt") if pn is not None else None,
            pg_num_start=pn.get(f"{W}start") if pn is not None else None,
            footnote_fmt=nf.get(f"{W}val") if nf is not None else None,
            footnote_restart=nr.get(f"{W}val") if nr is not None else None,
            starts_at=section_first[k] if k < len(section_first) else 0,
        ))

    stats = {
        "paragraphs_total": len(paragraphs),
        "paragraphs_body": len(body_paragraphs),
        "paragraphs_in_table": sum(1 for p in paragraphs if p.in_table),
        "paragraphs_in_sdt": sum(1 for p in paragraphs if p.in_sdt),
        "sections": len(sections),
        # 注意：这是"含引用域的段落数"，不是域的总个数
        "paragraphs_with_fields": len(fields),
    }
    return Doc(path=path, paragraphs=paragraphs, body_paragraphs=body_paragraphs,
               sections=sections, fields=fields, stats=stats, styles=styles,
               root=root)


if __name__ == "__main__":
    import sys
    sys.stdout.reconfigure(encoding="utf-8")
    d = load(sys.argv[1])
    for k, v in d.stats.items():
        print(f"  {k:24} {v}")
