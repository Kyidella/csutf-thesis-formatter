"""把 YAML 规则写进 Word 的样式定义（word/styles.xml）。

为什么走样式而不是逐段写直接格式：
  学生的"格式不对"有一半来自样式定义本身错了——某样本论文的二级标题的
  ascii 字体是"宋体"，导致 "1.2" 显示成宋体，27 段全错。改 1 处样式
  就能全好，比动 27 段的几百个 run 面小得多。

改动越少，出错面越小，这是选择这条路线的唯一理由。
"""

from dataclasses import dataclass

from lxml import etree

from docmodel import _rfonts_of_rpr, _sz_of_rpr

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
STYLES = "word/styles.xml"


@dataclass
class StyleChange:
    style: str
    what: str
    old: str
    new: str

    def render(self):
        return f"样式「{self.style}」{self.what}：{self.old} → {self.new}"


# --------------------------------------------------------------------------
# 取/建元素
# --------------------------------------------------------------------------
def _ensure(parent, tag, first=True):
    node = parent.find(f"{W}{tag}")
    if node is None:
        node = parent.makeelement(f"{W}{tag}", {})
        if first:
            parent.insert(0, node)
        else:
            parent.append(node)
    return node


def _rpr(style_el):
    return _ensure(style_el, "rPr")


def _ppr(style_el):
    return _ensure(style_el, "pPr")


def _set_rfonts(rpr, cn, en):
    rf = rpr.find(f"{W}rFonts")
    if rf is None:
        rf = rpr.makeelement(f"{W}rFonts", {})
        rpr.insert(0, rf)
    # 主题引用与显式字体名互斥：设了显式名就必须清掉 Theme 属性，
    # 否则 Word 会继续按主题取字体，我们写的名字不生效。
    for attr in ("asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme"):
        rf.attrib.pop(f"{W}{attr}", None)
    if cn:
        rf.set(f"{W}eastAsia", cn)
    if en:
        rf.set(f"{W}ascii", en)
        rf.set(f"{W}hAnsi", en)


def _set_size(rpr, size_pt):
    for tag in ("sz", "szCs"):
        node = rpr.find(f"{W}{tag}")
        if node is None:
            node = rpr.makeelement(f"{W}{tag}", {})
            rpr.append(node)
        node.set(f"{W}val", str(int(round(size_pt * 2))))


def _set_bold(rpr, bold):
    """显式写加粗开关。

    必须显式写、不能"不管"：Word 内置的 heading 2/3/4 样式自带 <w:b/>，
    我们改了字体字号却没管加粗，结果**把标题弄成了粗体**——规范要的是
    黑体 + 三号，从没说加粗（黑体本身已经够重，再叠加就是双重强调）。
    实测：用户拿另一份稿子测试时发现"标题不该粗的地方没更正"。

    位置有讲究：OOXML 的 rPr 子元素是有顺序的（rFonts → b → bCs → sz），
    所以塞在 rFonts 后面，不能直接 append 到末尾。
    """
    val = "1" if bold else "0"
    rf = rpr.find(f"{W}rFonts")
    idx = (list(rpr).index(rf) + 1) if rf is not None else 0
    for tag in ("b", "bCs"):
        node = rpr.find(f"{W}{tag}")
        if node is None:
            node = rpr.makeelement(f"{W}{tag}", {})
            rpr.insert(idx, node)
            idx += 1
        node.set(f"{W}val", val)


def _set_spacing(ppr, rule):
    sb, sa = rule.get("space_before"), rule.get("space_after")
    ls = rule.get("line_spacing")
    if not any((sb, sa, ls)):
        return
    sp = _ensure(ppr, "spacing")
    if ls:
        if ls["type"] in ("exact", "atLeast"):
            sp.set(f"{W}line", str(int(round(ls["value"] * 20))))
            sp.set(f"{W}lineRule", ls["type"])
        else:
            mult = 1.0 if ls["type"] == "single" else float(ls["value"])
            sp.set(f"{W}line", str(int(round(mult * 240))))
            sp.set(f"{W}lineRule", "auto")
    for tag, key in (("before", sb), ("after", sa)):
        if not key:
            continue
        if "lines" in key:
            sp.set(f"{W}{tag}Lines", str(int(round(key["lines"] * 100))))
            sp.attrib.pop(f"{W}{tag}", None)   # 有 *Lines 时磅值冗余
        else:
            sp.set(f"{W}{tag}", str(int(round(key["pt"] * 20))))
            sp.attrib.pop(f"{W}{tag}Lines", None)


def _set_indent(ppr, indent):
    if not indent:
        return
    ind = _ensure(ppr, "ind")
    for attr, key in (("firstLineChars", "first_line_chars"),
                      ("leftChars", "left_chars"),
                      ("rightChars", "right_chars")):
        if key not in indent or indent[key] is None:
            continue
        ind.set(f"{W}{attr}", str(int(round(indent[key] * 100))))
        # 字符数优先，清掉同义的磅值，避免两套单位打架
        ind.attrib.pop(f"{W}{attr.replace('Chars', '')}", None)


def _set_align(ppr, align):
    jc = ppr.find(f"{W}jc")
    if align is None:
        return
    if jc is None:
        jc = ppr.makeelement(f"{W}jc", {})
        ppr.append(jc)
    jc.set(f"{W}val", {"both": "both", "center": "center",
                       "left": "left", "right": "right"}[align])


def _set_outline(ppr, level):
    """规则里的 outline_level 是 1 起（1 = 一级标题）；OOXML 是 0 起。"""
    node = ppr.find(f"{W}outlineLvl")
    if level is None:
        if node is not None:
            ppr.remove(node)
        return
    if node is None:
        node = ppr.makeelement(f"{W}outlineLvl", {})
        ppr.append(node)
    node.set(f"{W}val", str(level - 1))


# --------------------------------------------------------------------------
def _read_style_rule_current(style_el):
    """读样式定义里我们关心的那几项，用于报告"改了什么"。"""
    rpr = style_el.find(f"{W}rPr")
    ppr = style_el.find(f"{W}pPr")
    cn, en = _rfonts_of_rpr(rpr) if rpr is not None else (None, None)
    b = rpr.find(f"{W}b") if rpr is not None else None
    has_b = None
    if b is not None:
        has_b = b.get(f"{W}val") not in ("0", "false")
    return {
        "font_cn": cn,
        "font_en": en,
        "size_pt": _sz_of_rpr(rpr) if rpr is not None else None,
        "bold": has_b,
        "align": (ppr.find(f"{W}jc").get(f"{W}val")
                  if ppr is not None and ppr.find(f"{W}jc") is not None else None),
    }


def style_id_map(styles_root):
    """样式名 → styleId。"""
    out = {}
    for s in styles_root.findall(f"{W}style"):
        nm = s.find(f"{W}name")
        if nm is not None and nm.get(f"{W}val"):
            out[nm.get(f"{W}val")] = s.get(f"{W}styleId")
    return out


# 内置样式的英文规范名 → 可能出现的其它写法（WPS 有时会存中文名）
BUILTIN_ALIASES = {
    "heading 1": ("heading 1", "标题 1", "标题1"),
    "heading 2": ("heading 2", "标题 2", "标题2"),
    "heading 3": ("heading 3", "标题 3", "标题3"),
    "heading 4": ("heading 4", "标题 4", "标题4"),
    "normal": ("normal", "正文"),
}


def find_style(styles_root, name, builtin=False):
    """按名字找样式；内置样式额外尝试别名。

    找不到就会新建——所以匹配错了的代价是给文档塞一套重复样式。
    """
    targets = {name.lower()}
    if builtin:
        for alias in BUILTIN_ALIASES.get(name.lower(), ()):
            targets.add(alias.lower())
    for s in styles_root.findall(f"{W}style"):
        nm = s.find(f"{W}name")
        if nm is not None and (nm.get(f"{W}val") or "").lower() in targets:
            return s
    return None


def _new_style_id(styles_root, name):
    existing = {s.get(f"{W}styleId") for s in styles_root.findall(f"{W}style")}
    base = "TF" + "".join(str(ord(c) % 10) for c in name)[:6]
    sid, i = base, 0
    while sid in existing:
        i += 1
        sid = f"{base}{i}"
    return sid


def ensure_styles(styles_root, rules, changes, notes):
    """按规则创建/修正样式定义。返回 类型键 → styleId 的映射。"""
    mapping = {}
    by_name = style_id_map(styles_root)
    default_id = None
    for s in styles_root.findall(f"{W}style"):
        if s.get(f"{W}default") == "1" and s.get(f"{W}type") == "paragraph":
            default_id = s.get(f"{W}styleId")

    for key, spec in rules.get("word_styles", {}).items():
        rule = rules["styles"].get(key)
        if rule is None:
            continue
        name = spec["name"]
        existing = find_style(styles_root, name, spec.get("builtin", False))

        if existing is None:
            sid = _new_style_id(styles_root, name)
            el = styles_root.makeelement(f"{W}style", {})
            el.set(f"{W}type", "paragraph")
            el.set(f"{W}styleId", sid)
            nm = el.makeelement(f"{W}name", {})
            nm.set(f"{W}val", name)
            el.append(nm)
            if default_id:
                bo = el.makeelement(f"{W}basedOn", {})
                bo.set(f"{W}val", default_id)
                el.append(bo)
            el.append(el.makeelement(f"{W}qFormat", {}))
            styles_root.append(el)
            by_name[name] = sid
            changes.append(StyleChange(name, "新建样式", "（无）", name))
        else:
            el = existing
            sid = existing.get(f"{W}styleId")

        before = _read_style_rule_current(el)
        rpr, ppr = _rpr(el), _ppr(el)

        if rule.get("font_cn") or rule.get("font_en"):
            _set_rfonts(rpr, rule.get("font_cn"), rule.get("font_en"))
        if rule.get("size_pt"):
            _set_size(rpr, rule["size_pt"])
        if "bold" in rule:
            _set_bold(rpr, bool(rule["bold"]))
        _set_spacing(ppr, rule)
        _set_indent(ppr, rule.get("indent"))
        _set_align(ppr, rule.get("align"))
        _set_outline(ppr, rule.get("outline_level"))

        after = _read_style_rule_current(el)
        for field, label, unit in (("font_cn", "中文字体", ""),
                                   ("font_en", "西文字体", ""),
                                   ("size_pt", "字号", "pt"),
                                   ("bold", "加粗", ""),
                                   ("align", "对齐", "")):
            if before[field] != after[field]:
                if field == "bold":
                    show = lambda v: {True: "加粗", False: "不加粗"}.get(v, "（继承）")
                    changes.append(StyleChange(name, label,
                                               show(before[field]), show(after[field])))
                    continue
                changes.append(StyleChange(
                    name, label,
                    f"{before[field]}{unit}" if before[field] else "（继承）",
                    f"{after[field]}{unit}" if after[field] else "（清除）"))
        mapping[key] = sid
    return mapping
