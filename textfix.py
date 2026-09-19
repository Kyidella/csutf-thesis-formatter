"""文本层修复：题注缺空格、异常字符、半角标点。

这一层和前面两层（文档级、样式级）的区别在于**它会改字**。因此纪律也最严：

  1. **只动 w:t 节点。** 绝不碰 w:instrText——Zotero 把自己的 JSON 存在那里，
     改坏一个字节，引用就成了废数据，而且它不会报错。
  2. **域内的 w:t 不动。** 用 w:fldChar begin/end 成对标记（可嵌套），
     加上整体成域的 w:fldSimple。这些内容由 Word / Zotero 重新生成，
     改了白改，还会触发"引注已被修改"的提示。
     注意这是**按节点**判的，不是"段落含引用域就整段不动"：引用标注旁边的
     普通文字（如"南方省(区、市)[4]"里的括号）不在域内，改了是安全的。
  3. **不新增、不删除、不合并 run。** 所有修复都是在已有 w:t 的文字里做增删改。
  4. 首尾出现空格必须补 `xml:space="preserve"`，否则 Word 会把空格吃掉。
     实测：NBSP 换成半角空格后若不加这个属性，英文单词会被粘在一起。

政策（"什么算问题"）写在 check_docx.py，本模块从那里取用，两处共用同一份判据，
避免出现"检查器报 8 处、格式化器只改 5 处"这种口径不一致。

用法（由 formatter.apply 调用，一般不单独用）:
    from textfix import fix_text
    fix_text(doc, rows, rules, changes, notes)
"""

import re
from collections import Counter
from dataclasses import dataclass

from docmodel import W

XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

# 半角 → 全角。只有这几个：方括号不在其内，理由见 punct_fixes。
FULLWIDTH = {"(": "（", ")": "）", ",": "，", ";": "；", ":": "："}


@dataclass
class TextChange:
    """一条文本层改动记录。一类问题一条：count 是改动个数，paras 是出处样例。"""

    what: str
    old: str
    new: str
    count: int
    paras: list

    def render(self):
        show = "、".join(f"段{i}" for i in self.paras[:3])
        more = " …" if len(self.paras) > 3 else ""
        return (f"（文本层）{self.what}：{self.old} → {self.new}，"
                f"共 {self.count} 处（{show}{more}）")


# --------------------------------------------------------------------------
# 位置映射：把段落文本的下标映射回具体的 w:t 节点
# --------------------------------------------------------------------------
def flatten(el):
    """把段落拍平成 (文本, 位置表)。

    位置表每项是 (w:t 节点, 起点, 终点)。口径必须与 docmodel._text_of 完全一致：
    tab 算一个 \\t、br 算一个 \\n、**不计 w:instrText**——偏一个字符，后续所有
    下标都会错位，改到不该改的地方。
    """
    parts, spans, pos = [], [], 0
    for node in el.iter():
        tag = node.tag.split("}")[-1]
        if tag == "t":
            s = node.text or ""
            if s:
                spans.append((node, pos, pos + len(s)))
        elif tag == "tab":
            s = "\t"
        elif tag in ("br", "cr"):
            s = "\n"
        else:
            continue
        parts.append(s)
        pos += len(s)
    return "".join(parts), spans


def node_at(spans, i):
    for node, a, b in spans:
        if a <= i < b:
            return node, i - a
    return None, None


def protected(el):
    """段落里落在域内的 w:t 节点（返回节点集合）。

    域由 w:fldChar 的 begin/separate/end 三元组成对标记，可以嵌套；
    w:fldSimple 则是整个元素自成一域，其内的 w:t 一律算域内。

    注意这里收集的是**元素本身**而不是 id(元素)：lxml 的元素代理是临时的，
    同一个节点每次访问都可能是新的 Python 对象，id() 会变（还可能被回收后
    被别的节点复用）。元素的 hash/== 才是按底层节点算的，稳定可用。
    这个坑实测踩到过：测试里 [False, False, False]，该保护的没保护上。
    """
    out, depth = set(), 0
    for node in el.iter():
        tag = node.tag.split("}")[-1]
        if tag == "fldChar":
            kind = node.get(f"{W}fldCharType")
            if kind == "begin":
                depth += 1
            elif kind == "end":
                depth = max(0, depth - 1)
        elif tag == "t" and (depth > 0 or _in_fldsimple(node, el)):
            out.add(node)
    return out


def _in_fldsimple(node, stop):
    p = node.getparent()
    while p is not None and p is not stop:
        if p.tag.split("}")[-1] == "fldSimple":
            return True
        p = p.getparent()
    return False


SPACE_CHARS = (" ", "\t")


def _edge_space(s, end=False):
    """s 的首字符（end=True 时取末字符）是不是空格。

    必须用元组成员判断，不能写成 `s[:1] in " \\t"`——空串是任何字符串的子串，
    `"" in " \\t"` 为真，会把空的 w:t 误判成"本来就带空格"，该补的
    xml:space 就没补上。测试抓到了这个。
    """
    return (s[-1:] if end else s[:1]) in SPACE_CHARS


def _set_text(node, text):
    """写回 w:t 的文字，必要时补 xml:space="preserve"。

    只在"新出现的首尾空格"上补——原本就有的（Word 本来就会吃掉）保持原样，
    免得顺手改变了别处的渲染结果。
    """
    old = node.text or ""
    node.text = text
    if (not _edge_space(old) and _edge_space(text)) or \
            (not _edge_space(old, end=True) and _edge_space(text, end=True)):
        node.set(XML_SPACE, "preserve")


# --------------------------------------------------------------------------
# 政策：与检查器共用（check_docx.py 从这里 import）
# --------------------------------------------------------------------------
CAPTION_RE = {
    "图": re.compile(r"^\s*图\s*\d+[.\-–]\d+"),
    "表": re.compile(r"^\s*表\s*\d+[.\-–]\d+"),
}


def caption_gap_pos(text, label):
    """题注的图号/表号后面缺空格吗？缺则返回号码最后一个字符的下标。

    末尾正好结束（如整段只有"图2.1"）也算缺——检查器一直这么报，
    修复时另行跳过，这样"报了多少"和"修了多少"对得上。
    """
    m = CAPTION_RE[label].match(text)
    if m is None:
        return None
    i = m.end()
    if i >= len(text) or not text[i].isspace():
        return i - 1
    return None


CJK_CH = re.compile(r"[一-鿿㐀-䶿]")
CJK_ANY = re.compile(r"[一-鿿㐀-䶿　-〿＀-￯]")
LATIN_ALPHA = re.compile(r"[A-Za-z]")
CITATION = re.compile(r"\[\s*\d+(?:\s*[，,、\-–—~至]\s*\d+)*\s*\]")
PAIR_OPEN = {"(": ")", "[": "]"}
# 方括号不自动改：全角方括号有 ［］(U+FF3B/FF3D) 与 【】(U+3010/3011) 两种写法，
# 语义不同（后者用于书名、强调），机械替换会改错意思。只提示，交给人。
NO_AUTOFIX = set("[]")


def punct_fixes(text):
    """半角标点分三堆，返回 (该改的, 要人看的, 豁免数)。

    "该改的"是一组组下标：括号对算一组（左右必须一起改，只改一边会变成
    "（区、市)" 这种半角全角混排，比不改更难看），逗号/分号/冒号各算一组。

    判据：两侧都是中文的标点用全角。括号内含拉丁字母的降级（"(RNA sequencing)"
    是英文缩写，括起来用半角才常见）；引用标注 [数字] 豁免（规范要求半角）。
    """
    spans = [m.span() for m in CITATION.finditer(text)]

    def exempt(i):
        return any(a <= i < b for a, b in spans)

    hi, lo, ex = [], [], 0
    for i, ch in enumerate(text):
        if ch not in PAIR_OPEN:
            continue
        j = text.find(PAIR_OPEN[ch], i + 1)
        if j < 0:
            continue
        if exempt(i) or exempt(j):
            ex += 1
            continue
        prev = text[i - 1] if i > 0 else ""
        nxt = text[j + 1] if j + 1 < len(text) else ""
        if not (prev and nxt and CJK_ANY.match(prev) and CJK_ANY.match(nxt)):
            continue
        if ch in NO_AUTOFIX:
            lo.append(i)
        elif LATIN_ALPHA.search(text[i + 1:j]):
            lo.append(i)              # 括号里是英文缩写，半角才常见
        else:
            hi.append([i, j])         # 成对记下，左右必须一起改

    colons = [i for i, ch in enumerate(text)
              if ch == ":" and 0 < i < len(text) - 1
              and CJK_CH.match(text[i - 1]) and CJK_CH.match(text[i + 1])]
    # 同段出现两个以上冒号，多半是"甲醛:乙酸:乙醇"这类比例式，整体降级
    ratio_like = len(colons) >= 2

    for i, ch in enumerate(text):
        if ch not in (",", ";", ":"):
            continue
        prev = text[i - 1] if i > 0 else ""
        nxt = text[i + 1] if i + 1 < len(text) else ""
        if not (prev and nxt and CJK_CH.match(prev) and CJK_CH.match(nxt)):
            continue
        if ch == ":" and ratio_like:
            lo.append(i)
        else:
            hi.append([i])
    return hi, lo, ex


# --------------------------------------------------------------------------
# 修复一：题注图号后缺空格
# --------------------------------------------------------------------------
def fix_caption_gaps(rows, rules, changes, notes):
    """在题注的图号/表号后插入半角空格，个数由规则决定（默认 1 个）。

    空格写进图号所在的那个 w:t（同节点则从中间插入，跨节点则挂到前一节点末尾），
    这样空格继承的是图号的字体（Times New Roman），中文名那侧不受影响。
    不新建 run——新建的 run 没有 rPr，会掉到样式默认字体上去。

    只补"一个空格都没有"的情况，图题表题各取自己那条 gap_after_number。
    已有空格但个数不对的不管（检查器也只报"缺空格"，不判个数）。
    """
    hits, skipped, at_end = {}, 0, 0

    for r in rows:
        if r.key not in ("figure_caption", "table_caption"):
            continue
        p = r.para
        label = "图" if r.key == "figure_caption" else "表"
        i = caption_gap_pos(p.text, label)
        if i is None:
            continue

        want = rules.get("styles", {}).get(r.key, {}).get("gap_after_number", 1)
        pad = " " * max(1, int(want or 1))

        text, spans = flatten(p.element)
        node, off = node_at(spans, i)
        if node is None or node in protected(p.element):
            skipped += 1
            continue

        cur = node.text or ""
        if off + 1 < len(cur):            # 图号和图名在同一个 w:t 里
            _set_text(node, cur[:off + 1] + pad + cur[off + 1:])
        elif i + 1 < len(text):           # 跨节点：挂到图号那个节点末尾
            _set_text(node, cur + pad)
        else:                             # 图号后面什么都没有
            at_end += 1
            continue
        hits.setdefault(pad, []).append(p.index)

    for pad in sorted(hits, key=len):
        changes.append(TextChange("题注图号后缺空格", "无空格",
                                  f"插入 {len(pad)} 个半角空格",
                                  len(hits[pad]), hits[pad]))
    if skipped:
        notes.append(f"跳过（题注落在引用域内）：{skipped} 处")
    if at_end:
        notes.append(f"跳过（图号后没有图名，无处插空格）：{at_end} 处")


# --------------------------------------------------------------------------
# 修复二：异常字符
# --------------------------------------------------------------------------
def _build_tables(cfg):
    """返回 (区间表, 同义字符表)。

    区间表来自规则的 bad 段，每项带 fix 处置：space / delete / report。
    处置写在 YAML 里而不是代码里——"哪些字符该删、哪些只报"是会变的政策，
    应该让改规则的人直接改得到，不用动代码。
    """
    ranges = []
    for it in cfg.get("bad", []):
        lo, hi = it["range"]
        ranges.append((lo, hi, it["name"], it.get("fix", "report")))
    syn = {}
    for s in cfg.get("synonyms", []):
        for ch in s["group"]:
            if ch != s["prefer"]:
                syn[ch] = (s["name"], s["prefer"])
    return ranges, syn


def fix_characters(doc, rules, changes, notes):
    """按黑名单清洗异常字符：换空间、删垃圾、统一同义字符。

    只处理 w:t 里的字。实测某样本 44 个可疑字符里 38 个落在 w:instrText
    （Zotero 的 JSON），照段落文本盲扫会把引用改废。
    """
    cfg = rules.get("characters", {})
    if not cfg.get("enabled"):
        return

    ranges, syn = _build_tables(cfg)
    counts, paras = {}, {}          # 键 (名称, 原字符, 结果) → 个数 / 段落集合
    skipped = 0

    for p in doc.paragraphs:
        prot = None
        for node in p.element.iter(f"{W}t"):
            old = node.text or ""
            if not old:
                continue
            new, found = _clean(old, ranges, syn)
            if not found:
                continue
            if prot is None:
                prot = protected(p.element)
            if node in prot:
                skipped += 1
                continue
            _set_text(node, new)
            for key, n in found.items():
                counts[key] = counts.get(key, 0) + n
                paras.setdefault(key, set()).add(p.index)

    for key in sorted(counts, key=lambda k: -counts[k]):
        name, old_ch, new_ch = key
        changes.append(TextChange(name, old_ch, new_ch, counts[key],
                                  sorted(paras[key])))

    if skipped:
        notes.append(f"跳过（落在引用域内的异常字符）：{skipped} 个")


def _clean(text, ranges, syn):
    """逐字处置，返回 (新文本, 记账)。

    记账的键就是给人看的三个字段 (名称, 原字符, 结果)，这样 fix_characters
    不用再翻译一遍。
    """
    out, found = [], {}
    for ch in text:
        cp = ord(ch)
        hit = next(((n, a) for lo, hi, n, a in ranges if lo <= cp <= hi), None)
        if hit is not None:
            name, action = hit
            if action == "space":
                out.append(" ")
                key = (name, f"U+{cp:04X}", "半角空格")
            elif action == "delete":
                key = (name, f"U+{cp:04X}", "删除")
            else:
                out.append(ch)       # report：只报不改（替换字符删了会丢信息）
                continue
        else:
            rep = syn.get(ch)
            if not rep:
                out.append(ch)
                continue
            out.append(rep[1])
            key = (rep[0], ch, rep[1])
        found[key] = found.get(key, 0) + 1
    return "".join(out), found


# --------------------------------------------------------------------------
# 修复三：半角标点全角化
# --------------------------------------------------------------------------
def fix_punctuation(doc, rules, changes, notes):
    """两侧都是中文的半角标点改成全角（高置信度那一类）。

    括号成对改：只改半边会变成"（区、市)"这种混排，比不改更糟。因此一组里
    只要有一个字符落在域内，整组都不动。
    """
    if not rules.get("punctuation", {}).get("enabled"):
        return

    counts, paras = {}, {}
    skipped = 0

    for p in doc.paragraphs:
        if not p.text.strip():
            continue
        groups, _, _ = punct_fixes(p.text)
        if not groups:
            continue

        text, spans = flatten(p.element)
        prot = protected(p.element)
        for group in groups:
            targets = []
            for idx in group:
                node, off = node_at(spans, idx)
                new_ch = FULLWIDTH.get(text[idx])
                if node is None or new_ch is None or node in prot:
                    targets = None          # 一组里有域内字符，整组放弃
                    break
                targets.append((node, off, idx, new_ch))

            if targets is None:
                skipped += 1
                continue

            # 同一节点里改多个字符，下标会互相影响：从后往前改
            for node, off, idx, new_ch in sorted(targets, key=lambda t: -t[1]):
                cur = node.text or ""
                _set_text(node, cur[:off] + new_ch + cur[off + 1:])
                key = (text[idx], new_ch)
                counts[key] = counts.get(key, 0) + 1
                paras.setdefault(key, set()).add(p.index)

    for key in sorted(counts, key=lambda k: -counts[k]):
        old_ch, new_ch = key
        changes.append(TextChange(f"半角「{old_ch}」全角化", old_ch, new_ch,
                                  counts[key], sorted(paras[key])))

    if skipped:
        notes.append(f"跳过（标点落在引用域内）：{skipped} 组")


# --------------------------------------------------------------------------
def fix_text(doc, rows, rules, changes, notes):
    """文本层修复的总入口。顺序无所谓，三者互不重叠。"""
    fix_caption_gaps(rows, rules, changes, notes)
    fix_characters(doc, rules, changes, notes)
    fix_punctuation(doc, rules, changes, notes)
