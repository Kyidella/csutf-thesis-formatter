"""位置感知的段落分类器。

纯文本特征分不出「摘要正文」和「正文」——它们长得一模一样。所以要先用章节
锚点（摘要 / ABSTRACT / 目录 / 参考文献 / 附录 / 致谢 / 章标题）把文档切成区段，
再让区段内的普通文字按「当前所在章节」归属类型。

层次（自上而下，先命中先返回）：
  1. 明确的文本标记（图题、表题、关键词、引用条目……）
  2. 标题（大纲级别 → 样式名 → 编号模式，逐个回退）
  3. 所在章节的默认正文类型（摘要正文 / 英文摘要正文 / 参考文献条目 / 致谢正文…）

用法:
    from classifier import classify
    rows = classify(doc)          # doc 来自 docmodel.load
"""

import re
from dataclasses import dataclass

# ---- 章节锚点：命中即开启一个新的区段 ----
ANCHORS = [
    (re.compile(r"^摘\s*要$"), "abstract_title", "abstract_body"),
    (re.compile(r"^ABSTRACT$", re.I), "abstract_en_title", "abstract_en_body"),
    (re.compile(r"^目\s*录$"), "toc_title", "toc_entry_1"),
    (re.compile(r"^参考文献$"), "reference_title", "reference_item"),
    (re.compile(r"^附录\s*[A-Z]?"), "appendix_title", "appendix_body"),
    (re.compile(r"^致\s*谢$"), "ack_title", "ack_body"),
]

# ---- 与位置无关的文本标记 ----
TEXT_RULES = [
    (re.compile(r"^\s*续\s*表"), "table_continued"),
    (re.compile(r"^\s*表\s*\d+[.\-–]\d+"), "table_caption"),
    (re.compile(r"^\s*图\s*\d+[.\-–]\d+"), "figure_caption"),
    (re.compile(r"^\s*\[\s*\d+"), "reference_item"),
    (re.compile(r"^关键词"), "abstract_keywords"),
    (re.compile(r"^Keywords\s*:", re.I), "abstract_en_keywords"),
    (re.compile(r"^数据来源"), "table_note"),
    (re.compile(r"^(图片来源|注：)"), "figure_note"),
]

# 标题候选的排除条件：太长、或带句末/句中标点，就不是标题
NOT_HEADING_TAIL = re.compile(r"[。！？；，,;]$")
HEADING_MAX_LEN = 45

# 编号模式 → 层级
NUM_PATTERNS = [
    (re.compile(r"^第\s*[一二三四五六七八九十百]+\s*章"), 1),
    (re.compile(r"^第\s*\d+\s*章"), 1),
    (re.compile(r"^(\d+)\.(\d+)\.(\d+)\.(\d+)(?=\D|$)"), 4),
    (re.compile(r"^(\d+)\.(\d+)\.(\d+)(?=\D|$)"), 3),
    (re.compile(r"^(\d+)\.(\d+)(?=\D|$)"), 2),
    (re.compile(r"^(\d+)(?=\D|$)"), 1),
]

# 纯阿拉伯数字开头的一级标题（如 "1  绪论"）必须很短，
# 否则 "498 Shaoshan South Road, Tianxin District" 这类封面地址也会中招。
ARABIC_L1_MAX_LEN = 30

STYLE_HEADING = re.compile(r"heading\s*(\d)", re.I)
LONE_HEADINGS = {"摘要", "abstract", "目录", "参考文献", "致谢", "结论"}


@dataclass
class Classified:
    index: int
    key: str
    level: int | None
    reason: str
    para: object


def _norm(s):
    return re.sub(r"\s+", "", s)


LIST_MARKER = re.compile(r"^\d+\s*[）)、]\s*")
DATE_LEAD = re.compile(r"^\d{4}\s*[年\-/.]")


def heading_level(p):
    """逐级回退地判断标题层级：大纲级别 → 样式名 → 编号模式。"""
    t = p.text.strip()
    if not t or len(t) > HEADING_MAX_LEN or NOT_HEADING_TAIL.search(t):
        return None, None
    if LIST_MARKER.match(t):
        return None, None          # "1）参加项目课题" 是列表项
    if DATE_LEAD.match(t):
        return None, None          # "2025年6月" 是日期

    if p.outline_level is not None and p.outline_level <= 3:
        return p.outline_level + 1, "大纲级别"

    if p.style_name:
        m = STYLE_HEADING.search(p.style_name)
        if m:
            return int(m.group(1)), "样式名"

    if _norm(t).lower() in LONE_HEADINGS:
        return 1, "固定章名"

    for pat, lvl in NUM_PATTERNS:
        if pat.match(t):
            if lvl == 1 and pat.pattern.startswith(r"^(\d+)(") \
                    and len(t) > ARABIC_L1_MAX_LEN:
                return None, None
            return lvl, "编号模式"
    return None, None


def is_toc_entry(p):
    """目录条目：样式名含 toc/目录，或带连续点前导符。"""
    s = (p.style_name or "").lower()
    if "toc" in s or "目录" in s:
        return True
    return bool(re.search(r"[.·]{3,}", p.text))


def classify(doc):
    """返回 list[Classified]，顺序与 doc.paragraphs 一致。"""
    rows = []
    ctx_key = None          # 当前区段的默认正文类型
    in_toc = False          # 是否处在目录区

    # 首个章节锚点之前是封面（含题名页、原创性声明等）。
    # 样本里封面段落被手动设过大纲级别 0，不挡住就会误判成一级标题。
    first_anchor = None
    for p in doc.paragraphs:
        t = p.text.strip()
        if t and len(t) <= 20 and any(
                pat.match(_norm(t)) for pat, _, _ in ANCHORS):
            first_anchor = p.index
            break

    for p in doc.paragraphs:
        t = p.text.strip()

        # --- 0a. 封面区 ---
        if first_anchor is not None and p.index < first_anchor:
            rows.append(Classified(p.index, "cover", None, "封面区", p))
            continue

        # --- 0b. 表格内的段落不做标题判定：单元格里 "1.5(μL)" 之类会大量误判 ---
        if p.in_table:
            rows.append(Classified(p.index, "table_content", None, "表格内容", p))
            continue

        # --- 1. 章节锚点 ---
        # 目录区里的 "附录A …… 100" 这类行不能当锚点，否则会打断目录区
        if t and len(t) <= 20 and not in_toc:
            hit = next(((pat, k, bk) for pat, k, bk in ANCHORS
                        if pat.match(_norm(t))), None)
            if hit:
                _, key, body_key = hit
                ctx_key = body_key
                in_toc = key == "toc_title"
                rows.append(Classified(p.index, key,
                                       1 if key != "toc_title" else None,
                                       f"章节锚点 {key}", p))
                continue

        # --- 2. 目录区：SDT 控件内的条目、带点前导符的、或样式名像 toc 的 ---
        if in_toc:
            if p.in_sdt or is_toc_entry(p):
                s = (p.style_name or "").lower()
                m = re.search(r"toc\s*(\d)", s)
                key = f"toc_entry_{m.group(1)}" if m else \
                    ("toc_entry_2" if re.search(r"[.·]{3,}", p.text) else "toc_entry_1")
                rows.append(Classified(p.index, key, None, "目录区", p))
                continue
            in_toc = False        # 遇到目录之外的内容，目录区结束

        # --- 3. 与位置无关的文本标记 ---
        hit = next((k for pat, k in TEXT_RULES if pat.match(t)), None)
        if hit:
            rows.append(Classified(p.index, hit, None, f"文本标记 {hit}", p))
            continue

        # --- 4. 标题 ---
        lvl, why = heading_level(p)
        if lvl:
            # 只有一级标题才代表进入新的章节；二/三级是小节，不该重置上下文
            if lvl == 1:
                ctx_key = None
            rows.append(Classified(p.index, f"heading{lvl}", lvl, f"标题（{why}）", p))
            continue

        # --- 5. 按所在章节归属 ---
        key = ctx_key or "body"
        rows.append(Classified(p.index, key, None, f"章节上下文({ctx_key or '正文'})", p))

    return rows


def summarize(rows):
    """统计各类型的段落数，便于核对分类是否合理。"""
    from collections import Counter
    c = Counter(r.key for r in rows)
    reasons = Counter(r.reason.split()[0] for r in rows)
    return c, reasons


if __name__ == "__main__":
    import sys
    sys.stdout.reconfigure(encoding="utf-8")
    sys.path.insert(0, ".")
    from docmodel import load

    doc = load(sys.argv[1])
    rows = classify(doc)
    counts, reasons = summarize(rows)
    print("按类型:")
    for k, n in counts.most_common():
        print(f"  {n:5d}  {k}")
    print("\n按判定依据:")
    for k, n in reasons.most_common():
        print(f"  {n:5d}  {k}")
