"""侦察：文本层问题的真实形态。

要回答三个问题，答案决定 textfix 怎么写：
  1. 图题缺空格的那些段，图号和图名是不是在同一个 w:t 里？
  2. 异常字符落单还是成串？
  3. 高置信度标点所在段，有没有含引用域？

用法:
    python exp/probe_textfix.py 论文.docx
"""

import re
import sys
sys.stdout.reconfigure(encoding='utf-8')
from collections import Counter
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from docmodel import W, load
from check_docx import classify_punct
from classifier import classify
import yaml

import json

ROOT = Path(__file__).resolve().parent.parent
rules = yaml.safe_load(open(ROOT / "rules" / "中南林科大_硕士_学术学位_理工类.yaml",
                            encoding="utf-8"))

doc = load(sys.argv[1])
rows = classify(doc)

# ---- 文本扁平化：复刻 docmodel._text_of 的顺序，给出 index → 节点 的映射 ----
def flatten(el):
    """返回 (text, nodes)，nodes 形如 [(kind, node, start, end)]。"""
    parts, nodes = [], []
    pos = 0
    for node in el.iter():
        t = node.tag.split("}")[-1]
        s = None
        if t == "t":
            s = node.text or ""
        elif t == "instrText":
            s = node.text or ""
        elif t == "tab":
            s = "\t"
        elif t in ("br", "cr"):
            s = "\n"
        if s is None:
            continue
        kind = "instr" if t == "instrText" else ("t" if t == "t" else t)
        nodes.append((kind, node, pos, pos + len(s)))
        parts.append(s)
        pos += len(s)
    return "".join(parts), nodes


def node_at(nodes, i):
    for kind, node, a, b in nodes:
        if a <= i < b:
            return kind, node, i - a
    return None, None, None


print("=" * 70)
print("1. 图题/表题缺空格")
print("=" * 70)
PAT = re.compile(r"^\s*(图|表)\s*\d+[.\-–]\d+(?!\s)")
n_missing = 0
struct = Counter()
in_field = 0
samples = []
for r in rows:
    if r.key not in ("figure_caption", "table_caption"):
        continue
    t = r.para.text.strip()
    m = PAT.match(t)
    if not m:
        continue
    n_missing += 1
    if r.para.has_field:
        in_field += 1
    text, nodes = flatten(r.para.element)
    # 图号最后一个字符的位置
    i = m.end() - 1
    kind, node, off = node_at(nodes, i)
    after = node_at(nodes, i + 1)
    struct[(kind, after[0] if after else "PARA_END",
            node is after[1] if after else None)] += 1
    if len(samples) < 8:
        samples.append((r.para.index, t[:30], r.para.has_field,
                        [(k, (n.text or "")[:20]) for k, n, _, _ in nodes[:6]]))

print(f"缺空格 {n_missing} 处，其中含引用域的段落 {in_field} 处")
print(f"图号末字符所在节点 / 其后字符所在节点 / 是否同一节点：")
for k, v in struct.most_common():
    print(f"   {k} × {v}")
for s in samples:
    print(f"   段{s[0]} field={s[2]} {s[1]!r}")
    print(f"        {s[3]}")

print()
print("=" * 70)
print("2. 异常字符")
print("=" * 70)
ranges = rules["characters"]["bad"]
def bad_name(cp):
    for it in ranges:
        lo, hi = it["range"]
        if lo <= cp <= hi:
            return it["name"]
    return None

cnt = Counter()
where = Counter()
for p in doc.paragraphs:
    text, nodes = flatten(p.element)
    for i, ch in enumerate(text):
        nm = bad_name(ord(ch))
        if not nm:
            continue
        cnt[(nm, hex(ord(ch)))] += 1
        kind, node, off = node_at(nodes, i)
        where[kind] += 1
for k, v in cnt.most_common():
    print(f"   {k[0]:<16} {k[1]:<8} × {v}")
print(f"   所在节点类型：{dict(where)}")

print()
print("=" * 70)
print("3. 高置信度半角标点")
print("=" * 70)
pchar = Counter()
field_para = 0
total = 0
for p in doc.paragraphs:
    if not p.text.strip():
        continue
    hi, lo, ex = classify_punct(p.text)
    if not hi:
        continue
    total += len(hi)
    if p.has_field:
        field_para += 1
    for snip in hi:
        for ch in re.findall(r"[(),;]", snip):
            pchar[ch] += 1
print(f"   高置信度 {total} 处（粗计），涉及含引用域段落 {field_para} 段")
print(f"   字符分布：{dict(pchar)}")

print()
print("=" * 70)
print("4. w:t 与 w:instrText 的总量对照")
print("=" * 70)
n_t = n_instr = 0
for el in doc.root.iter():
    tag = el.tag.split("}")[-1]
    if tag == "t":
        n_t += 1
    elif tag == "instrText":
        n_instr += 1
print(f"   w:t {n_t} 个，w:instrText {n_instr} 个")
