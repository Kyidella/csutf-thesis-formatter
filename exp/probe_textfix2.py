"""侦察 2：把修复所需的细节问清楚。

  1. w:t 里的异常字符，具体落在哪些段，那些段有没有引用域？
  2. 同义字符（µ/μ）有没有出现？分布在哪种节点？
  3. 高置信度标点的精确索引、字符，所在段有没有引用域？
  4. 现有 w:t 里的首尾空格是怎么处理的（有没有 xml:space="preserve"）？

用法:
    python exp/probe_textfix2.py 论文.docx
"""

import re
import sys
from collections import Counter
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
sys.stdout.reconfigure(encoding="utf-8")

import yaml

from check_docx import classify_punct
from classifier import classify
from docmodel import load

XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
ROOT = Path(__file__).resolve().parent.parent
rules = yaml.safe_load(open(ROOT / "rules" / "中南林科大_硕士_学术学位_理工类.yaml",
                            encoding="utf-8"))

doc = load(sys.argv[1])
rows = classify(doc)


def flatten(el):
    """复刻 docmodel._text_of 的顺序（只算 w:t / tab / br），返回 index 映射。"""
    parts, nodes, pos = [], [], 0
    for node in el.iter():
        t = node.tag.split("}")[-1]
        s = {"t": lambda: node.text or "",
             "tab": lambda: "\t",
             "br": lambda: "\n",
             "cr": lambda: "\n"}.get(t, lambda: None)()
        if s is None:
            continue
        nodes.append((node, pos, pos + len(s)))
        parts.append(s)
        pos += len(s)
    return "".join(parts), nodes


def node_at(nodes, i):
    for node, a, b in nodes:
        if a <= i < b:
            return node, i - a
    return None, None


print("=" * 72)
print("1. w:t 里的异常字符（w:instrText 里的不列）")
print("=" * 72)
ranges = rules["characters"]["bad"]
def bad_name(cp):
    for it in ranges:
        lo, hi = it["range"]
        if lo <= cp <= hi:
            return it["name"]
    return None

n_text = n_instr = 0
for p in doc.paragraphs:
    for node in p.element.iter("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t"):
        for ch in (node.text or ""):
            nm = bad_name(ord(ch))
            if nm:
                n_text += 1
                print(f"   段{p.index} field={p.has_field} U+{ord(ch):04X} {nm} "
                      f"节点文本={node.text!r}")
    for node in p.element.iter("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}instrText"):
        n_instr += sum(1 for ch in (node.text or "") if bad_name(ord(ch)))
print(f"   —— w:t 共 {n_text} 个，w:instrText 共 {n_instr} 个")

print()
print("=" * 72)
print("2. 同义字符 µ / μ")
print("=" * 72)
syn = rules["characters"]["synonyms"][0]
found = Counter()
for p in doc.paragraphs:
    for node in p.element.iter():
        t = node.tag.split("}")[-1]
        if t not in ("t", "instrText"):
            continue
        for ch in (node.text or ""):
            if ch in syn["group"]:
                found[(ch, t, p.has_field)] += 1
for k, v in found.most_common():
    print(f"   {k[0]!r} U+{ord(k[0]):04X} 节点={k[1]} field段={k[2]} × {v}")
if not found:
    print("   （无）")

print()
print("=" * 72)
print("3. 高置信度半角标点：精确索引")
print("=" * 72)
FIXABLE = {"(": "（", ")": "）", ",": "，", ";": "；"}
for p in doc.paragraphs:
    if not p.text.strip():
        continue
    hi, lo, ex = classify_punct(p.text)
    if not hi:
        continue
    text, nodes = flatten(p.element)
    for snip in hi:
        pos = text.find(snip)
        chars = []
        for k, ch in enumerate(snip):
            if ch in "(),;[]":
                gi = pos + k
                node, off = node_at(nodes, gi)
                chars.append((ch, FIXABLE.get(ch, "—不可自动修—"), node is not None))
        print(f"   段{p.index} field={p.has_field} {snip!r}")
        print(f"        {chars}")

print()
print("=" * 72)
print("4. w:t 的首尾空格 / xml:space")
print("=" * 72)
lead = tail = 0
no_preserve = []
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
for el in doc.root.iter(f"{W}t"):
    s = el.text or ""
    if s != s.strip():
        if s[:1].isspace():
            lead += 1
        if s[-1:].isspace():
            tail += 1
        if not el.get(XML_SPACE):
            no_preserve.append(s)
print(f"   首空格 {lead} 个 · 尾空格 {tail} 个 · 其中未标 xml:space=preserve 的 {len(no_preserve)} 个")
for s in no_preserve[:5]:
    print(f"      {s!r}")
