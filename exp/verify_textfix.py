"""对照实验：文本层修复到底动了什么。

文本层是唯一"改字"的一层，所以验的东西和前面几层不同：
  1. w:instrText 逐字节不变（Zotero 的 JSON）
  2. 域标记（fldChar）数量与顺序不变
  3. run 数量、w:t 节点数量不变（没有新增/合并/删除）
  4. 除 document.xml / styles.xml 外，zip 里其余部件哈希全不变
  5. 正文文字的变化**只有**预期的那些（逐段比对，列出全部差异）

用法:
    python exp/verify_textfix.py 原文件.docx 修正后.docx
"""

import hashlib
import sys
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
sys.stdout.reconfigure(encoding="utf-8")

from lxml import etree

from docmodel import W, load

XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"


def parts(path):
    with zipfile.ZipFile(path) as z:
        return {n: z.read(n) for n in z.namelist()}


def digest(b):
    return hashlib.sha256(b).hexdigest()[:12]


def collect(root):
    instr = [n.text or "" for n in root.iter(f"{W}instrText")]
    fld = [(n.get(f"{W}fldCharType"), n.get(f"{W}instr"))
           for n in root.iter(f"{W}fldChar")]
    runs = list(root.iter(f"{W}r"))
    ts = list(root.iter(f"{W}t"))
    return instr, fld, runs, ts


old_path, new_path = sys.argv[1], sys.argv[2]
a, b = parts(old_path), parts(new_path)

print("=" * 72)
print("1. zip 部件")
print("=" * 72)
only_a = set(a) - set(b)
only_b = set(b) - set(a)
if only_a or only_b:
    print(f"   ！！部件增删：多的={only_a} 少的={only_b}")
changed = [n for n in a if n in b and a[n] != b[n]]
print(f"   部件数 {len(a)} → {len(b)}；内容变化的：{changed}")

print()
print("=" * 72)
print("2. 域的完整性")
print("=" * 72)
ra = etree.fromstring(a["word/document.xml"])
rb = etree.fromstring(b["word/document.xml"])
ia, fa, run_a, ta = collect(ra)
ib, fb, run_b, tb = collect(rb)

print(f"   w:instrText {len(ia)} → {len(ib)}")
same = sum(1 for x, y in zip(ia, ib) if x == y)
print(f"   逐条相同的 {same}/{len(ia)}"
      + ("  ✓ Zotero 的 JSON 一个字节没动" if same == len(ia) else "  ！！有改动"))
if same != len(ia):
    for k, (x, y) in enumerate(zip(ia, ib)):
        if x != y:
            print(f"      第{k}条变化：{x[:60]!r} → {y[:60]!r}")
print(f"   fldChar 标记 {len(fa)} → {len(fb)}；序列一致：{fa == fb}")
print(f"   w:r 节点 {len(run_a)} → {len(run_b)}；数量一致：{len(run_a) == len(run_b)}")
print(f"   w:t 节点 {len(ta)} → {len(tb)}；数量一致：{len(ta) == len(tb)}")

print()
print("=" * 72)
print("3. 正文文字的变化（应当只有预期的那些）")
print("=" * 72)
def diff_window(x, y, half=22):
    """截出两者首个差异附近的窗口，否则长段落里改的那个字看不见。"""
    i = 0
    while i < min(len(x), len(y)) and x[i] == y[i]:
        i += 1
    lo = max(0, i - half)
    return x[lo:i + half], y[lo:i + half]


da, db = load(old_path), load(new_path)
assert len(da.paragraphs) == len(db.paragraphs), "段落数变了，不该发生"
n_diff = 0
for pa, pb in zip(da.paragraphs, db.paragraphs):
    if pa.text == pb.text:
        continue
    n_diff += 1
    x, y = diff_window(pa.text, pb.text)
    print(f"   段{pa.index}: {x!r}")
    print(f"          → {y!r}")
print(f"   —— 共 {n_diff} 段文字有变化（段落总数 {len(da.paragraphs)}）")

print()
print("=" * 72)
print("4. 新增 xml:space=\"preserve\" 的节点")
print("=" * 72)
la = list(ra.iter(f"{W}t"))
lb = list(rb.iter(f"{W}t"))
gained = [(k, a_t, b_t) for k, (a_t, b_t) in enumerate(zip(la, lb))
          if b_t.get(XML_SPACE) and not a_t.get(XML_SPACE)]
lost = [k for k, (a_t, b_t) in enumerate(zip(la, lb))
        if a_t.get(XML_SPACE) and not b_t.get(XML_SPACE)]
for k, a_t, b_t in gained:
    print(f"   w:t #{k}: {(a_t.text or '')[:36]!r}")
    print(f"          → {(b_t.text or '')[:36]!r}")
print(f"   —— 新增 {len(gained)} 个，去掉 {len(lost)} 个"
      f"（改前带该属性的共 {sum(1 for n in la if n.get(XML_SPACE))} 个）")
