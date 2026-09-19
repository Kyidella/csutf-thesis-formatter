"""对照实验：Zotero 刷新会不会覆盖我们写进去的格式。

这是目前唯一还没实测过的安全假设。前面每条纪律都做过对照实验，这条没有。

做法：往三类位置写一个**一眼可见**的格式（红色 + 18pt），让用户在 Zotero 里
点刷新，看哪些变回去了：

  A. Zotero 书目域内的参考文献条目   ← 会被域管理器重新生成
  B. 正文里的正常段落                ← 对照组，不该变
  C. 含引用标注的段落                ← 引用域所在的段落

如果 A 变回去、B 不变，说明"域生成的内容"确实会被覆盖，
格式化器就应该对这类段落只报告不修改。

用法:
    python zotero_refresh_test.py input.docx -o out/刷新实验.docx
"""

import argparse
import copy
import sys
import zipfile
from pathlib import Path

from lxml import etree

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
DOCUMENT = "word/document.xml"


def _has_rpr_run(p):
    rpr = p.find(f"{W}rPr")
    return rpr is not None


def _set_visible_mark(run):
    """把 run 标成红色 22pt —— 肉眼一眼能看出有没有被覆盖。"""
    rpr = run.find(f"{W}rPr")
    if rpr is None:
        rpr = run.makeelement(f"{W}rPr", {})
        run.insert(0, rpr)
    for tag, attrs in (("color", {"val": "FF0000"}),
                       ("sz", {"val": "44"}),        # 22pt
                       ("szCs", {"val": "44"})):
        node = rpr.find(f"{W}{tag}")
        if node is None:
            node = rpr.makeelement(f"{W}{tag}", {})
            rpr.append(node)
        for k, v in attrs.items():
            node.set(f"{W}{k}", v)


def classify_paragraphs(body):
    """按文档顺序给每段打标：在不在 Zotero 书目的域区间内 / 在不在引用域内。"""
    out = []
    depth = 0
    bibl_start = None
    current_field = None

    for p in body.iter(f"{W}p"):
        text = "".join(t.text or "" for t in p.iter(f"{W}t"))
        marks = set()
        for r in p.iter(f"{W}r"):
            for child in r:
                tag = etree.QName(child).localname
                if tag == "fldChar":
                    t = child.get(f"{W}fldCharType")
                    if t == "begin":
                        depth += 1
                        current_field = None     # 等下一条 instrText 才知道是什么
                        bibl_start = depth
                    elif t == "end":
                        if current_field == "bibl":
                            current_field = None
                            bibl_start = None
                        depth = max(0, depth - 1)
                elif tag == "instrText":
                    txt = child.text or ""
                    if "ZOTERO_BIBL" in txt:
                        current_field = "bibl"
                    elif "ZOTERO_ITEM" in txt:
                        if current_field != "bibl":
                            marks.add("citation")
        if bibl_start is not None and current_field == "bibl":
            marks.add("bibliography")
        out.append((p, text, marks))
    return out


def apply_mark(paragraph, kinds, stat):
    for run in paragraph.findall(f"{W}r"):
        if run.find(f"{W}t") is None:
            continue
        _set_visible_mark(run)
        for k in kinds:
            stat[k] = stat.get(k, 0) + 1


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("input")
    ap.add_argument("-o", "--output", required=True)
    args = ap.parse_args()

    z = zipfile.ZipFile(args.input)
    root = etree.fromstring(z.read(DOCUMENT))
    body = root.find(f"{W}body")

    paras = classify_paragraphs(body)

    # A. 书目域内的条目
    stat = {}
    a_count = 0
    for p, text, marks in paras:
        if "bibliography" in marks and text.strip():
            apply_mark(p, ["bibl"], stat)
            a_count += 1

    # B. 对照组：正文正常段落（不含任何域），取 3 段
    b_count = 0
    for p, text, marks in paras:
        if marks or not text.strip():
            continue
        if len(text.strip()) < 60:
            continue
        apply_mark(p, ["body"], stat)
        b_count += 1
        if b_count >= 3:
            break

    # C. 含引用标注的段落，取 3 段
    c_count = 0
    for p, text, marks in paras:
        if "citation" in marks and "bibliography" not in marks and text.strip():
            apply_mark(p, ["cite"], stat)
            c_count += 1
            if c_count >= 3:
                break

    out = Path(args.output)
    out.parent.mkdir(parents=True, exist_ok=True)
    new_xml = etree.tostring(root, xml_declaration=True,
                             encoding="UTF-8", standalone=True)
    with zipfile.ZipFile(args.input) as zin, \
            zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = new_xml if item.filename == DOCUMENT else zin.read(item.filename)
            zout.writestr(item, data)

    print(f"已写出 {out}\n")
    print(f"  A. 书目域内条目  {a_count:>4} 段  —— 会被 Zotero 重新生成")
    print(f"  B. 正文对照组    {b_count:>4} 段  —— 不该变")
    print(f"  C. 含引用标注段  {c_count:>4} 段  —— 引用域所在段落")
    print("\n下一步（需要你在 WPS + Zotero 里操作）：")
    print("  1. 用 WPS 打开这个文件，确认 A/B/C 三处都变成了红色 22pt")
    print("  2. 在 Zotero 里点刷新（或随便改一条引用再刷新）")
    print("  3. 看哪些位置的红色消失了：")
    print("       A 消失、B 还在  -> 域生成的内容会被覆盖，格式化器应只报告不改")
    print("       三个都还在      -> 我们的格式不会被覆盖，可以照常处理")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
