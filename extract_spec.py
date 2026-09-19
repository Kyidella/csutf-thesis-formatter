"""从学校模板的 Word 批注里抽出格式规范原文，并定位每条批注锚定的正文位置。

用法:
    python extract_spec.py templates/xxx.docx -o spec_raw.json
"""

import argparse
import json
import sys
import zipfile

from lxml import etree

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
WNS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": WNS}


def para_text(p):
    """把一个 w:p 里的所有 w:t 直接拼起来（中文常被拆进多个 run）。"""
    parts = []
    for node in p.iter():
        tag = etree.QName(node).localname
        if tag == "t":
            parts.append(node.text or "")
        elif tag in ("tab",):
            parts.append("\t")
        elif tag in ("br", "cr"):
            parts.append("\n")
    return "".join(parts)


def extract_comments(z):
    if "word/comments.xml" not in z.namelist():
        return []
    root = etree.fromstring(z.read("word/comments.xml"))
    out = []
    for c in root.findall(f"{W}comment"):
        paras = [para_text(p) for p in c.findall(f"{W}p")]
        while paras and not paras[-1].strip():
            paras.pop()
        out.append({
            "id": int(c.get(f"{W}id")),
            "author": c.get(f"{W}author"),
            "date": c.get(f"{W}date"),
            "text": "\n".join(paras),
        })
    return out


def anchor_map(z):
    """comment id -> 锚定的段落序号与段落文本。"""
    root = etree.fromstring(z.read("word/document.xml"))
    body = root.find(f"{W}body")

    # 按文档顺序遍历顶层段落，记录 commentRangeStart 落在哪个段落
    anchors = {}
    for idx, p in enumerate(body.iter(f"{W}p")):
        for node in p.iter():
            if etree.QName(node).localname == "commentRangeStart":
                cid = int(node.get(f"{W}id"))
                anchors.setdefault(cid, {"para_index": None, "para_text": None})
                if anchors[cid]["para_index"] is None:
                    anchors[cid]["para_index"] = idx
                    anchors[cid]["para_text"] = para_text(p)[:60]
            elif etree.QName(node).localname == "commentReference":
                cid = int(node.get(f"{W}id"))
                anchors.setdefault(cid, {"para_index": None, "para_text": None})
                if anchors[cid]["para_index"] is None:
                    anchors[cid]["para_index"] = idx
                    anchors[cid]["para_text"] = para_text(p)[:60]
    return anchors


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("input")
    ap.add_argument("-o", "--output", default="spec_raw.json")
    args = ap.parse_args()

    z = zipfile.ZipFile(args.input)
    comments = extract_comments(z)
    anchors = anchor_map(z)

    for c in comments:
        a = anchors.get(c["id"], {})
        c["anchor_para_index"] = a.get("para_index")
        c["anchor_para_text"] = a.get("para_text")

    with open(args.output, "w", encoding="utf-8") as f:
        json.dump(comments, f, ensure_ascii=False, indent=1)

    print(f"共 {len(comments)} 条批注，已写入 {args.output}\n")
    for c in comments:
        t = c["text"].replace("\n", " ⏎ ")
        print(f"--- id={c['id']} 锚定段落#{c['anchor_para_index']} "
              f"锚定文本={c['anchor_para_text']!r}")
        print(f"    {t}")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
