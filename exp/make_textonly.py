"""只跑文本层，产出一个用于真机验证的孤立文件。

为什么要孤立：要给用户验证的是"文本层会不会动坏 Zotero 引用"。
如果同时把页边距、样式也都改了，万一出问题就分不清是哪一层的锅。
所以这个文件**只替换 word/document.xml**，其余 70 个部件逐字节照抄。

用法:
    python exp/make_textonly.py 输入.docx 输出.docx
"""

import sys
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
sys.stdout.reconfigure(encoding="utf-8")

import yaml
from lxml import etree

from classifier import classify
from docmodel import load
from textfix import fix_text

ROOT = Path(__file__).resolve().parent.parent
DOCUMENT = "word/document.xml"

src, dst = sys.argv[1], sys.argv[2]
rules = yaml.safe_load(open(ROOT / "rules" / "中南林科大_硕士_学术学位_理工类.yaml",
                            encoding="utf-8"))

doc = load(src)
changes, notes = [], []
fix_text(doc, classify(doc), rules, changes, notes)

new_doc = etree.tostring(doc.root, xml_declaration=True,
                         encoding="UTF-8", standalone=True)

changed = []
with zipfile.ZipFile(src) as zin, \
        zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename == DOCUMENT:
            changed.append(item.filename)
            data = new_doc
        zout.writestr(item, data)

print(f"# 只改文本层：{src} → {dst}")
for c in changes:
    print(f"- {c.render()}")
for n in notes:
    print(f"- {n}")
print(f"\n被替换的部件：{changed}（其余全部逐字节照抄）")
