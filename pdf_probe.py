"""从 PDF（渲染结果）反向测量排版参数，用于与规范/其他样本交叉验证。

PDF 里没有样式名和大纲级别，只有字形、字号、坐标。可测量的是：
  页面尺寸 / 页边距 / 各级字号 / 中西文字体 / 行距 / 对齐
用法:
    python pdf_probe.py input.pdf
"""

import argparse
import collections
import statistics
import sys

import pymupdf

PT_PER_CM = 72 / 2.54

FONT_CN = {
    "simsun": "宋体", "nsimsun": "新宋体", "simhei": "黑体", "simkai": "楷体",
    "kaiti": "楷体", "fangsong": "仿宋", "msyh": "微软雅黑", "dengxian": "等线",
    "stsong": "宋体", "sthei": "黑体", "stkaiti": "楷体", "huawenxingkai": "华文行楷",
}


def cn_font(name):
    base = name.split("+")[-1].lower()
    for k, v in FONT_CN.items():
        if k in base:
            return v
    return name


def is_en(name):
    b = name.lower()
    return "times" in b or "arial" in b or "cambria" in b or "calibri" in b


def repeated_text(doc, min_pages_ratio=0.3):
    """找出在多数页面上重复出现的行文本 —— 即页眉页脚。"""
    per_text = collections.defaultdict(set)
    for pno, page in enumerate(doc):
        for blk in page.get_text("dict")["blocks"]:
            if blk.get("type") != 0:
                continue
            for line in blk["lines"]:
                t = "".join(s["text"] for s in line["spans"]).strip()
                if t:
                    per_text[t].add(pno)
    threshold = max(3, int(doc.page_count * min_pages_ratio))
    return {t for t, pages in per_text.items() if len(pages) >= threshold}


PAGE_NUM = __import__("re").compile(r"^[\dIVXivx]+$")


def page_lines(page, repeated):
    """把一页的 span 按纵向位置聚成行，剔除页眉页脚与页码，返回每行的 y / 左右缘。

    直接用 PyMuPDF 的 block 分组不行：这类 PDF 常常一块只有一行，
    组内求相邻行差会得到空结果。
    """
    h = page.rect.height
    spans = []
    for blk in page.get_text("dict")["blocks"]:
        if blk.get("type") != 0:
            continue
        for line in blk["lines"]:
            t = "".join(s["text"] for s in line["spans"]).strip()
            if not t or t in repeated:
                continue
            y = line["bbox"][1]
            # 页面上下 12% 带内的纯数字行 = 页码
            if PAGE_NUM.match(t) and (y < 0.12 * h or y > 0.88 * h):
                continue
            for s in line["spans"]:
                if s["text"].strip():
                    spans.append(s)

    spans.sort(key=lambda s: (s["bbox"][1], s["bbox"][0]))
    lines = []
    for s in spans:
        y = s["bbox"][1]
        if lines and abs(y - lines[-1]["y"]) < 3:
            cur = lines[-1]
            cur["x0"] = min(cur["x0"], s["bbox"][0])
            cur["x1"] = max(cur["x1"], s["bbox"][2])
        else:
            lines.append({"y": y, "y1": s["bbox"][3],
                          "x0": s["bbox"][0], "x1": s["bbox"][2]})
        lines[-1]["y1"] = max(lines[-1].get("y1", 0), s["bbox"][3])
    return lines


def pct(vals, p):
    if not vals:
        return None
    vals = sorted(vals)
    return vals[min(len(vals) - 1, max(0, int(len(vals) * p)))]


def measure(path):
    doc = pymupdf.open(path)
    repeated = repeated_text(doc)
    sizes = collections.Counter()
    fonts = collections.Counter()
    lefts, rights, tops, bots, gaps = [], [], [], [], []

    for page in doc:
        h = page.rect.height
        lines = page_lines(page, repeated)
        if not lines:
            continue
        for l in lines:
            lefts.append(l["x0"])
            rights.append(l["x1"])
        # 页眉页脚已剔除，首行/末行的 y 即版心上下边界；
        # 跨页取 10%/90% 分位，忽略章首页（标题有段前距）与末页（正文不满）
        tops.append(lines[0]["y"])
        bots.append(lines[-1]["y1"])   # 行底，而非行顶
        # 行距只在字号一致的相邻行之间算，避免标题/图题干扰
        for a, b in zip(lines, lines[1:]):
            g = b["y"] - a["y"]
            if 5 < g < 60:
                gaps.append(round(g, 1))

        for blk in page.get_text("dict")["blocks"]:
            if blk.get("type") != 0:
                continue
            for line in blk["lines"]:
                t = "".join(s["text"] for s in line["spans"]).strip()
                if not t or t in repeated:
                    continue
                for s in line["spans"]:
                    if not s["text"].strip():
                        continue
                    sizes[round(s["size"], 1)] += len(s["text"])
                    fonts[cn_font(s["font"])] += len(s["text"])

    rect = doc[0].rect
    return {
        "doc": doc, "rect": rect, "sizes": sizes, "fonts": fonts,
        "repeated": sorted(repeated),
        "margins": {
            "left": pct(lefts, 0.10) / PT_PER_CM,
            "right": (rect.width - pct(rights, 0.90)) / PT_PER_CM,
            "top": pct(tops, 0.10) / PT_PER_CM,
            "bottom": (rect.height - pct(bots, 0.90)) / PT_PER_CM,
        },
        "gaps": gaps,
    }


def report(path):
    m = measure(path)
    doc, r = m["doc"], m["rect"]
    print(f"# PDF 排版测量: {path}")
    print(f"页数 {doc.page_count}")

    print(f"\n## 页面")
    print(f"- 尺寸: {r.width/PT_PER_CM:.2f} x {r.height/PT_PER_CM:.2f} cm")
    mg = m["margins"]
    print(f"- **推算页边距**（取 10%/90% 分位，避开页眉页脚与个别行）")
    print(f"    上 {mg['top']:.2f} / 下 {mg['bottom']:.2f} / "
          f"左 {mg['left']:.2f} / 右 {mg['right']:.2f} cm")

    print(f"\n## 字号分布（按字符数加权，仅版心）")
    total = sum(m["sizes"].values())
    for sz, n in m["sizes"].most_common(10):
        print(f"- {sz:>5}pt  {n:7d} 字  {n/total*100:5.1f}%")

    print(f"\n## 字体分布")
    for f, n in m["fonts"].most_common(10):
        print(f"- {f:<20} {n:7d} 字")

    if m["gaps"]:
        med = statistics.median(m["gaps"])
        print(f"\n## 行距（相邻行基线差，n={len(m['gaps'])}）")
        print(f"- 中位数 {med:.1f}pt   众数分布:")
        for g, n in collections.Counter(m["gaps"]).most_common(6):
            print(f"    {g:>5}pt  ×{n}")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    ap = argparse.ArgumentParser()
    ap.add_argument("input")
    report(ap.parse_args().input)
