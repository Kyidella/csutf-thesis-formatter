"""变异测试：故意改坏源码，看测试能不能抓到。

测试全绿不代表测试有效——可能是断言太松、覆盖不到。这个脚本把已知的
几处逻辑逐个改坏，跑一遍测试，统计"抓到几个、漏掉几个"。

漏掉的说明那条逻辑缺测试，应该补断言。

用法:
    python tests/mutation_check.py
"""

import pathlib
import subprocess
import sys

ROOT = pathlib.Path(__file__).resolve().parent.parent

# (文件, 原文, 改后, 说明)  —— 每一条都对应一个真实踩过的坑
MUTATIONS = [
    ("docmodel.py", "TWIPS_PER_CM = 1440 / 2.54", "TWIPS_PER_CM = EMU_PER_CM",
     "页边距换算改回 EMU（2.54cm 会算成 0.004cm）"),
    ("docmodel.py", 'ea = ea if ea is not None else defaults.get("font_cn")',
     "ea = ea if ea is not None else None", "去掉文档级默认中文字体回退"),
    ("docmodel.py", "if asc is None:\n            asc = a",
     "if False:\n            asc = a", "西文字体不沿样式链回退"),
    ("docmodel.py", "def _rfonts_of_rpr(rpr):",
     "def _rfonts_of_rpr(rpr):\n    return None, None", "文档级默认字体读不出"),
    ("docmodel.py", "if ol is None:\n            ol = _style_outline(styles, chain_id)",
     "if False:\n            ol = _style_outline(styles, chain_id)",
     "大纲级别不再继承样式"),
    ("classifier.py", "if p.in_table:", "if False:", "去掉「表格内不判标题」守卫"),
    ("classifier.py", "if lvl == 1:\n                ctx_key = None",
     "if lvl >= 1:\n                ctx_key = None", "二三级标题也重置章节上下文"),
    ("classifier.py", "if DATE_LEAD.match(t):", "if False:", "日期守卫失效"),
    ("classifier.py", "if first_anchor is not None and p.index < first_anchor:",
     "if False:", "封面区不再排除"),
    ("classifier.py", "if t and len(t) <= 20 and not in_toc:", "if False:",
     "目录区不再识别章节锚点"),
    ("check_docx.py", 'if rule.get("font_en") and info["fonts_en"]:',
     "if False:", "不再检查西文字体"),
    ("check_docx.py", 'elif i["level"] == "warn" and i.get("ok") == 0:',
     "elif False:", "优先区不捞「整体不符规范」的条目"),
    ("check_docx.py", "if key in SKIP_FIX_KEYS:\n        return None",
     "if False:\n        return None",
     "报告把排版器不碰的目录条目也说成能自动修"),
    ("check_docx.py", "if not p.text.strip():\n            continue",
     "if False:\n            continue", "空段也计入格式统计"),
    ("check_docx.py", 'if not cfg.get("enabled"):', "if True:",
     "异常字符检查被整个关掉"),
    ("check_docx.py", "if len(found) > 1:", "if False:",
     "同义字符混用不再报"),
    ("check_docx.py", "if not name:\n                continue",
     "if False:\n                continue", "异常字符判定恒真（会误报一切字符）"),
    # ---- 格式化器 ----
    ("formatter.py", "return abs(int(cur) - int(target)) <= tol",
     "return cur == target", "页边距容差失效（会报 1.5cm→1.5cm 假改动）"),
    ("formatter.py", 'if pgsz.get(f"{W}orient") == "landscape":', "if False:",
     "横向页守卫失效"),
    ("formatter.py", "if src.resolve() == dst.resolve():", "if False:",
     "去掉「拒绝原地修改」防护"),
    ("formatter.py", "if changes and output_path and not dry_run:",
     "if output_path:", "预演模式也写文件"),
    ("formatter.py", 'pgsz.set(f"{W}w", w_tw)\n            pgsz.set(f"{W}h", h_tw)',
     'pgsz.set(f"{W}w", h_tw)\n            pgsz.set(f"{W}h", w_tw)', "页面宽高写反"),
    ("formatter.py", "if fm <= sec.index < body:", "if fm <= sec.index <= body:",
     "正文起页边界差一节"),
    ("formatter.py", 'if p.has_field:', 'if False:', "段落映射不再跳过引用域段落"),
    ("formatter.py", 'if key in SKIP_FIX_KEYS:', 'if False:', "段落映射不再跳过目录条目"),
    ("formatter.py",
     'for tag in RPR_STRIP:\n            node = rpr.find(f"{W}{tag}")\n'
     "            if node is not None:\n                rpr.remove(node)\n"
     "                n += 1",
     'for tag in PPR_STRIP:\n            node = rpr.find(f"{W}{tag}")\n'
     "            if node is not None:\n                rpr.remove(node)\n"
     "                n += 1",
     "run 级清除误用段落级的标签集"),
    ("formatter.py", 'if stat["styled"]:', 'if True:', "幂等性：无改动也报条目"),
    ("formatter.py",
     'data = new.get(item.filename) or zin.read(item.filename)',
     'data = (new[DOCUMENT] if item.filename == STYLES\n'
     '                else new.get(item.filename) or zin.read(item.filename))',
     "把 document.xml 的内容错写到 styles.xml"),
    # ---- 文本层 ----
    ("textfix.py", "        elif tag == \"tab\":\n            s = \"\\t\"",
     "        elif False:\n            s = \"\\t\"",
     "flatten 不把制表符算进文本（下标与 p.text 对不上）"),
    ("textfix.py", "out.add(node)", "out.add(id(node))",
     "域内节点用 id() 收集（lxml 代理是临时的，保护会失效）"),
    ("textfix.py", "elif kind == \"end\":\n                depth = max(0, depth - 1)",
     "elif kind == \"end\":\n                depth = depth",
     "fldChar end 不结束域（域范围一路扩到段末）"),
    ("textfix.py", "elif i + 1 < len(text):           # 跨节点：挂到图号那个节点末尾",
     "elif False:                       # 跨节点：挂到图号那个节点末尾",
     "跨节点的题注不处理（实测有 2 处就是这种）"),
    ("textfix.py", 'elif action == "delete":',
     'elif False:',
     "私有区/零宽字符不再删除"),
    ("textfix.py",
     "            hi.append([i, j])         # 成对记下，左右必须一起改",
     "            hi.append([i])",
     "括号只改左半边（会变成半角全角混排）"),
    ("textfix.py", "        if ch in NO_AUTOFIX:\n            lo.append(i)",
     "        if False:\n            lo.append(i)",
     "方括号也自动全角化（可能改成【】，语义变了）"),
    ("textfix.py", "if node is None or new_ch is None or node in prot:",
     "if node is None or new_ch is None:",
     "域内的标点也改"),
    ("textfix.py", "            if node in prot:\n                skipped += 1",
     "            if False:\n                skipped += 1",
     "域内的异常字符也改"),
    ("textfix.py",
     "    if (not _edge_space(old) and _edge_space(text)) or \\\n"
     "            (not _edge_space(old, end=True) and _edge_space(text, end=True)):\n"
     "        node.set(XML_SPACE, \"preserve\")",
     "    if False:\n        node.set(XML_SPACE, \"preserve\")",
     "不补 xml:space（换出来的首尾空格会被 Word 吃掉）"),
    ("textfix.py", 'return (s[-1:] if end else s[:1]) in SPACE_CHARS',
     'return (s[-1:] if end else s[:1]) in " \\t"',
     "首尾空格判定用字符串包含（空串是任何字符串的子串，会误判）"),
    ("textfix.py", "hit = next(((n, a) for lo, hi, n, a in ranges if lo <= cp <= hi), None)",
     "hit = next(((n, a) for lo, hi, n, a in reversed(ranges) if lo <= cp <= hi), None)",
     "字符区间反向匹配（特例被通用区间盖掉，U+E5E5 会被删而不是换成空格）"),
    ("formatter.py", "        if has_drawing(p.element):", "        if False:",
     "图片段落也套样式（正文的固定行距会把嵌入型图片裁掉）"),
    ("formatter.py",
     "            if ls and ls[0] == \"exact\":" + chr(10) + "                _set_single_spacing(p.element)",
     "            if False:" + chr(10) + "                _set_single_spacing(p.element)",
     "图片段落上的固定行距不再救回"),
    ("check_docx.py", "        if has_drawing(p.element):", "        if False:",
     "图片段落计入样式统计（报出改不了的问题）"),
    ("check_docx.py", "        if NEWLINE not in p.text or r.key.startswith(\"heading\"):",
     "        if True:", "标题与正文同段不再报"),
    ("check_docx.py", "        if len(first) > 45 or not _looks_like_numbered_heading(first):",
     "        if False:", "标题与正文同段判定恒真（把正常段落也报出来）"),
    ("wordstyles.py", 'val = "1" if bold else "0"', 'val = "1"',
     "加粗开关恒为开（标题会被弄成粗体）"),
    ("formatter.py", "        if has_drawing(p.element) and not p.text.strip():",
     "        if has_drawing(p.element):",
     "只要含图就整段跳过（图片锚在标题里时，标题排不上）"),
    ("formatter.py",
     '        if has_inline_drawing(el) and _style_uses_fixed_spacing(rules, key):',
     "        if False:",
     "嵌入型图片段落不钉单倍行距（图会被裁）"),
]


def run_tests():
    r = subprocess.run([sys.executable, "-m", "pytest", "tests/", "-q",
                        "--no-header", "--tb=no"],
                       cwd=ROOT, capture_output=True, text=True,
                       encoding="utf-8", errors="replace", timeout=600)
    return r.returncode != 0, r.stdout


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    caught = escaped = skipped = 0
    escaped_list = []

    for fname, old, new, desc in MUTATIONS:
        p = ROOT / fname
        orig = p.read_text(encoding="utf-8")
        if old not in orig:
            print(f"  [跳过] 锚点找不到：{desc}")
            skipped += 1
            continue
        p.write_text(orig.replace(old, new, 1), encoding="utf-8")
        try:
            failed, _ = run_tests()
        finally:
            p.write_text(orig, encoding="utf-8")

        if failed:
            caught += 1
            print(f"  [抓到] {desc}")
        else:
            escaped += 1
            escaped_list.append(desc)
            print(f"  [漏掉] {desc}  ← 缺对应测试")

    total = len(MUTATIONS) - skipped
    print(f"\n共 {total} 处人为破坏：抓到 {caught}，漏掉 {escaped}")
    if escaped_list:
        print("\n需要补测试的：")
        for d in escaped_list:
            print(f"  - {d}")

    ok, out = run_tests()
    last = [l for l in out.splitlines() if l.strip()]
    print(f"\n还原后：{last[-1] if last else '(无输出)'}")
    return 1 if escaped else 0


if __name__ == "__main__":
    sys.exit(main())
