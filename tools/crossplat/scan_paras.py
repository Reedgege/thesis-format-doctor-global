# -*- coding: utf-8 -*-
"""扫描全文每个非空段落的字体/字号/行距/缩进，用来判断「到底哪些段没改到」。

用法：python tools/crossplat/scan_paras.py <docx>
"""
from __future__ import annotations

import sys
from collections import Counter

import docx


def _eff(para):
    """段落首个 run 的字体（空 run 则回退到样式链）。"""
    name = size = None
    for r in para.runs:
        if r.font.name or r.font.size:
            name = r.font.name or name
            size = r.font.size.pt if r.font.size else size
            break
    try:
        if name is None:
            name = para.style.font.name
        if size is None and para.style.font.size:
            size = para.style.font.size.pt
    except Exception:
        pass
    return name, size


def main(argv=None) -> int:
    argv = argv if argv is not None else sys.argv[1:]
    if not argv:
        print("usage: scan_paras.py <docx>", file=sys.stderr)
        return 2
    d = docx.Document(argv[0])
    stat = Counter()
    rows = []
    for i, p in enumerate(d.paragraphs):
        t = (p.text or "").strip()
        if not t:
            continue
        name, size = _eff(p)
        try:
            sname = p.style.name
        except Exception:
            sname = "?"
        pf = p.paragraph_format
        fli = round(pf.first_line_indent / 914400.0, 2) if pf.first_line_indent is not None else None
        li = round(pf.left_indent / 914400.0, 2) if pf.left_indent is not None else None
        stat[(name, size, sname)] += 1
        rows.append((i, sname, name, size, pf.line_spacing, fli, li, t[:44]))

    print("=== 字体/字号/样式 分布 ===")
    for (name, size, sname), n in stat.most_common():
        print("  %-22s %-8s %-14s x%d" % (name, size, sname, n))

    print("\n=== 逐段（前 60 段非空） 行距/首行缩进/左缩进 ===")
    for i, sname, name, size, ls, fli, li, t in rows[:60]:
        print("  #%-3d %-12s %-18s %-5s ls=%-4s fli=%-5s li=%-5s %s"
              % (i, sname, name, size, ls, fli, li, t))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
