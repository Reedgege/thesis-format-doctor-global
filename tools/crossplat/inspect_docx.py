# -*- coding: utf-8 -*-
"""把输出文档的真实属性打印出来 —— 「报告说改了」不算，文档里得真有。

用途：跨平台验收时逐个平台跑一遍，比对这三件事：
  1. 页眉里是否真有 PAGE 域、是否右对齐；
  2. 页边距 / 行距 / 首行缩进 / 字体是否按规范改到位；
  3. 参考文献是否悬挂缩进、是否另起一页。

用法：python tools/crossplat/inspect_docx.py <docx> [--json]
"""
from __future__ import annotations

import json
import sys
import zipfile

import docx


def _inch(v):
    return round(v / 914400.0, 3) if v is not None else None


def inspect(path: str) -> dict:
    d = docx.Document(path)
    sec = d.sections[0]
    out = {
        "margins_in": {
            "top": _inch(sec.top_margin),
            "bottom": _inch(sec.bottom_margin),
            "left": _inch(sec.left_margin),
            "right": _inch(sec.right_margin),
        },
        "header": [],
        "footer": [],
        "body_first": None,
        "references": [],
        "zip": {},
    }

    try:
        for i, para in enumerate(sec.header.paragraphs):
            xml = para._p.xml
            out["header"].append({
                "i": i,
                "align": str(para.alignment),
                "has_page_field": "PAGE" in xml,
                "text": (para.text or "")[:40],
            })
    except Exception as e:
        out["header"] = "error: %r" % (e,)

    try:
        for i, para in enumerate(sec.footer.paragraphs):
            out["footer"].append({
                "i": i,
                "align": str(para.alignment),
                "has_page_field": "PAGE" in para._p.xml,
                "text": (para.text or "")[:40],
            })
    except Exception as e:
        out["footer"] = "error: %r" % (e,)

    for para in d.paragraphs:
        t = (para.text or "").strip()
        if out["body_first"] is None and len(t) > 120:
            pf, r = para.paragraph_format, (para.runs[0] if para.runs else None)
            out["body_first"] = {
                "text": t[:40],
                "font": r.font.name if r else None,
                "size_pt": r.font.size.pt if (r and r.font.size) else None,
                "line_spacing": pf.line_spacing,
                "first_line_indent_in": _inch(pf.first_line_indent),
            }
        if t.startswith(("Brown, T.", "Davenport, T.", "Gawande, A.", "Suchman, L.")):
            pf = para.paragraph_format
            out["references"].append({
                "text": t[:40],
                "left_in": _inch(pf.left_indent),
                "first_in": _inch(pf.first_line_indent),
                "line_spacing": pf.line_spacing,
            })

    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        out["zip"] = {
            "entries": len(names),
            "hf_parts": sorted(n.split("/")[-1] for n in names
                               if n.startswith("word/header") or n.startswith("word/footer")),
            "has_comments": any("comments" in n for n in names),
        }
    return out


def main(argv=None) -> int:
    argv = argv if argv is not None else sys.argv[1:]
    if not argv:
        print("usage: inspect_docx.py <docx> [--json]", file=sys.stderr)
        return 2
    path = argv[0]
    data = inspect(path)
    if "--json" in argv:
        print(json.dumps(data, ensure_ascii=False, indent=2))
        return 0

    print("页边距(in): " + json.dumps(data["margins_in"], ensure_ascii=False))
    print("页眉: " + json.dumps(data["header"], ensure_ascii=False))
    print("页脚: " + json.dumps(data["footer"], ensure_ascii=False))
    print("正文首段: " + json.dumps(data["body_first"], ensure_ascii=False))
    print("参考文献: " + json.dumps(data["references"], ensure_ascii=False))
    print("包结构: " + json.dumps(data["zip"], ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
