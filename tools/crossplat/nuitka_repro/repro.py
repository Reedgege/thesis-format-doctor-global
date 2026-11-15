#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""最小复现：只用 Nuitka 打包 python-docx，验证「创建页眉」在产物里能不能work。

背景：线上产物在 macOS / Linux 上修不出页码（页眉 part 根本没创建），
Windows 正常；同一份源码在 macOS 上用外部 python-docx 跑又完全正常。
所以嫌疑落在「Nuitka 打包后的 python-docx」这一层。

这个脚本只做一件事：走完 HeaderPart.new 的每一步并把中间值打出来
（__file__ 在哪、默认模板 XML 读不读得到、header.paragraphs 能不能访问），
三个平台各编译一次对照，一眼看出断点。

用 Nuitka 编译后运行：
  python -m nuitka --standalone --assume-yes-for-downloads repro.py
  ./repro.dist/repro <fixture.docx> <out.docx>
"""

import os
import sys
import traceback
import zipfile

for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass


def main() -> int:
    fixture = sys.argv[1] if len(sys.argv) > 1 else "in.docx"
    out = sys.argv[2] if len(sys.argv) > 2 else "out.docx"

    print("frozen = %s" % (getattr(sys, "frozen", False),))
    print("platform = %s" % sys.platform)

    try:
        import docx
        print("docx     loaded, __file__ = %r" % (getattr(docx, "__file__", None),))
    except Exception:
        print("!! import docx 失败")
        traceback.print_exc()
        return 1

    # ① 关键子模块能不能 import，__file__ 指向哪
    for m in ("docx.parts.hdrftr", "docx.oxml.section", "docx.opc.packuri",
              "docx.parts.document"):
        try:
            mod = __import__(m, fromlist=["_"])
            print("%-22s __file__ = %r" % (m, getattr(mod, "__file__", None)))
        except Exception as e:
            print("!! %s import 失败：%r" % (m, e))
            traceback.print_exc()

    # ② HeaderPart.new 第一步：读默认页眉模板
    try:
        from docx.parts.hdrftr import HeaderPart
        raw = HeaderPart._default_header_xml()
        print("_default_header_xml OK：%d bytes" % len(raw))
    except Exception:
        print("!! _default_header_xml 失败（页眉模板读不到 → 建不了页眉 part）")
        traceback.print_exc()
        raw = None

    # ③ 真正走一遍：打开文档 → 访问 header.paragraphs（会创建 part）→ 保存
    try:
        d = docx.Document(fixture)
        print("sections = %d" % len(list(d.sections)))
        s = d.sections[0]
        try:
            ps = s.header.paragraphs
            print("header.paragraphs OK：%d 段" % len(ps))
            tgt = ps[0] if ps else s.header.add_paragraph()
            tgt.alignment = 2
            print("写入对齐成功")
        except Exception:
            print("!! header 操作失败")
            traceback.print_exc()
        d.save(out)
        with zipfile.ZipFile(out) as z:
            hf = [n for n in z.namelist() if "header" in n or "footer" in n]
        print("保存后 hf parts = %s" % (hf or "(none)"))
    except Exception:
        print("!! 文档处理失败")
        traceback.print_exc()
    return 0


if __name__ == "__main__":
    sys.exit(main())
