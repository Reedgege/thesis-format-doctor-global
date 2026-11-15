#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""探针：在目标平台上，用**本机 Python + 同版本 python-docx** 复现引擎对页眉的操作。

为什么需要它：功能验证在 macOS 上失败时，原因可能有两层——
  A. python-docx / lxml 在这个平台上的 API 行为不同；
  B. Nuitka 打包进产物的那份库与外部安装的不一致。
这个脚本只跑 A 层（外部安装的 python-docx），所以：
  · 这里就失败  → 是 python-docx/lxml 平台层的问题，与打包无关；
  · 这里正常    → 打包层的问题，需要另找。

用法：python3 tools/crossplat/probe_docx.py <fixture.docx> <out.docx>
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
    fixture = sys.argv[1] if len(sys.argv) > 1 else "fixture.docx"
    out = sys.argv[2] if len(sys.argv) > 2 else "probe_out.docx"

    import docx
    ver = getattr(docx, "__version__", None) or "unknown"
    try:
        from docx.oxml.ns import qn  # noqa: F401
        ns_ok = True
    except Exception:
        ns_ok = False
    print("python-docx = %s ; qn 可用 = %s" % (ver, ns_ok))
    if not ns_ok:
        traceback.print_exc()

    d = docx.Document(fixture)
    print("sections = %d" % len(list(d.sections)))

    s = d.sections[0]
    print("sectPr 存在 = %s" % (s._sectPr is not None))

    try:
        hp = s.header.paragraphs
        print("header.paragraphs = %d 段" % len(hp))
        target = hp[0] if hp else s.header.add_paragraph()
        target.alignment = 2   # RIGHT
        print("写入右对齐成功")
    except Exception:
        print("!! 页眉操作抛异常：")
        traceback.print_exc()

    try:
        print("settings 可读 = %s"
              % (d.settings.element is not None))
    except Exception:
        print("!! settings 读取抛异常：")
        traceback.print_exc()

    d.save(out)
    with zipfile.ZipFile(out) as z:
        hf = [n for n in z.namelist() if "header" in n or "footer" in n]
    print("保存后页眉/页脚 part = %s" % (hf or "(none)"))
    return 0


if __name__ == "__main__":
    sys.exit(main())
