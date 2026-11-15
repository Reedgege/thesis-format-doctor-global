#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""在目标平台上用**源码**复现「页码修正」这一环，把每一步的值都打印出来。

为什么要它：线上产物在 macOS / Linux 上修不出页码，但 Windows 正常，
而引擎代码里没有任何平台分支。单看产物只能看到"没写页眉"这个结果，
看不到是哪一环断的。这个脚本走源码，逐环打印：

  target.page_number_position → fix_dims 里有没有 page_number
  → doc.sections 数量 → _apply_page_number 的返回值 → 保存后有没有 header part

不需要激活码：只跑 check + 直接调用一个内部函数，不走授权门禁。

用法：python3 tools/crossplat/debug_page_number.py <fixture.docx> <out.docx>
"""

import os
import sys
import traceback

for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

_REPO = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
if _REPO not in sys.path:
    sys.path.insert(0, _REPO)


def line(msg=""):
    print(msg, flush=True)


def main() -> int:
    fixture = sys.argv[1] if len(sys.argv) > 1 else "fixture.docx"
    out = sys.argv[2] if len(sys.argv) > 2 else "debug_out.docx"

    import docx
    from docx.section import Section, Sections

    line("python-docx = %s" % (getattr(docx, "__version__", "unknown"),))
    line("python = %s" % sys.version.replace("\n", " "))
    line("platform = %s %s" % (sys.platform, os.uname().machine if hasattr(os, "uname") else "?"))

    from src.engine.questionnaire import build_target
    from src.engine import checker, fixer

    SPEC = "APA-Student"
    target = build_target(SPEC)
    line("")
    line("target.page_number_position = %r" % (getattr(target, "page_number_position", "<无此属性>"),))

    prof, findings = checker.run_check(fixture, target)
    fix_dims = {f.fix.dimension for f in findings if getattr(f, "fix", None) is not None}
    line("findings 条数 = %d" % len(findings))
    line("fix_dims = %s" % sorted(fix_dims))
    line("'page_number' in fix_dims = %s" % ("page_number" in fix_dims))
    for f in findings:
        if getattr(f, "dimension", "") == "page_number":
            fxv = getattr(f, "fix", None)
            line("  page_number finding: severity=%r fix=%r"
                 % (getattr(f, "severity", None), fxv))

    line("")
    d = docx.Document(fixture)
    try:
        secs = list(d.sections)
        line("doc.sections = %d 个" % len(secs))
        s0 = secs[0] if secs else None
        if s0 is not None:
            line("  s0 是 Section: %s" % isinstance(s0, Section))
            line("  s0.header = %r" % (s0.header,))
            line("  is_linked_to_previous = %r" % (s0.header.is_linked_to_previous,))
            line("  different_first_page = %r" % (s0.different_first_page_header_footer,))
    except Exception:
        line("!! list(doc.sections) 抛异常：")
        traceback.print_exc()

    # 打点：看 _ensure_page_number_in_part 有没有被调用、被谁调用
    _orig_ensure = fixer._ensure_page_number_in_part
    calls = []

    def _traced(part, align):
        calls.append((type(part).__name__, align))
        try:
            r = _orig_ensure(part, align)
            line("  [_ensure_page_number_in_part] part=%s align=%r -> %r"
                 % (type(part).__name__, align, r))
            return r
        except Exception:
            line("  [_ensure_page_number_in_part] 抛异常：")
            traceback.print_exc()
            raise

    fixer._ensure_page_number_in_part = _traced

    line("")
    line("--- 直接调用 _apply_page_number ---")
    try:
        res = fixer._apply_page_number(d, target.page_number_position)
        line("返回值 = %r" % (res,))
    except Exception:
        line("!! _apply_page_number 抛异常：")
        traceback.print_exc()
    line("_ensure 被调用 %d 次：%r" % (len(calls), calls))

    fixer._ensure_page_number_in_part = _orig_ensure

    line("")
    try:
        d.save(out)
        import zipfile
        with zipfile.ZipFile(out) as z:
            names = z.namelist()
            hf = [n for n in names if "header" in n or "footer" in n]
        line("保存后 zip_entries=%d hf_parts=%s" % (len(names), hf or "(none)"))
    except Exception:
        line("!! 保存失败：")
        traceback.print_exc()
    return 0


if __name__ == "__main__":
    sys.exit(main())
