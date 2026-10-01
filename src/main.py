"""PaperFormat Pro — command-line entry point (offline).

Usage:
  python -m src.main gui                       # launch the desktop GUI
  python -m src.main check thesis.docx --spec APA
  python -m src.main check thesis.docx --spec APA --template school-template.docx
  python -m src.main check thesis.docx --ai ai-filled-form.json
  python -m src.main check thesis.docx --latex school-template.tex
  python -m src.main check thesis.docx --spec APA --json          # structured output
  python -m src.main fix   thesis.docx --spec APA [--out output.docx] [--preview] [--json]
  python -m src.main activate <code>
  python -m src.main status                    # show license status
  python -m src.main export-schema out.yaml    # export the AI questionnaire template

Notes:
- Checking (check) is always free, with no gate.
- Fixing (fix) is free the first time, then requires an activation code; if the
  document already complies, nothing is written and no trial credit is used.
- JSON output is always emitted as a UTF-8 byte stream, parseable by any terminal.
"""

from __future__ import annotations

import argparse
import json
import os
import sys

from .engine import checker, report as report_mod
from .engine.docx_reader import DocxReadError
from .engine.fixer import (
    compute_changes, fix_docx, preview_fix, same_file, unique_output_path,
)
from .engine.questionnaire import build_target
from .engine.specs import SPEC_ORDER
from .versioninfo import app_version

_TEMPLATE = os.path.join(os.path.dirname(__file__), "data", "ai_questionnaire_template.yaml")

_CONSOLE_READY = False


def _fatal_dialog(msg: str):
    """把启动期的致命错误弹给用户看（打包成窗口程序后没有控制台可看）。

    优先用 Windows 原生 MessageBox：此时 Tk 可能还没起来（或者就是 Tk 起不来），
    不能指望用 Tk 弹窗。非 Windows 退化为 stderr。
    """
    try:
        import ctypes

        ctypes.windll.user32.MessageBoxW(0, msg, "PaperFormat Pro", 0x10)
        return
    except Exception:
        pass
    try:
        print(msg, file=sys.stderr)
    except Exception:
        pass


def _has_console() -> bool:
    """打包成 GUI 程序（--windows-console-mode=disable）时 stdout/stderr 为 None。"""
    try:
        return sys.stdout is not None and sys.stderr is not None
    except Exception:
        return False


def _init_console():
    """让命令行输出在各类环境下都不崩、且可被外部程序解析（幂等）。

    - 交互式终端（isatty）：保持终端原编码（中文 Windows 为 GBK，中文显示正常），
      仅放宽错误处理，避免报告里的 ✅/❌ 触发 UnicodeEncodeError 直接崩溃；
    - 管道 / 重定向 / 被 subprocess 捕获：强制 UTF-8，保证输出可被外部程序正确解析。

    竞态/异常一律吞掉：控制台输出问题永远不该让命令失败。
    """
    global _CONSOLE_READY
    if _CONSOLE_READY:
        return
    _CONSOLE_READY = True
    for stream in (sys.stdout, sys.stderr):
        try:
            if not stream.isatty():
                stream.reconfigure(encoding="utf-8", errors="replace")
            else:
                stream.reconfigure(errors="replace")
        except Exception:
            pass


def _build_target(args) -> "object":
    questionnaire: dict = None
    q_path = getattr(args, "questionnaire", None)
    if q_path:
        with open(q_path, "r", encoding="utf-8") as f:
            questionnaire = json.load(f)
    return build_target(
        args.spec,
        template=getattr(args, "template", None),
        ai=getattr(args, "ai", None),
        questionnaire=questionnaire,
        latex=getattr(args, "latex", None),
    )


def _write_json(text: str):
    """JSON 一律以 UTF-8 字节输出（机器可解析、重定向不乱码）。

    绕开终端编码（Windows 中文控制台为 GBK）导致的 errors="replace" 降级，
    否则中文/箭头会被替换成 '?'，外部程序无法解析。
    无 sys.stdout.buffer（自定义流）时退化为文本写 + 放宽错误处理。
    """
    try:
        buf = sys.stdout.buffer
        buf.write(text.encode("utf-8"))
        buf.write(b"\n")
        buf.flush()
        return
    except Exception:
        pass
    try:
        sys.stdout.reconfigure(errors="replace")
    except Exception:
        pass
    sys.stdout.write(text + "\n")
    sys.stdout.flush()


def report_to_dict(docx_path, target, prof, findings) -> dict:
    """把体检结果整理成结构化 dict（供 --json 输出 / 外部集成复用）。"""
    rep = report_mod.build_report(docx_path, target, prof, findings)
    return {
        "meta": rep.meta,
        "summary": rep.summary,
        "resolution": [{"dimension": d, "source": s} for d, s in rep.resolution],
        "conflicts": [{"dimension": d, "from": l, "to": r} for d, l, r in rep.conflicts],
        "notes": list(getattr(target, "notes", ()) or ()),
        "findings": [
            {"dimension": f.dimension, "severity": f.severity, "message": f.message,
             "expected": f.expected, "actual": f.actual, "source": f.source}
            for f in findings
        ],
        "profile": {
            "margins": prof.margins,
            "margins_inconsistent": prof.margins_inconsistent,
            "font_family": prof.font_family,
            "font_size_pt": prof.font_size_pt,
            "line_spacing": prof.line_spacing,
            "body_alignment": prof.body_alignment,
            "first_line_indent_in": prof.first_line_indent_in,
            "space_after_pt": prof.space_after_pt,
            "table_paragraph_count": prof.table_paragraph_count,
            "heading_levels_used": prof.heading_levels_used,
            "reference_entries": len(prof.reference_entries),
            "reference_hanging_indent_in": prof.reference_hanging_indent_in,
        },
    }


def cmd_check(args):
    _init_console()
    if not os.path.isfile(args.docx):
        print(f"Error: document not found: {args.docx}", file=sys.stderr)
        return 2
    try:
        target = _build_target(args)
        prof, findings = checker.run_check(args.docx, target)
    except DocxReadError as e:
        print(f"Error: {e}", file=sys.stderr)
        return 2
    except (ValueError, FileNotFoundError, json.JSONDecodeError) as e:
        print(f"Error: failed to build target profile: {e}", file=sys.stderr)
        return 2
    if getattr(args, "json", False):
        text = json.dumps(report_to_dict(args.docx, target, prof, findings),
                          ensure_ascii=False, indent=2)
        _write_json(text)
        if args.out:
            try:
                with open(args.out, "w", encoding="utf-8") as f:
                    f.write(text)
            except OSError as e:
                print(f"Error: cannot write {args.out}: {e}", file=sys.stderr)
                return 2
            print(f"\nJSON report written to: {args.out}", file=sys.stderr)
        return 0
    rep = report_mod.build_report(args.docx, target, prof, findings)
    print(rep.markdown)
    if args.out:
        out = args.out
        try:
            if out.lower().endswith(".docx"):
                report_mod.build_docx_report(rep, out)
            else:
                with open(out, "w", encoding="utf-8") as f:
                    f.write(rep.markdown)
        except OSError as e:
            print(f"Error: cannot write {out}: {e}", file=sys.stderr)
            return 2
        print(f"\nReport written to: {out}")
    return 0


def cmd_gui(args):
    """启动桌面 GUI。

    任何失败都要**弹出来告诉用户**：本程序在 Windows 上以窗口子系统编译
    （无控制台），一旦这里静默 return，用户看到的就是"双击没反应"。
    """
    try:
        from .ui.app import main as gui_main
    except Exception as e:
        _fatal_dialog("Cannot start the graphical interface: Tk (_tkinter) is missing from this environment.\n\n"
                      "Details: %s\n\n"
                      "You can use the `check` subcommand from the command line instead." % e)
        return 1
    try:
        gui_main()
        return 0
    except Exception as e:  # noqa: BLE001  GUI crashes must still be visible to the user
        _fatal_dialog("The graphical interface exited unexpectedly:\n\n%s: %s" % (type(e).__name__, e))
        return 1


def cmd_export_schema(args):
    if not os.path.isfile(_TEMPLATE):
        print("Error: built-in template file not found.", file=sys.stderr)
        return 2
    with open(_TEMPLATE, "r", encoding="utf-8") as src, \
        open(args.out, "w", encoding="utf-8") as out:
        out.write(src.read())
    print(f"AI questionnaire template exported to: {args.out}")
    return 0


def _default_fixed_path(input_path: str) -> str:
    """默认输出文件名：原名_fixed.docx。"""
    base, ext = os.path.splitext(input_path)
    return f"{base}_fixed{ext}"


def _summary_lines(summary: dict) -> list:
    """把 fix_docx 的变更摘要整理成人可读的行。"""
    out: list = []
    page = summary.get("page") or {}
    applied = summary.get("applied") or {}
    if page:
        margins = {k.replace("margin_", "").replace("_in", ""): v
                   for k, v in page.items() if k.startswith("margin_") and v is not None}
        if margins:
            out.append("Margins → " + ", ".join(f"{s} {v}″" for s, v in margins.items()))
        if page.get("font_family"):
            sz = f" {page['font_size_pt']}pt" if page.get("font_size_pt") else ""
            out.append(f"Body font → {page['font_family']}{sz}")
        if page.get("line_spacing") is not None:
            out.append(f"Line spacing → {page['line_spacing']}x")
        if applied.get("body_alignment"):
            out.append(f"Body alignment → {applied['body_alignment']}")
        if applied.get("first_line_indent_in") is not None:
            out.append(f"Body first-line indent → {applied['first_line_indent_in']}″")
        if applied.get("space_after_pt") is not None:
            out.append(f"Space after paragraph → {applied['space_after_pt']}pt")
    if summary.get("reference_style"):
        kind = "Numbered [n]" if summary.get("reference_numbered") else "Author-Year"
        out.append(f"References → {summary['reference_style']}"
                   f" (hanging indent {summary.get('reference_hanging_indent_in')}″, {kind})")
        added = summary.get("reference_numbering_added") or 0
        if added:
            out.append(f"Reference numbering → added [n] prefix to {added} unnumbered entries (prefix only, text unchanged)")
    out.append("Heading paragraphs skipped (heading size/font unchanged, original visual hierarchy preserved)")
    return out


def _fix_gate_cli(gate, as_json: bool) -> int:
    """门禁不通过时的输出。返回退出码。"""
    if as_json:
        _write_json(json.dumps({"ok": False, "reason": gate["reason"],
                               "message": gate["message"]},
                               ensure_ascii=False, indent=2))
    else:
        print(f"✗ {gate['message']}", file=sys.stderr)
        print("  Tip: run `python -m src.main activate <code>` to unlock unlimited fixes.",
              file=sys.stderr)
    return 3


def cmd_fix(args):
    """一键修正：查免费、修正首次免费、之后需激活码；纯格式层。

    执行顺序有意为之（**先校验入参，再谈授权**）：
      1) 输入文件是否存在 → rc=2；
      2) ``--out`` 是否指向原件 → rc=2（参数错误，不该被报成「试用已用完」）；
      3) 构建目标画像（spec/模板/问卷/AI/LaTeX 任一有问题 → rc=2）；
      4) 预览（受门禁）；
      5) 已合规 → rc=0 且**不消耗试用、不触发门禁**（与 `check` 同级的免费体检）；
      6) 授权门禁 → rc=3；
      7) 落盘 + 计次。
    """
    _init_console()
    as_json = bool(getattr(args, "json", False))

    # ① Input file
    if not os.path.isfile(args.docx):
        print(f"Error: document not found: {args.docx}", file=sys.stderr)
        return 2

    # ② Output path safety: a parameter-level error, returned before the license decision
    if not args.preview and args.out and same_file(args.out, args.docx):
        print("Error: output path points to the same file as the input (original is not overwritten).", file=sys.stderr)
        return 2

    # ③ 目标画像（含 spec / 模板 / 问卷 / AI / LaTeX 的读取与校验）
    try:
        target = _build_target(args)
    except DocxReadError as e:
        print(f"Error: {e}", file=sys.stderr)
        return 2
    except (ValueError, FileNotFoundError, json.JSONDecodeError) as e:
        print(f"Error: failed to build target profile: {e}", file=sys.stderr)
        return 2

    from .license import license as lic

    # ④ 预览模式：受门禁，只列改动、不落盘（也不消耗试用）
    if args.preview:
        gate = lic.require_fix_entitlement()
        if not gate["allowed"]:
            return _fix_gate_cli(gate, as_json)
        try:
            changes = preview_fix(args.docx, target)
        except DocxReadError as e:
            print(f"Error: {e}", file=sys.stderr)
            return 2
        if as_json:
            _write_json(json.dumps({"ok": True, "mode": "preview", "changes": changes,
                                    "license": gate}, ensure_ascii=False, indent=2))
            return 0
        print("Format fixes to apply (preview — only the items that will actually change):")
        for line in changes:
            print("  - " + line)
        print(f"  Trial status: {gate['message']}")
        return 0

    # ⑤ 已合规 → 不生成副本、不消耗试用额度、不触发门禁
    #    （与 check 同级：单纯读文档比对，无任何副作用，也告诉用户「不用花钱」）
    try:
        changes = compute_changes(args.docx, target)
    except DocxReadError as e:
        print(f"Error: {e}", file=sys.stderr)
        return 2
    if not changes:
        msg = "This document already meets the target format. No copy was generated and no trial credit was used."
        if as_json:
            _write_json(json.dumps({"ok": True, "mode": "noop", "output": None,
                                    "changes": [], "message": msg},
                                   ensure_ascii=False, indent=2))
        else:
            print(f"✓ {msg}")
        return 0

    # ⑥ 授权门禁（模型 B：未激活 → 首次免费；试用已用 → 需激活）
    gate = lic.require_fix_entitlement()
    if not gate["allowed"]:
        return _fix_gate_cli(gate, as_json)

    # ⑦ 落盘
    out = args.out or _default_fixed_path(args.docx)
    if not args.out:
        out = unique_output_path(out)   # 默认输出不覆盖上一次的修正稿
    try:
        summary = fix_docx(args.docx, out, target)
    except DocxReadError as e:
        print(f"Error: {e}", file=sys.stderr)
        return 2
    except ValueError as e:
        print(f"Error: {e}", file=sys.stderr)
        return 2
    except OSError as e:
        print(f"Error: {e}", file=sys.stderr)
        return 2

    counted = lic.record_fix_used()
    after = lic.post_fix_message(gate, counted)
    detail = _summary_lines(summary)
    if as_json:
        _write_json(json.dumps({"ok": True, "mode": "fix", "output": out,
                                "summary": summary, "changes": detail,
                                "license": gate, "license_after": after,
                                "counted": counted},
                               ensure_ascii=False, indent=2))
        return 0
    print(f"✓ Fixed document generated: {out}")
    print(f"  {after}")
    print("  Applied this run:")
    for line in detail:
        print("    " + line)
    return 0


def cmd_activate(args):
    from .license import license as lic
    if getattr(args, "offline", False):
        r = lic.activate_offline(args.code)
    else:
        r = lic.activate(args.code)
    print(r["message"])
    return 0 if r["ok"] else 4


def cmd_status(args):
    from .license import license as lic
    s = lic.status()
    print(f"Activated: {'Yes' if s['activated'] else 'No'}")
    print(f"Type: {s['kind'] or '-'}")
    print(f"Trial used: {'Yes' if s['trial_fix_used'] else 'No'} (limit {s['trial_limit']})")
    if s.get("revoked"):
        print("Status: revoked (reactivation required)")
    print(f"Machine fingerprint: {s['machine']}")
    print(f"Local state file: {'OK' if s.get('state_ok') else 'missing/corrupted'}")
    return 0


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(description="PaperFormat Pro — offline thesis format checker & fixer")
    # --version 用 argparse 内置 action：解析到就打印并 SystemExit(0)，
    # 不进入任何子命令 —— 版本号读 VERSION 文件（打包时随产物一起带出）。
    p.add_argument("-V", "--version", action="version",
                   version="PaperFormat Pro %s" % app_version())
    sub = p.add_subparsers(dest="cmd")

    g = sub.add_parser("gui", help="Launch the desktop GUI")
    g.set_defaults(func=cmd_gui)

    c = sub.add_parser("check", help="Run a format check from the command line")
    c.add_argument("docx", help="Path to the Word document to check")
    c.add_argument("--spec", default="APA", choices=SPEC_ORDER, help="Citation style (default APA)")
    c.add_argument("--template", help="School template docx (optional; page layout follows the school)")
    c.add_argument("--ai", help="AI-filled questionnaire config (YAML/JSON, optional)")
    c.add_argument("--questionnaire", help="Manually filled questionnaire JSON (optional)")
    c.add_argument("--latex", help="LaTeX template path (optional, 5th source)")
    c.add_argument("--out", help="Report output path (.md or .docx; .json when used with --json)")
    c.add_argument("--json", action="store_true", help="Output structured results as JSON (for integration/automation)")
    c.set_defaults(func=cmd_check)

    f = sub.add_parser("fix", help="One-click fix (format only; first use free, then requires activation)")
    f.add_argument("docx", help="Path to the Word document to fix")
    f.add_argument("--spec", default="APA", choices=SPEC_ORDER, help="Citation style (default APA)")
    f.add_argument("--template", help="School template docx (optional)")
    f.add_argument("--ai", help="AI-filled questionnaire config (YAML/JSON, optional)")
    f.add_argument("--questionnaire", help="Manually filled questionnaire JSON (optional)")
    f.add_argument("--latex", help="LaTeX template path (optional)")
    f.add_argument("--out", help="Output path (default <name>_fixed.docx; auto-numbered if it already exists)")
    f.add_argument("--preview", action="store_true", help="Preview the changes only; do not write to disk")
    f.add_argument("--json", action="store_true", help="Output structured results as JSON (for integration/automation)")
    f.set_defaults(func=cmd_fix)

    a = sub.add_parser("activate", help="Enter an activation code to unlock unlimited fixes")
    a.add_argument("code", help="Activation code")
    a.add_argument("--offline", action="store_true",
                   help="Treat as an offline activation code (issued by the seller, no internet required)")
    a.set_defaults(func=cmd_activate)

    s = sub.add_parser("status", help="Show the current license status")
    s.set_defaults(func=cmd_status)

    e = sub.add_parser("export-schema", help="Export the AI questionnaire template")
    e.add_argument("out", help="Output YAML path")
    e.set_defaults(func=cmd_export_schema)
    return p


def main(argv=None):
    _init_console()
    parser = build_parser()
    args = parser.parse_args(argv)
    if not getattr(args, "cmd", None):
        # 无子命令：
        #  - 有控制台（源码运行 / 终端里跑）-> 打印帮助，保持 CLI 习惯；
        #  - 没有控制台（打包成窗口程序，stdout 为 None）-> 必须进 GUI，
        #    否则用户双击 exe 时"什么都没发生"（帮助文本没有任何地方能显示）。
        if _has_console():
            try:
                parser.print_help()
                return 0
            except Exception:
                pass
        return cmd_gui(args)
    return args.func(args)


if __name__ == "__main__":
    raise SystemExit(main())
