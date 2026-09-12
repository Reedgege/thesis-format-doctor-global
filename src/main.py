"""Thesis Format Doctor Global —— 命令行入口（离线）。

用法：
  python -m src.main gui                       # 启动桌面 GUI
  python -m src.main check 论文.docx --spec APA
  python -m src.main check 论文.docx --spec APA --template 学校模板.docx
  python -m src.main check 论文.docx --ai ai填的表单.json
  python -m src.main check 论文.docx --latex 学校模板.tex
  python -m src.main check 论文.docx --spec APA --json          # 结构化输出
  python -m src.main fix   论文.docx --spec APA [--out 输出.docx] [--preview] [--json]
  python -m src.main activate 激活码
  python -m src.main status                    # 查看授权状态
  python -m src.main export-schema out.json    # 导出 AI 填表模板

说明：
- 检查（check）永远免费、无门禁；
- 修正（fix）首次免费，之后需激活码；文档已合规时不落盘、不消耗试用；
- JSON 输出一律走 UTF-8 字节流，任何终端都可被外部程序解析。
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

_TEMPLATE = os.path.join(os.path.dirname(__file__), "data", "ai_questionnaire_template.json")

_CONSOLE_READY = False


def _fatal_dialog(msg: str):
    """把启动期的致命错误弹给用户看（打包成窗口程序后没有控制台可看）。

    优先用 Windows 原生 MessageBox：此时 Tk 可能还没起来（或者就是 Tk 起不来），
    不能指望用 Tk 弹窗。非 Windows 退化为 stderr。
    """
    try:
        import ctypes

        ctypes.windll.user32.MessageBoxW(0, msg, "Thesis Format Doctor Global", 0x10)
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
        print(f"错误：找不到文档 {args.docx}", file=sys.stderr)
        return 2
    try:
        target = _build_target(args)
        prof, findings = checker.run_check(args.docx, target)
    except DocxReadError as e:
        print(f"错误：{e}", file=sys.stderr)
        return 2
    except (ValueError, FileNotFoundError, json.JSONDecodeError) as e:
        print(f"错误：构建目标画像失败：{e}", file=sys.stderr)
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
                print(f"错误：无法写入 {args.out}：{e}", file=sys.stderr)
                return 2
            print(f"\nJSON 报告已写入：{args.out}", file=sys.stderr)
        return 0
    rep = report_mod.build_report(args.docx, target, prof, findings)
    print(rep.markdown)
    if args.out:
        try:
            with open(args.out, "w", encoding="utf-8") as f:
                f.write(rep.markdown)
        except OSError as e:
            print(f"错误：无法写入 {args.out}：{e}", file=sys.stderr)
            return 2
        print(f"\n报告已写入：{args.out}")
    return 0


def cmd_gui(args):
    """启动桌面 GUI。

    任何失败都要**弹出来告诉用户**：本程序在 Windows 上以窗口子系统编译
    （无控制台），一旦这里静默 return，用户看到的就是"双击没反应"。
    """
    try:
        from .ui.app import main as gui_main
    except Exception as e:
        _fatal_dialog("无法启动图形界面：当前环境缺少 Tk（_tkinter）。\n\n"
                      "细节：%s\n\n"
                      "可改用命令行 `check` 子命令做体检。" % e)
        return 1
    try:
        gui_main()
        return 0
    except Exception as e:  # noqa: BLE001  界面异常退出也要让用户看得见
        _fatal_dialog("图形界面异常退出：\n\n%s: %s" % (type(e).__name__, e))
        return 1


def cmd_export_schema(args):
    if not os.path.isfile(_TEMPLATE):
        print("错误：找不到内置模板文件。", file=sys.stderr)
        return 2
    with open(_TEMPLATE, "r", encoding="utf-8") as src, \
            open(args.out, "w", encoding="utf-8") as out:
        out.write(src.read())
    print(f"AI 问卷模板已导出：{args.out}")
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
            out.append("页边距 → " + "，".join(f"{s} {v}″" for s, v in margins.items()))
        if page.get("font_family"):
            sz = f" {page['font_size_pt']}pt" if page.get("font_size_pt") else ""
            out.append(f"正文字体 → {page['font_family']}{sz}")
        if page.get("line_spacing") is not None:
            out.append(f"行距 → {page['line_spacing']}x")
        if applied.get("body_alignment"):
            out.append(f"正文对齐 → {applied['body_alignment']}")
        if applied.get("first_line_indent_in") is not None:
            out.append(f"正文首行缩进 → {applied['first_line_indent_in']}″")
        if applied.get("space_after_pt") is not None:
            out.append(f"段后间距 → {applied['space_after_pt']}pt")
    if summary.get("reference_style"):
        kind = "顺序编码 [n]" if summary.get("reference_numbered") else "著者-出版年"
        out.append(f"参考文献 → {summary['reference_style']}"
                   f"（悬挂缩进 {summary.get('reference_hanging_indent_in')}″，{kind}）")
        added = summary.get("reference_numbering_added") or 0
        if added:
            out.append(f"参考文献编号 → 为 {added} 条未编号条目补 [n] 前缀（仅新增前缀，不改文字）")
    out.append("标题段已跳过（不改动标题字号/字体，保留原视觉层级）")
    return out


def _fix_gate_cli(gate, as_json: bool) -> int:
    """门禁不通过时的输出。返回退出码。"""
    if as_json:
        _write_json(json.dumps({"ok": False, "reason": gate["reason"],
                               "message": gate["message"]},
                               ensure_ascii=False, indent=2))
    else:
        print(f"✗ {gate['message']}", file=sys.stderr)
        print("  提示：执行 `python -m src.main activate 激活码` 解锁无限修正。",
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

    # ① 输入文件
    if not os.path.isfile(args.docx):
        print(f"错误：找不到文档 {args.docx}", file=sys.stderr)
        return 2

    # ② 输出路径安全：参数层面的错误，先于授权裁决返回
    if not args.preview and args.out and same_file(args.out, args.docx):
        print("错误：输出路径与输入指向同一文件（不覆盖原件）。", file=sys.stderr)
        return 2

    # ③ 目标画像（含 spec / 模板 / 问卷 / AI / LaTeX 的读取与校验）
    try:
        target = _build_target(args)
    except DocxReadError as e:
        print(f"错误：{e}", file=sys.stderr)
        return 2
    except (ValueError, FileNotFoundError, json.JSONDecodeError) as e:
        print(f"错误：构建目标画像失败：{e}", file=sys.stderr)
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
            print(f"错误：{e}", file=sys.stderr)
            return 2
        if as_json:
            _write_json(json.dumps({"ok": True, "mode": "preview", "changes": changes,
                                    "license": gate}, ensure_ascii=False, indent=2))
            return 0
        print("将进行的格式修正（预览，仅列实际会变的项）：")
        for line in changes:
            print("  · " + line)
        print(f"  试用状态：{gate['message']}")
        return 0

    # ⑤ 已合规 → 不生成副本、不消耗试用额度、不触发门禁
    #    （与 check 同级：单纯读文档比对，无任何副作用，也告诉用户「不用花钱」）
    try:
        changes = compute_changes(args.docx, target)
    except DocxReadError as e:
        print(f"错误：{e}", file=sys.stderr)
        return 2
    if not changes:
        msg = "当前文档已符合目标格式，未生成副本、未消耗试用额度。"
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
        print(f"错误：{e}", file=sys.stderr)
        return 2
    except ValueError as e:
        print(f"错误：{e}", file=sys.stderr)
        return 2
    except OSError as e:
        print(f"错误：{e}", file=sys.stderr)
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
    print(f"✓ 已生成修正稿：{out}")
    print(f"  {after}")
    print("  本次应用：")
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
    print(f"激活：{'是' if s['activated'] else '否'}")
    print(f"种类：{s['kind'] or '-'}")
    print(f"试用已用：{'是' if s['trial_fix_used'] else '否'}（共 {s['trial_limit']} 次）")
    if s.get("revoked"):
        print("状态：已被吊销（需重新激活）")
    print(f"机器指纹：{s['machine']}")
    print(f"本地状态文件：{'正常' if s.get('state_ok') else '缺失/损坏'}")
    return 0


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(description="Thesis Format Doctor Global — offline thesis format checker & fixer")
    # --version 用 argparse 内置 action：解析到就打印并 SystemExit(0)，
    # 不进入任何子命令 —— 版本号读 VERSION 文件（打包时随产物一起带出）。
    p.add_argument("-V", "--version", action="version",
                   version="Thesis Format Doctor Global %s" % app_version())
    sub = p.add_subparsers(dest="cmd")

    g = sub.add_parser("gui", help="启动桌面 GUI")
    g.set_defaults(func=cmd_gui)

    c = sub.add_parser("check", help="命令行体检")
    c.add_argument("docx", help="待检 Word 文档路径")
    c.add_argument("--spec", default="APA", choices=SPEC_ORDER, help="引用规范（默认 APA）")
    c.add_argument("--template", help="学校模板 docx（可选，页面以学校为准）")
    c.add_argument("--ai", help="AI 填好的问卷 JSON（可选）")
    c.add_argument("--questionnaire", help="手填问卷 JSON（可选）")
    c.add_argument("--latex", help="LaTeX 模板路径（可选，第⑤种来源）")
    c.add_argument("--out", help="报告输出路径（md；配 --json 时为 .json）")
    c.add_argument("--json", action="store_true", help="以 JSON 输出结构化结果（便于集成/自动化）")
    c.set_defaults(func=cmd_check)

    f = sub.add_parser("fix", help="一键修正（仅改格式；首次免费，之后需激活）")
    f.add_argument("docx", help="待修正 Word 文档路径")
    f.add_argument("--spec", default="APA", choices=SPEC_ORDER, help="引用规范（默认 APA）")
    f.add_argument("--template", help="学校模板 docx（可选）")
    f.add_argument("--ai", help="AI 填好的问卷 JSON（可选）")
    f.add_argument("--questionnaire", help="手填问卷 JSON（可选）")
    f.add_argument("--latex", help="LaTeX 模板路径（可选）")
    f.add_argument("--out", help="输出路径（默认 原名_fixed.docx，已存在时自动加序号不覆盖）")
    f.add_argument("--preview", action="store_true", help="只预览要改的项，不落盘")
    f.add_argument("--json", action="store_true", help="以 JSON 输出结构化结果（便于集成/自动化）")
    f.set_defaults(func=cmd_fix)

    a = sub.add_parser("activate", help="输入激活码解锁无限修正")
    a.add_argument("code", help="激活码")
    a.add_argument("--offline", action="store_true",
                   help="作为离线激活码处理（卖家签发，无需联网）")
    a.set_defaults(func=cmd_activate)

    s = sub.add_parser("status", help="查看当前授权状态")
    s.set_defaults(func=cmd_status)

    e = sub.add_parser("export-schema", help="导出 AI 填表模板")
    e.add_argument("out", help="输出 JSON 路径")
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
