# -*- coding: utf-8 -*-
"""生成 / 刷新跨平台功能验收基准（Windows 侧跑，作为其他平台的比对参照）。

基准内容 = 同一份素材在 Windows 上跑出来的：
    check（修前）→ fix --preview（可修项）→ fix（真改）→ check（修后）

⚠️ 发版后必须重跑一次：verify_cli 只在「实测版本 == 基准版本」时严格比对，
版本一变就自动降成宽松比对（只告警不判红），基准放着不刷新等于失去保护。

用法：
    python tools/crossplat/make_baseline.py [--exe <产物路径>]

不给 --exe 时用源码跑（`python -m src.main`），这也是基准的历史口径。
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import shutil
import subprocess
import sys
import tempfile

for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(os.path.dirname(HERE))
FIXTURE = os.path.join(HERE, "fixtures", "paper_apa_student.docx")
SPEC = "APA-Student"


def _run(exe, args):
    cmd = [exe] + args if exe else [sys.executable, "-m", "src.main"] + args
    p = subprocess.run(cmd, cwd=REPO, capture_output=True)
    out = p.stdout.decode("utf-8", errors="replace")
    start, end = out.find("{"), out.rfind("}")
    if start < 0 or end <= start:
        raise SystemExit("no JSON in output (rc=%s)\nstdout:\n%s\nstderr:\n%s"
                         % (p.returncode, out[:1500],
                            p.stderr.decode("utf-8", errors="replace")[:1500]))
    return json.loads(out[start:end + 1])


def _version(exe) -> str:
    """`--version` 打的是纯文本（"PaperFormat Pro 2.3.122"），单独取。"""
    import re

    cmd = [exe] + ["--version"] if exe else [sys.executable, "-m", "src.main", "--version"]
    p = subprocess.run(cmd, cwd=REPO, capture_output=True)
    text = (p.stdout or p.stderr).decode("utf-8", errors="replace")
    m = re.search(r"(\d+\.\d+\.\d+)", text)
    return m.group(1) if m else ""


def _fp(payload):
    triples, summary = [], payload.get("summary") or {}
    for f in payload.get("findings") or []:
        triples.append([str(f.get("dimension")), str(f.get("severity")),
                        str(f.get("message"))])
    triples.sort()
    return summary, triples


def main(argv=None) -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--exe", default=None, help="产物路径；留空则用源码跑")
    ap.add_argument("--out", default=os.path.join(HERE, "baseline_apa_student.json"))
    args = ap.parse_args(argv)

    work = tempfile.mkdtemp(prefix="baseline_")
    doc_in = os.path.join(work, "input.docx")
    shutil.copyfile(FIXTURE, doc_in)
    version = _version(args.exe)
    before = _run(args.exe, ["check", doc_in, "--spec", SPEC, "--json"])
    b_sum, b_fp = _fp(before)

    prev = _run(args.exe, ["fix", doc_in, "--spec", SPEC, "--preview", "--json"])
    changes = prev.get("changes") or []

    doc_out = os.path.join(work, "fixed.docx")
    fix = _run(args.exe, ["fix", doc_in, "--spec", SPEC, "--out", doc_out, "--json"])
    fx_sum = fix.get("summary") or {}

    after = _run(args.exe, ["check", doc_out, "--spec", SPEC, "--json"])
    a_sum, a_fp = _fp(after)

    data = {
        "note": ("Windows 侧跑出的基准（%s）。macOS / Linux 同版本必须逐条一致；"
                 "不一致 = 跨平台功能性回归。素材结构自洽：题名页 / 摘要 / 多级标题 / "
                 "References 标题 + 6 条条目，格式故意全错。" % ("产物" if args.exe else "源码")),
        "spec": SPEC,
        "app_version": version,
        "produced_on": args.exe or "source (python -m src.main)",
        "fixture": "tools/crossplat/fixtures/paper_apa_student.docx",
        "fixture_sha256": hashlib.sha256(open(FIXTURE, "rb").read()).hexdigest(),
        "required_applied": ["page_number"],
        "min_comments_added": 1,
        "observed_comments_added": fx_sum.get("comments_added"),
        "before": {"summary": b_sum, "findings": b_fp},
        "after": {"summary": a_sum, "findings": a_fp},
        "preview_change_count": len(changes),
        "preview_changes": changes,
        "observed_applied": sorted((fx_sum.get("applied") or {}).keys()),
        "observed_changes": fix.get("changes") or [],
        "soft_errors": fx_sum.get("soft_errors"),
    }
    with open(args.out, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    print("written: %s" % args.out)
    print("  version      : %s" % version)
    print("  before       : %s (%d findings)" % (json.dumps(b_sum, ensure_ascii=False), len(b_fp)))
    print("  after        : %s (%d findings)" % (json.dumps(a_sum, ensure_ascii=False), len(a_fp)))
    print("  preview      : %d 项" % len(changes))
    print("  comments     : %s" % fx_sum.get("comments_added"))
    print("  applied keys : %s" % ", ".join(data["observed_applied"]))
    print("  soft_errors  : %s" % (fx_sum.get("soft_errors") or "无"))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
