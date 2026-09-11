#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""用 Codex CLI 对海外版「本轮改动」做只读审查（round3）。

与 run_review.py 的区别：prompt 与报告路径可指定，便于每轮单独留档。

用法：
  python tools/codex_review/run_round3.py                      # 用默认 round3 prompt/report
  python tools/codex_review/run_round3.py <prompt> <report>    # 自定义

前置：~/.codex/config.toml 已配 deepseek provider；环境变量 DEEPSEEK_API_KEY 已设置。
安全：以 `-s read-only` 运行，codex 只读，不会改动仓库文件。
"""
import os
import subprocess
import sys

ROOT = r"D:\AgentSpace\thesis-format-doctor-global"
HERE = os.path.dirname(os.path.abspath(__file__))
CODEX_EXE = r"C:\Users\yunwu\AppData\Local\OpenAI\Codex\bin\8e5b6932251c2c1c\codex.exe"

PROMPT_FILE = (sys.argv[1] if len(sys.argv) > 1 else os.path.join(HERE, "prompt_round3.txt"))
REPORT_FILE = (sys.argv[2] if len(sys.argv) > 2
               else os.path.join(ROOT, "docs", "codex_review_2026-09-11_round3.md"))


def main() -> int:
    if not os.path.isfile(CODEX_EXE):
        print("[codex-review] codex.exe not found: " + CODEX_EXE, flush=True)
        return 2
    if not os.environ.get("DEEPSEEK_API_KEY"):
        print("[codex-review] DEEPSEEK_API_KEY not set", flush=True)
        return 2

    with open(PROMPT_FILE, encoding="utf-8") as fh:
        prompt = fh.read().strip()
    print("[codex-review] prompt=%s (%d chars)" % (os.path.basename(PROMPT_FILE), len(prompt)),
          flush=True)
    print("[codex-review] report=%s" % REPORT_FILE, flush=True)

    cmd = [
        CODEX_EXE, "exec", prompt,
        "-C", ROOT,
        "-s", "read-only",
        "--skip-git-repo-check",
        "-o", REPORT_FILE,
        "--ephemeral",
        "--color", "never",
    ]
    print("[codex-review] launching read-only review (no file changes) ...", flush=True)
    r = subprocess.run(cmd, capture_output=True, text=True,
                       encoding="utf-8", errors="replace", cwd=ROOT, timeout=3600)
    print("[codex-review] exit=%d report=%s"
          % (r.returncode, "yes" if os.path.isfile(REPORT_FILE) else "no"), flush=True)
    if r.stdout and r.stdout.strip():
        print("=== codex stdout tail ===\n" + r.stdout.strip()[-2000:], flush=True)
    if r.stderr and r.stderr.strip():
        print("=== codex stderr tail ===\n" + r.stderr.strip()[-2000:], flush=True)
    if os.path.isfile(REPORT_FILE):
        print("[codex-review] report size=%d bytes" % os.path.getsize(REPORT_FILE), flush=True)
    return 0 if os.path.isfile(REPORT_FILE) else 1


if __name__ == "__main__":
    raise SystemExit(main())
