#!/usr/bin/env python3
# 复用脚本：用 OpenAI Codex CLI 对海外版代码做只读审查。
# 用法（在 managed python 下）：
#   python tools/codex_review/run_review.py
# 前置：
#   - codex 二进制已安装（本机路径见 CODEX_EXE）
#   - ~/.codex/config.toml 已配置 deepseek provider + 本项目 trust_level=trusted
#   - 环境变量 DEEPSEEK_API_KEY 已设置
# 注意：脚本以 -s read-only 运行，codex 只读不写，绝不会改动仓库文件。
# 用法补充：可传一个参数指定报告输出路径，例如
#   python run_review.py D:\...\docs\codex_review_xxx.md
import os
import subprocess
import sys

ROOT = r"D:\AgentSpace\thesis-format-doctor-global"
HERE = os.path.dirname(os.path.abspath(__file__))
CODEX_EXE = r"C:\Users\yunwu\AppData\Local\OpenAI\Codex\bin\8e5b6932251c2c1c\codex.exe"
PROMPT_FILE = os.path.join(HERE, "prompt.txt")
# 报告文件名可通过命令行参数覆盖（默认落 round2，避免覆盖历史报告）
REPORT_FILE = (sys.argv[1] if len(sys.argv) > 1
               else os.path.join(ROOT, "docs", "codex_review_2026-09-11_round2.md"))


def main():
    with open(PROMPT_FILE, encoding="utf-8") as f:
        prompt = f.read().strip()

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
                       encoding="utf-8", errors="replace", cwd=ROOT, timeout=600)
    print(f"[codex-review] exit={r.returncode} report={'yes' if os.path.isfile(REPORT_FILE) else 'no'}",
          flush=True)
    if r.stderr.strip():
        print("[codex-review] stderr tail:\n" + r.stderr.strip()[-800:], flush=True)


if __name__ == "__main__":
    main()
