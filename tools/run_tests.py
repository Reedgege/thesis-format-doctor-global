"""跑全量 pytest，并绕开本机两个环境坑。

为什么需要它（不要直接用 `pytest tests`）：
1. 本机有 safe-delete 钩子，pytest 清理含 >50 文件的 `tmp_path` 时会报
   `SAFE_DELETE_BULK_CONFIRM_REQUIRED` → 每次用**全新 UUID 的 `--basetemp`** 绕开。
2. 本机 PowerShell 捕获不到 python 的 stdout（拿到空串），且 `>`/`Out-File`
   默认写出 UTF-16 → 这里直接把 fd 1/2 `dup2` 到 **UTF-8 文件**，跑完读文件即可。

用法::

    .venv\\Scripts\\python.exe tools\\run_tests.py            # 输出到 %TEMP%\\tfd_pytest_log.txt
    .venv\\Scripts\\python.exe tools\\run_tests.py -o out.txt  # 指定路径
    .venv\\Scripts\\python.exe tools\\run_tests.py -k license  # 透传给 pytest

退出码与 pytest 一致（0 = 全绿）。
"""

from __future__ import annotations

import argparse
import os
import sys
import tempfile
import uuid

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def main() -> int:
    ap = argparse.ArgumentParser(add_help=True)
    ap.add_argument("-o", "--out", default=None)
    ap.add_argument("-k", "--keyword", default=None, help="透传给 pytest 的 -k")
    ap.add_argument("-x", "--exitfirst", action="store_true")
    args, extra = ap.parse_known_args()

    import pytest

    os.chdir(ROOT)
    sys.path.insert(0, ROOT)

    log_path = args.out or os.path.join(tempfile.gettempdir(), "tfd_pytest_log.txt")
    base = os.path.join(tempfile.gettempdir(), "tfd_bt_" + uuid.uuid4().hex[:8])

    argv = ["tests", "-q", "-p", "no:cacheprovider", "--tb=short",
            "--basetemp", base] + extra
    if args.keyword:
        argv += ["-k", args.keyword]
    if args.exitfirst:
        argv += ["-x"]

    # 把 fd 1/2 指向 UTF-8 文件：既避开 PowerShell 的 UTF-16，也留下可读日志。
    # 先备份原 stdout，跑完用它把日志路径回显到真正的终端。
    try:
        orig_stdout = os.dup(1)
    except OSError:
        orig_stdout = None
    fd = os.open(log_path, os.O_WRONLY | os.O_CREAT | os.O_TRUNC)
    os.dup2(fd, 1)
    os.dup2(fd, 2)
    sys.stdout = os.fdopen(1, "w", encoding="utf-8", buffering=1, closefd=False)
    sys.stderr = sys.stdout

    code = pytest.main(argv)
    print(f"\nEXIT={code}\nBASE={base}\nLOG={log_path}")
    sys.stdout.flush()
    if orig_stdout is not None:
        try:
            os.write(orig_stdout, f"LOG={log_path}\nEXIT={code}\n".encode("utf-8"))
            os.close(orig_stdout)
        except OSError:
            pass
    return int(code)


if __name__ == "__main__":
    # 注意：fd 已被重定向，用 os._exit 避免解释器关闭阶段 flush 旧流。
    rc = main()
    try:
        sys.stdout.flush()
    except Exception:  # noqa: BLE001
        pass
    os._exit(rc)
