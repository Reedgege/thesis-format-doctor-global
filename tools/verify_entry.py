# -*- coding: utf-8 -*-
"""入口行为验证：无参数启动必须进 GUI（这就是「双击打不开」的回归防线）。

背景（2026-09-11，v2.0.2 修）：打包时用了 ``--windows-console-mode=disable``，
双击 exe 不传任何参数；而旧版 ``main()`` 在没有子命令时只 ``print_help()`` 就退出
—— 黑框被隐藏、窗口又不开，用户看到的就是「双击毫无反应」。

本脚本用**带 tkinter 的解释器**真实拉起入口进程，按「进程是否存活」判定有没有
进 GUI（GUI 会常驻直到窗口被关）。无法打包 Nuitka 的机器上，这是最接近真机的验证。

用法（必须用带 tkinter 的 python 跑）：
  python tools/verify_entry.py

退出码 0 = 全部通过。
"""

from __future__ import annotations

import os
import subprocess
import sys
import time

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
ENTRY = os.path.join(ROOT, "main.py")

# 窗口至少要能撑过这么多秒才算「真起来了」；进 GUI 后进程会一直活着
GUI_ALIVE_SECONDS = 5.0


def _probe(args, label):
    """起一个入口进程：超时仍在 → 说明进了 GUI；提前退出 → 打印退出码。"""
    env = os.environ.copy()
    env["PYTHONPATH"] = ROOT + os.pathsep + env.get("PYTHONPATH", "")
    env.setdefault("PYTHONIOENCODING", "utf-8")
    # 隔离授权状态，别污染本机真正的 license 文件
    env["TFD_LICENSE_FILE"] = os.path.join(ROOT, ".verify_entry_license.json")
    env["TFD_SETTINGS_FILE"] = os.path.join(ROOT, ".verify_entry_settings.json")

    t0 = time.time()
    proc = subprocess.Popen(
        [sys.executable, ENTRY] + args, cwd=ROOT, env=env,
        stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
    try:
        out, _ = proc.communicate(timeout=GUI_ALIVE_SECONDS)
        return False, "%.1fs 后退出 rc=%s" % (time.time() - t0, proc.returncode), \
            (out or b"").decode("utf-8", "replace")
    except subprocess.TimeoutExpired:
        proc.kill()
        try:
            out, _ = proc.communicate(timeout=10)
        except Exception:
            out = b""
        return True, "存活超 %.0fs（已进 GUI）" % GUI_ALIVE_SECONDS, \
            (out or b"").decode("utf-8", "replace")


def main() -> int:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(errors="replace")
        except Exception:
            pass

    if not os.path.isfile(ENTRY):
        print("[fail] 找不到入口文件：%s" % ENTRY)
        return 2

    try:
        import tkinter  # noqa: F401
        print("[info] 解释器自带 tkinter，可做真机入口验证：%s" % sys.executable)
    except Exception as e:
        print("[skip] 当前解释器没有 tkinter（%s）——请用带 tk 的 Python 跑本脚本。" % e)
        return 0

    failures = []

    # ① 无参数 = 双击：必须进 GUI（这条是本次 P0 修复的核心防线）
    alive, why, out = _probe([], "no-args")
    print("[test] 无参数启动 -> %s  %s" % ("PASS" if alive else "FAIL", why))
    if not alive:
        failures.append("no-args")
        print("       输出：%s" % out.strip()[:400])

    # ② 显式 gui 子命令：同样必须进 GUI
    alive, why, out = _probe(["gui"], "gui")
    print("[test] `gui` 子命令  -> %s  %s" % ("PASS" if alive else "FAIL", why))
    if not alive:
        failures.append("gui")

    # ③ --help：必须快速退出且 rc=0（CLI 语义没被入口改动破坏）
    env = os.environ.copy()
    env["PYTHONPATH"] = ROOT + os.pathsep + env.get("PYTHONPATH", "")
    p = subprocess.run([sys.executable, ENTRY, "--help"], cwd=ROOT, env=env,
                       capture_output=True, timeout=60)
    ok = p.returncode == 0 and b"usage" in p.stdout.lower()
    print("[test] `--help`      -> %s  rc=%s" % ("PASS" if ok else "FAIL", p.returncode))
    if not ok:
        failures.append("--help")

    # ④ --version：快速退出、打印版本号
    p = subprocess.run([sys.executable, ENTRY, "--version"], cwd=ROOT, env=env,
                       capture_output=True, timeout=60)
    ver = (p.stdout or b"").decode("utf-8", "replace").strip()
    ok = p.returncode == 0 and ver
    print("[test] `--version`   -> %s  rc=%s  %r" % ("PASS" if ok else "FAIL", p.returncode, ver))
    if not ok:
        failures.append("--version")

    # 清理临时状态文件
    for f in (".verify_entry_license.json", ".verify_entry_settings.json"):
        try:
            os.remove(os.path.join(ROOT, f))
        except Exception:
            pass

    print()
    if failures:
        print("[result] 失败项：%s" % ", ".join(failures))
        return 1
    print("[result] 入口验证通过：无参数/gui 均进 GUI，--help 与 --version 语义正常")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
