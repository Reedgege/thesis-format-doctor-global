# -*- coding: utf-8 -*-
"""v2.3.83 · 进程 DPI 感知锁定（老板 2026-10-21 报「保存修正稿时窗口变小、字变清楚」）。

为什么必须用子进程测：``SetProcessDpiAwarenessContext`` 是**进程级一次性设定**，
在 pytest 进程里调一次就把感知锁死了，后续所有 GUI 用例都在"被锁过的进程"里跑，
等于给整个测试套件偷换环境。所以这里一律 ``subprocess`` 起独立解释器。

被测行为（三条）：
  ① 锁定之后进程感知 = UNAWARE（0）；
  ② 锁定之后**任何**提升尝试都失败（返回 False）——这是整个修法的命门：
     原生 shell 对话框就是因为能提升才把我们窗口的拉伸撤掉的，锁住就撤不了；
  ③ ``main()`` 必须在 ``tk.Tk()`` **之前**调用锁定（窗口一旦建好就晚了）。
"""
import os
import re
import subprocess
import sys

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

_SUBPROBE = r"""
import ctypes, sys
sys.path.insert(0, %r)
from src.ui import dpi

u = ctypes.windll.user32
u.SetProcessDpiAwarenessContext.argtypes = [ctypes.c_void_p]
u.SetProcessDpiAwarenessContext.restype = ctypes.c_bool

print("before", dpi.awareness_code())
print("pin", dpi.pin_process_dpi_awareness())
print("after", dpi.awareness_code())
for name, val in (("v2", dpi.DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2),
                  ("gdiscaled", dpi.DPI_AWARENESS_CONTEXT_UNAWARE_GDISCALED),
                  ("system", dpi.DPI_AWARENESS_CONTEXT_SYSTEM_AWARE)):
    ok = u.SetProcessDpiAwarenessContext(ctypes.c_void_p(val))
    print("promote", name, int(bool(ok)), dpi.awareness_code())
""" % REPO


def _run_probe():
    env = dict(os.environ)
    env["PYTHONPATH"] = REPO
    out = subprocess.run([sys.executable, "-c", _SUBPROBE],
                         capture_output=True, text=True, env=env, timeout=120)
    assert out.returncode == 0, out.stderr
    return out.stdout.strip().splitlines()


def test_pin_locks_process_awareness_to_unaware():
    lines = _run_probe()
    d = {}
    for line in lines:
        parts = line.split()
        d[parts[0]] = parts[1:]
    assert d["pin"] == ["UNAWARE"], lines
    assert d["after"] == ["0"], lines          # 0 = UNAWARE


def test_pin_blocks_every_later_promotion():
    """命门：锁定后 V2 / GDIscaled / SYSTEM 三种提升全部失败。"""
    lines = _run_probe()
    promote = [ln.split() for ln in lines if ln.startswith("promote")]
    assert promote, lines
    for _tag, name, ok, code in promote:
        assert ok == "0", "%s 竟然提升成功了：%s" % (name, lines)
        assert code == "0", "%s 之后感知变成了 %s" % (name, code)


def test_main_pins_dpi_before_creating_tk_root():
    """源码顺序断言：锁定必须早于 tk.Tk()，否则窗口已按旧感知建好。"""
    src = open(os.path.join(REPO, "src", "ui", "app.py"), encoding="utf-8").read()
    body = src[src.index("def main():"):]
    i_pin = body.index("pin_process_dpi_awareness()")
    i_tk = body.index("root = tk.Tk()")
    assert i_pin < i_tk, "锁定必须写在 root = tk.Tk() 之前"
    # 别退化成"锁了但没被调用"
    assert re.search(r"^\s*_dpi_mod\.pin_process_dpi_awareness\(\)", body, re.M)


def test_save_flow_snapshots_and_heals_display():
    """保存链路（确认 → 另存为 → 完成）必须前后各抓/还原一次显示状态。"""
    src = open(os.path.join(REPO, "src", "ui", "app.py"), encoding="utf-8").read()
    body = src[src.index("def _run_fix("):src.index("def _save_report(")]
    assert "_snap_display()" in body
    assert body.count("_heal_display(_disp)") >= 4, "每条退出路径都要自愈"
    i_snap = body.index("_snap_display()")
    i_ask = body.index("askyesno(")
    i_done = body.index("_heal_display(_disp)", body.index("showinfo("))
    assert i_snap < i_ask < i_done
