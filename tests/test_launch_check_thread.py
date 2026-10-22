# -*- coding: utf-8 -*-
"""启动授权核验的线程安全（v2.3.92）：worker 只塞队列，Tk 只从主线程碰。

修的是实测过的偶发崩溃：``_refresh_license_at_launch`` 跑在 worker 线程里，却直接调
``self.root.after()``。Tk 明确不是线程安全的，``after()`` 内部会走 ``tk.createcommand``，
在主循环之外注册命令就会抛::

    RuntimeError: main thread is not in main loop

2026-10-22 全量回归里它**偶发**把后续用例的 setup 打成 ERROR（同一文件单跑 11 passed，
连跑出现 5 errors，再跑又全绿）。这类「跑十次错一次」的时序问题只能靠结构消除：
worker 只 ``queue.put``，主线程用 ``after`` 轮询。

本文件三条契约：
  ① worker 线程**一个 Tk 属性都不许碰**（用「一碰就报错」的假 root 钉死）；
  ② 主线程轮询取到结果后必须立刻停表、只调一次 ``_on_launch_check``；
  ③ 窗口已销毁 / 队列还空时不能炸，也不能无限重排。
"""

from __future__ import annotations

import os
import queue
import sys
import threading

import pytest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from src.ui import app as ui_app


class _ExplodingRoot:
    """任何属性访问都立刻失败 —— 用来证明 worker 线程碰了 Tk。"""

    def __getattr__(self, name):
        raise AssertionError("worker 线程碰到了 Tk：root.%s" % name)


class _RecordingRoot:
    """记录 after() 调用（不真跑事件循环）。"""

    def __init__(self):
        self.calls = []

    def after(self, ms, fn=None, *args):
        self.calls.append((ms, fn))
        return "after#%d" % len(self.calls)


class _FakeApp:
    """借用 App 的三个真方法（不建窗口），其余用最小替身。"""

    _start_launch_check = ui_app.App._start_launch_check
    _poll_launch_check = ui_app.App._poll_launch_check
    _refresh_license_at_launch = ui_app.App._refresh_license_at_launch

    def __init__(self, root):
        self.root = root
        self._closing = False
        self._launch_q = None
        self._launch_polling = False
        self.seen = []
        self.seen_threads = []

    def _on_launch_check(self, res):
        self.seen.append(res)
        self.seen_threads.append(threading.current_thread().name)


@pytest.fixture
def fake_lic(monkeypatch):
    """把中台调用换成固定返回值（不联网、不读本机授权）。"""
    payload = {"locked": False, "version": {"update_available": False}}
    monkeypatch.setattr(ui_app.lic, "refresh_license_on_launch",
                        lambda: dict(payload), raising=True)
    return payload


# ---------------------------------------------------------------------------
# ① worker 线程绝不碰 Tk
# ---------------------------------------------------------------------------

def test_worker_thread_never_touches_tk(fake_lic):
    a = _FakeApp(_ExplodingRoot())
    a._launch_q = queue.Queue()
    # 直接在本线程调也必须是「不碰 root」——碰了会立刻 AssertionError
    a._refresh_license_at_launch()
    assert a._launch_q.get_nowait() == fake_lic


def test_worker_in_real_thread_only_fills_queue(fake_lic):
    a = _FakeApp(_ExplodingRoot())
    a._launch_q = queue.Queue()
    t = threading.Thread(target=a._refresh_license_at_launch, daemon=True)
    t.start()
    t.join(timeout=5)
    assert not t.is_alive()
    assert a._launch_q.get_nowait() == fake_lic
    assert a.seen == []            # worker 不许直接处理 UI


# ---------------------------------------------------------------------------
# ② 主线程轮询：取到结果 → 只调一次 → 停表
# ---------------------------------------------------------------------------

def test_poller_consumes_result_once_and_stops(fake_lic):
    root = _RecordingRoot()
    a = _FakeApp(root)
    a._launch_q = queue.Queue()
    a._launch_polling = True
    a._launch_q.put(dict(fake_lic))

    a._poll_launch_check()

    assert a.seen == [fake_lic]
    assert a.seen_threads == [threading.current_thread().name], "必须在主线程处理"
    assert a._launch_polling is False, "拿到结果后必须停表"
    assert root.calls == [], "拿到结果后不许再排下一轮"


def test_poller_keeps_waiting_while_queue_empty():
    root = _RecordingRoot()
    a = _FakeApp(root)
    a._launch_q = queue.Queue()
    a._launch_polling = True

    a._poll_launch_check()

    assert a.seen == []
    assert a._launch_polling is True
    assert root.calls and root.calls[0][0] == 250, "队列空时应隔 250ms 再看一次"
    assert root.calls[0][1] == a._poll_launch_check


# ---------------------------------------------------------------------------
# ③ 窗口已销毁时不能炸
# ---------------------------------------------------------------------------

def test_poller_stops_when_window_closed():
    root = _RecordingRoot()
    a = _FakeApp(root)
    a._launch_q = queue.Queue()
    a._launch_polling = True
    a._closing = True
    a._launch_q.put({"locked": True})

    a._poll_launch_check()

    assert a._launch_polling is False
    assert a.seen == [], "已关窗时不许再动 UI"
    assert root.calls == []


def test_start_is_idempotent(fake_lic):
    """切语言会重建界面并再次调用；轮询器只许挂一次，线程可以再起。"""
    root = _RecordingRoot()
    a = _FakeApp(root)
    a._start_launch_check()
    a._start_launch_check()
    assert len(root.calls) == 1, "轮询器不许重复挂"


def test_poller_survives_root_after_failure(fake_lic):
    """根治：root 已经销毁时 after() 会抛 TclError，轮询器要安静退出而不是冒泡。"""

    class _DeadRoot:
        def after(self, *a, **kw):
            raise RuntimeError("application has been destroyed")

    a = _FakeApp(_DeadRoot())
    a._launch_q = queue.Queue()
    a._launch_polling = True
    a._poll_launch_check()          # 不抛异常即通过
    assert a._launch_polling is False
