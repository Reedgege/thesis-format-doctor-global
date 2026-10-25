"""启动一次性核验（授权 + 版本）的 UI 状态机断言（无需 Tk 主窗）。

直接调用 ``App._on_launch_check``（决策逻辑在 UI 线程内执行，但本身
不依赖 Tk 控件），用一个轻量 stub 记录它调用了哪些分支。
同时锁住 i18n 新文案在两套语言里都存在。
"""

from __future__ import annotations

import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

import src.ui.i18n as i18n
from src.ui import app as app_mod


# ---------------------------------------------------------------------------
# i18n：新文案两套语言都必须存在
# ---------------------------------------------------------------------------
def test_i18n_launch_keys_present():
    for key in ("launch_offline", "upd_title", "upd_body", "upd_btn", "upd_later"):
        assert key in i18n.STRINGS["en"], f"EN missing {key}"
        assert key in i18n.STRINGS["zh"], f"ZH missing {key}"


# ---------------------------------------------------------------------------
# 状态机：用一个 stub 记录 _on_launch_check 走了哪条分支
# ---------------------------------------------------------------------------
class _LaunchStub:
    """只具备 _on_launch_check 用到的成员，避免真正建 Tk 控件。"""
    def __init__(self):
        self.calls = []
        self.version = "2.3.58"
        # tr 原样把参数拼成字符串，方便断言
        self.tr = lambda *a: " ".join(str(x) for x in a)

    def _set_status(self, *a, **k):
        self.calls.append(("status", a, k))

    def _set_bar(self, *a, **k):
        self.calls.append(("bar", a, k))

    def _upgrade_dialog(self, *a, **k):
        self.calls.append(("upgrade", a, k))

    def _refresh_commercial(self, *a, **k):
        # 启动核验判定吊销时必须刷新商业区（老板 2026-10-25 报的状态行不同步）
        self.calls.append(("commercial", a, k))

    def _show_update_banner(self, *a, **k):
        self.calls.append(("banner", a, k))

    def _on_launch_check(self, res):
        # 以未绑定方法方式调用真实实现
        return app_mod.App._on_launch_check(self, res)


def _has(calls, name):
    return any(c[0] == name for c in calls)


def test_launch_state_revoked_opens_activate_window():
    s = _LaunchStub()
    s._on_launch_check({"locked": True, "network_unreachable": False,
                         "version": {}})
    assert _has(s.calls, "upgrade"), "吊销/过期必须弹激活窗"
    assert _has(s.calls, "status"), "吊销/过期必须刷新状态行"


def test_launch_state_offline_shows_prompt():
    s = _LaunchStub()
    s._on_launch_check({"locked": False, "network_unreachable": True,
                         "version": {"update_available": False}})
    assert _has(s.calls, "status"), "网络不可达必须给软提示"
    assert not _has(s.calls, "upgrade"), "离线宽延绝不该锁/弹激活窗"
    assert not _has(s.calls, "banner")


def test_launch_state_update_shows_banner():
    s = _LaunchStub()
    s._on_launch_check({"locked": False, "network_unreachable": False,
                         "version": {"update_available": True,
                                     "latest_version": "9.9.9",
                                     "download_url": "https://x"}})
    assert _has(s.calls, "banner"), "有新版本必须弹软横幅"
    assert not _has(s.calls, "upgrade")


def test_launch_state_valid_latest_is_silent():
    s = _LaunchStub()
    s._on_launch_check({"locked": False, "network_unreachable": False,
                         "version": {"update_available": False}})
    # 有效 + 最新 = 真静默：什么 UI 都不触发
    assert s.calls == [], f"有效+最新应真静默，却触发了 {s.calls}"


def test_launch_state_revoked_suppresses_banner():
    s = _LaunchStub()
    s._on_launch_check({"locked": True, "network_unreachable": False,
                         "version": {"update_available": True,
                                     "latest_version": "9.9.9",
                                     "download_url": "https://x"}})
    # 即便有更新，失效状态优先弹激活窗，不应再弹更新横幅
    assert _has(s.calls, "upgrade")
    assert not _has(s.calls, "banner")
