"""授权被停用（revoked）后，UI 状态行必须与授权状态同步，并给出再次购买的入口。

老板 2026-10-25 报的现象（海外版 PaperFormat Pro）：
    后台把激活码停用 → 客户再点「检查」确实会弹出激活窗，但主界面底部那行仍显示
    "Licensed · unlimited checks"（已授权·不限次），而且主页没有 Upgrade 入口
    —— 客户看到的是「我明明买了」，点下去却被弹窗拦，客户没处再买。

根因（已取证）：
    状态行与 Upgrade 按钮是在**界面搭建时**按本地缓存渲染的（那时心跳还没跑），
    启动核验 / 检查流程发现 revoked 之后**没有再刷新商业区** → UI 与真实授权状态脱节。
    （`i18n` 里其实早就有 `trial_revoked` 文案，但从来没被渲染出来过。）

本文件锁住三件事：
    ① 启动核验判定 locked 时，必须刷新商业区；
    ② `_refresh_commercial` 在 revoked 态：文案 = "已过期或已撤销"，Upgrade 按钮**可见**；
       在 activated 态：文案 = "已授权"，Upgrade 按钮隐藏（付费用户不该被推销升级）；
    ③ i18n 新键 `up_buy_email` 两套语言都在（邮件购买 = 海外站在线结算暂停时的唯一通路）。
"""

from __future__ import annotations

import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

import src.ui.i18n as i18n
from src.ui import app as app_mod


# ---------------------------------------------------------------------------
# 轻量 stub：只具备被测方法用到的成员，避免真正建 Tk 控件
# ---------------------------------------------------------------------------
class _Lbl:
    def __init__(self):
        self.text = None

    def config(self, **kw):
        if "text" in kw:
            self.text = kw["text"]


class _Btn:
    def __init__(self):
        self.visible = None

    def set_visible(self, visible):
        self.visible = visible


class _CommercialStub:
    def __init__(self):
        self._trial_lbl = _Lbl()
        self._upgrade_btn = _Btn()
        self.refreshed = 0

    def tr(self, key, *args):
        s = i18n.STRINGS["en"].get(key, key)
        try:
            return s % args if args else s
        except Exception:
            return s

    def _refresh_commercial(self):
        self.refreshed += 1
        return app_mod.App._refresh_commercial(self)


class _LaunchStub:
    """只具备 `_on_launch_check` 用到的成员。"""

    def __init__(self):
        self.calls = []
        self.version = "2.3.58"
        self.commercial = _CommercialStub()
        self.tr = lambda *a: " ".join(str(x) for x in a)

    def _set_status(self, *a, **k):
        self.calls.append(("status", a, k))

    def _set_bar(self, *a, **k):
        self.calls.append(("bar", a, k))

    def _upgrade_dialog(self, *a, **k):
        self.calls.append(("upgrade", a, k))

    def _show_update_banner(self, *a, **k):
        self.calls.append(("banner", a, k))

    def _refresh_commercial(self):
        self.calls.append(("commercial", (), {}))
        self.commercial.refreshed += 1

    def _on_launch_check(self, res):
        return app_mod.App._on_launch_check(self, res)


# ---------------------------------------------------------------------------
# ① 启动核验：locked → 必须刷新商业区
# ---------------------------------------------------------------------------
def test_launch_locked_refreshes_commercial():
    s = _LaunchStub()
    s._on_launch_check({"locked": True, "network_unreachable": False,
                        "version": {}})
    assert any(c[0] == "commercial" for c in s.calls), \
        "启动核验发现授权被停用时，必须刷新商业区（否则状态行还显示已授权）"
    assert any(c[0] == "upgrade" for c in s.calls), "吊销/过期必须弹激活窗"


def test_launch_valid_does_not_refresh_commercial():
    s = _LaunchStub()
    s._on_launch_check({"locked": False, "network_unreachable": False,
                        "version": {"update_available": False}})
    assert not any(c[0] == "commercial" for c in s.calls), \
        "有效授权时不该乱刷商业区"


# ---------------------------------------------------------------------------
# ② _refresh_commercial 的分档
# ---------------------------------------------------------------------------
def _status_stub(monkeypatch, payload):
    monkeypatch.setattr(app_mod.lic, "status", lambda: payload)


def test_refresh_commercial_revoked_shows_revoked_and_upgrade(monkeypatch):
    _status_stub(monkeypatch, {"activated": False, "revoked": True,
                                "kind": None, "expires_at": None,
                                "trial_fix_used": True, "trial_limit": 1})
    s = _CommercialStub()
    app_mod.App._refresh_commercial(s)
    assert s._trial_lbl.text == i18n.STRINGS["en"]["trial_revoked"], \
        f"停用态文案应为 trial_revoked，实得 {s._trial_lbl.text!r}"
    assert s._upgrade_btn.visible is True, \
        "停用态必须露出 Upgrade 入口（客户要有再次购买的路径）"


def test_refresh_commercial_activated_hides_upgrade(monkeypatch):
    _status_stub(monkeypatch, {"activated": True, "revoked": False,
                                "kind": "lifetime", "expires_at": None,
                                "trial_fix_used": False, "trial_limit": 1})
    s = _CommercialStub()
    app_mod.App._refresh_commercial(s)
    assert s._trial_lbl.text == i18n.STRINGS["en"]["trial_licensed"]
    assert s._upgrade_btn.visible is False, \
        "已付费用户（终身/月卡）不该被推销升级"


def test_refresh_commercial_trial_shows_upgrade(monkeypatch):
    _status_stub(monkeypatch, {"activated": False, "revoked": False,
                                "kind": None, "expires_at": None,
                                "trial_fix_used": False, "trial_limit": 2})
    s = _CommercialStub()
    app_mod.App._refresh_commercial(s)
    assert "2" in (s._trial_lbl.text or ""), "试用态应显示剩余次数"
    assert s._upgrade_btn.visible is True


# ---------------------------------------------------------------------------
# ③ 文案：必须能"重新购买"，且两套语言都有邮件购买入口
# ---------------------------------------------------------------------------
def test_i18n_buy_email_present():
    for lang in ("en", "zh"):
        assert "up_buy_email" in i18n.STRINGS[lang], f"{lang} 缺 up_buy_email"


def test_revoked_copy_mentions_buy_again():
    en = i18n.STRINGS["en"]["trial_revoked"].lower()
    zh = i18n.STRINGS["zh"]["trial_revoked"]
    assert "buy" in en, f"英文停用文案要提到重新购买：{en!r}"
    assert "再次购买" in zh or "购买" in zh, f"中文停用文案要提到购买：{zh!r}"


def test_gate_revoked_copy_mentions_buy_again():
    en = i18n.STRINGS["en"]["gate_revoked"].lower()
    assert "buy" in en, f"英文 gate_revoked 要提到重新购买：{en!r}"
