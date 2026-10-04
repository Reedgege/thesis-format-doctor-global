"""一键修正门禁的 UI 状态机断言（无需 Tk 主窗）。

直接调用 ``App._run_fix``，用 stub 记录它走了哪条分支。

重点回归（2026-10-04）：授权不足时点击「修改」必须弹出合并激活窗引导购买。
旧实现在 ``try`` 内 early-return，导致下方的「授权不足」分支（弹激活窗）永远
到不了，用户只看到红状态栏却找不到买码入口。该 bug 同时影响试用路径与
正式（已激活却被吊销/过期/需重新激活）路径 —— 因为两段走的是同一个 _run_fix。

覆盖理由：
- trial_exhausted：试用次数用尽
- revoked：激活码被停用（中台心跳 status=revoked/expired 都归到这个 reason，
           故「过期」一并覆盖）
- needs_reactivation：需重新激活
- state_error：程序故障，走普通警告而非购买窗（不应弹窗）
- allowed：正常进入预览/确认，不应弹购买窗
"""

from __future__ import annotations

import os
import sys
import types

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

import pytest
from src.ui import app as app_mod


# 这些 reason 代表「该引导购买/重新激活」，必须弹合并激活窗
BUY_REASONS = ("trial_exhausted", "revoked", "needs_reactivation")


class _FixStub:
    """只具备 _run_fix 用到的成员，避免真正建 Tk 控件。"""
    def __init__(self):
        self.calls = []
        self._busy = False
        self._phase = "ready"
        self._last_issues = 0
        self._last_output = None
        self.lang = "en"
        # docx_path.get() 返回一个真实存在的文件，跳过「未选论文」分支
        self.docx_path = types.SimpleNamespace(
            get=lambda: os.path.abspath(__file__))
        # tr 原样把参数拼成字符串，方便断言
        self.tr = lambda *a: " ".join(str(x) for x in a)

    def _set_status(self, *a, **k):
        self.calls.append(("status", a, k))

    def _set_bar(self, *a, **k):
        self.calls.append(("bar", a, k))

    def _set_busy(self, v, *a, **k):
        self.calls.append(("busy", v, k))

    def _set_report(self, *a, **k):
        self.calls.append(("report", a, k))

    def _fail(self, *a, **k):
        self.calls.append(("fail", a, k))

    def _sync_ui(self, *a, **k):
        self.calls.append(("sync_ui", a, k))

    def _flush_frame(self, *a, **k):
        self.calls.append(("flush", a, k))

    def _gate_message(self, gate):
        return "GATE:" + str((gate or {}).get("reason"))

    def _upgrade_dialog(self, *a, **k):
        self.calls.append(("upgrade", a, k))

    def _build_target(self, *a, **k):
        return object()

    def _run_fix(self):
        # 以未绑定方法方式调用真实实现
        return app_mod.App._run_fix(self)


def _has(calls, name):
    return any(c[0] == name for c in calls)


def _patch_gates(monkeypatch, reason, allowed=False):
    monkeypatch.setattr(app_mod.lic, "require_fix_entitlement",
                        lambda: {"allowed": allowed, "reason": reason})
    # 返回非空改动，确保不会走「已合规」早退（那条分支本就不该弹激活窗）
    monkeypatch.setattr(app_mod, "compute_changes", lambda *a, **k: ["x"])


@pytest.mark.parametrize("reason", BUY_REASONS)
def test_fix_gate_disallowed_pops_activate_window(monkeypatch, reason):
    """授权不足（试用耗尽 / 吊销 / 过期 / 需重新激活）点击修改必须弹激活窗。"""
    s = _FixStub()
    _patch_gates(monkeypatch, reason)
    # 普通警告在此分支不应被调用（防止误走到 showwarning）
    warns = []
    monkeypatch.setattr(app_mod.messagebox, "showwarning",
                        lambda *a, **k: warns.append(a))
    s._run_fix()
    assert _has(s.calls, "upgrade"), f"reason={reason} 必须弹激活窗引导购买"
    assert _has(s.calls, "status"), f"reason={reason} 必须刷新状态行"
    assert not warns, f"reason={reason} 不应走普通警告（应弹购买窗）"


def test_fix_gate_state_error_uses_warning_not_upgrade(monkeypatch):
    """state_error 是程序故障，走普通警告，不应弹购买窗。"""
    s = _FixStub()
    warns = []
    monkeypatch.setattr(app_mod.messagebox, "showwarning",
                        lambda *a, **k: warns.append(a))
    _patch_gates(monkeypatch, "state_error")
    s._run_fix()
    assert not _has(s.calls, "upgrade"), "state_error 不应弹购买窗"
    assert warns, "state_error 应走普通警告"


def test_fix_gate_allowed_reaches_confirm(monkeypatch):
    """授权通过应进入预览/确认，不弹激活窗、不再走购买流程。"""
    s = _FixStub()
    monkeypatch.setattr(app_mod.messagebox, "askyesno",
                        lambda *a, **k: False)  # 用户取消，不应落盘
    _patch_gates(monkeypatch, "trial", allowed=True)
    s._run_fix()
    assert _has(s.calls, "report"), "allowed 应进入预览"
    assert not _has(s.calls, "upgrade"), "allowed 不应弹激活窗"
