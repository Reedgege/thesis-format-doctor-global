"""license_summary() 阈值逻辑断言测试（纯离线，不依赖 tkinter / 中台）。

锁住四档授权的展示摘要：
- times 次卡：剩余次数、≤1 次 warn、用尽 exhausted；
- quarterly / annual 定时卡：剩余天数、≤7 天 warn、过期 exhausted；
- lifetime：无到期、不 warn / 不 exhausted。
"""

import time

import pytest

from src.license import license as lic


def _state(kind="lifetime", uses_total=None, uses_used=0, expires_at=None,
           activated=True, owed=0):
    base = lic._default_state()
    base.update({
        "activated": activated,
        "kind": kind,
        "uses_total": uses_total,
        "uses_used": uses_used,
        "expires_at": expires_at,
        "owed": owed,
    })
    return base


@pytest.fixture
def patch_state(monkeypatch):
    def _set(st):
        monkeypatch.setattr(lic, "load_state", lambda: st)
    return _set


def test_times_remaining_and_warn(patch_state):
    patch_state(_state(kind=lic.KIND_TIMES, uses_total=2, uses_used=1))
    s = lic.license_summary()
    assert s["uses_left"] == 1
    assert s["warn"] is True
    assert s["exhausted"] is False
    assert s["kind_label"] == "Two-use pack"


def test_times_exhausted(patch_state):
    patch_state(_state(kind=lic.KIND_TIMES, uses_total=2, uses_used=2))
    s = lic.license_summary()
    assert s["uses_left"] == 0
    assert s["exhausted"] is True


def test_times_normal_no_warn(patch_state):
    patch_state(_state(kind=lic.KIND_TIMES, uses_total=5, uses_used=2))
    s = lic.license_summary()
    assert s["uses_left"] == 3
    assert s["warn"] is False
    assert s["exhausted"] is False


def test_quarterly_warn_within_7_days(patch_state):
    now = int(time.time() * 1000)
    patch_state(_state(kind=lic.KIND_QUARTERLY, expires_at=now + 5 * 86400000))
    s = lic.license_summary()
    assert s["days_left"] == 5
    assert s["warn"] is True
    assert s["exhausted"] is False


def test_quarterly_expired(patch_state):
    now = int(time.time() * 1000)
    patch_state(_state(kind=lic.KIND_QUARTERLY, expires_at=now - 1000))
    s = lic.license_summary()
    assert s["exhausted"] is True


def test_annual_normal(patch_state):
    now = int(time.time() * 1000)
    patch_state(_state(kind=lic.KIND_ANNUAL, expires_at=now + 200 * 86400000))
    s = lic.license_summary()
    assert s["warn"] is False
    assert s["exhausted"] is False


def test_lifetime_no_expiry(patch_state):
    patch_state(_state(kind=lic.KIND_LIFETIME))
    s = lic.license_summary()
    assert s["days_left"] is None
    assert s["exhausted"] is False
    assert s["warn"] is False
