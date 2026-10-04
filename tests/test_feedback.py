"""反馈窗逻辑断言（无需 Tk 主窗）。

直接测 ``App._is_valid_email`` / ``App._feedback_deliver`` / ``App._feedback_post``
这几个纯函数 / 静态方法，用 monkeypatch 替换网络与 mailto，断言分支裁决正确。

覆盖理由（2026-10-04 老板定的反馈窗）：
- 邮箱格式边界：空 / 无@ / 无域名点 / 含空格 → 非法；正常 → 合法。
- _feedback_deliver 裁决：邮箱填了但非法 → invalid_email；内容空 → empty；
  网络可达且 ok → sent；否则回落 mailto 兜底 → mailto。
- 邮箱选填：没填邮箱也能 sent（老板要的是「能收到反馈」，回复靠用户填邮箱）。
- _feedback_post：2xx+ok=true → True；2xx+ok=false / 非2xx / 异常 → False（fail-soft）。
"""

from __future__ import annotations

import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from src.ui import app as app_mod

App = app_mod.App


# --------------------------------------------------------------------------- 邮箱校验
def test_is_valid_email_boundaries():
    assert App._is_valid_email("") is False
    assert App._is_valid_email("  ") is False
    assert App._is_valid_email("notanemail") is False
    assert App._is_valid_email("a@b") is False          # 域名无点
    assert App._is_valid_email("a@b@c.com") is False    # 多个 @
    assert App._is_valid_email("a b@c.com") is False    # 含空格
    assert App._is_valid_email("a@b.com") is True
    assert App._is_valid_email("UPPER@B.COM") is True
    assert App._is_valid_email("a.b@c.co.uk") is True
    assert App._is_valid_email("user+tag@example.org") is True


# --------------------------------------------------------------------------- 提交裁决
def test_deliver_invalid_email(monkeypatch):
    """邮箱填了但格式不对 → invalid_email；绝不发网络请求 / 不弹 mailto。"""
    posted, mailed = [], []
    monkeypatch.setattr(App, "_feedback_post", lambda *a, **k: posted.append(a) or True)
    monkeypatch.setattr(App, "_feedback_mailto", lambda *a, **k: mailed.append(a))
    status, _ = App._feedback_deliver("bad@", "hello", "2.3.60")
    assert status == "invalid_email"
    assert posted == [] and mailed == []


def test_deliver_empty_message(monkeypatch):
    """内容必填，空 → empty；不发送。"""
    posted, mailed = [], []
    monkeypatch.setattr(App, "_feedback_post", lambda *a, **k: posted.append(a) or True)
    monkeypatch.setattr(App, "_feedback_mailto", lambda *a, **k: mailed.append(a))
    status, _ = App._feedback_deliver("a@b.com", "   ", "2.3.60")
    assert status == "empty"
    assert posted == [] and mailed == []


def test_deliver_sent(monkeypatch):
    """网络可达且服务端 ok → sent；不落 mailto 兜底。"""
    mailed = []
    monkeypatch.setattr(App, "_feedback_post", lambda *a, **k: True)
    monkeypatch.setattr(App, "_feedback_mailto", lambda *a, **k: mailed.append(a))
    status, _ = App._feedback_deliver("a@b.com", "great app", "2.3.60")
    assert status == "sent"
    assert mailed == []


def test_deliver_email_optional(monkeypatch):
    """邮箱选填：没填邮箱也能 sent（老板要的是收到反馈，回复靠用户填邮箱）。"""
    mailed = []
    monkeypatch.setattr(App, "_feedback_post", lambda *a, **k: True)
    monkeypatch.setattr(App, "_feedback_mailto", lambda *a, **k: mailed.append(a))
    status, _ = App._feedback_deliver("", "found a bug", "2.3.60")
    assert status == "sent"
    assert mailed == []


def test_deliver_mailto_fallback(monkeypatch):
    """网络不可达 / 服务端 ok=false → 回落 mailto 兜底，返回 mailto。"""
    posted_args, mailed = [], []
    monkeypatch.setattr(App, "_feedback_post",
                        lambda *a, **k: posted_args.append(a) or False)
    monkeypatch.setattr(App, "_feedback_mailto", lambda *a, **k: mailed.append(a))
    status, _ = App._feedback_deliver("a@b.com", "offline?", "2.3.60")
    assert status == "mailto"
    assert posted_args and mailed  # 既尝试了 POST，又回落 mailto


# --------------------------------------------------------------------------- 网络 helper
class _FakeResp:
    def __init__(self, status=200, body=b'{"ok":true}'):
        self.status = status
        self._body = body

    def __enter__(self):
        return self

    def __exit__(self, *a):
        return False

    def read(self):
        return self._body


def test_feedback_post_ok_true(monkeypatch):
    monkeypatch.setattr(app_mod.urllib.request, "urlopen",
                        lambda *a, **k: _FakeResp(200, b'{"ok":true}'))
    assert App._feedback_post("a@b.com", "hi", "2.3.60") is True


def test_feedback_post_ok_false(monkeypatch):
    monkeypatch.setattr(app_mod.urllib.request, "urlopen",
                        lambda *a, **k: _FakeResp(200, b'{"ok":false}'))
    assert App._feedback_post("a@b.com", "hi", "2.3.60") is False


def test_feedback_post_non_2xx(monkeypatch):
    monkeypatch.setattr(app_mod.urllib.request, "urlopen",
                        lambda *a, **k: _FakeResp(500, b'{}'))
    assert App._feedback_post("a@b.com", "hi", "2.3.60") is False


def test_feedback_post_exception_fails_soft(monkeypatch):
    def boom(*a, **k):
        raise OSError("offline")
    monkeypatch.setattr(app_mod.urllib.request, "urlopen", boom)
    assert App._feedback_post("a@b.com", "hi", "2.3.60") is False
