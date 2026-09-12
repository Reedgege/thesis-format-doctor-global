"""授权模块断言式测试（pytest，完全离线：三个中台调用被 monkeypatch 掉）。

覆盖：
- 未激活首次免费试用允许；
- 试用一次后 record_fix_used → 第 2 次要求激活；
- 激活码通过中台验证后写入本地状态 → 允许（终身）；
- 中台验证失败 → 不激活；
- 中台不可达 → 不崩溃（返回错误信息）；
- 状态文件损坏 → fail-closed（按试用已用处理）；
- 状态文件来自其他机器 → 不认激活；
- record_fix_used 在写盘失败时也不抛异常。
"""

from __future__ import annotations

import json
import os
import sys

import base64
import time

import pytest
from cryptography.hazmat.primitives.asymmetric.ed25519 import Ed25519PrivateKey

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from src.license import license as lic
from src.license import crypto as crypto_mod


# 离线程测试用临时 Ed25519 密钥对：真实私钥只在卖家本机，CI/单测都不持有。
# 这里生成一份临时密钥对——公钥（32 字节原始值）注入 crypto.verify_offline_code 的
# 校验路径，私钥用于签发测试离线码，闭环完全离线、不依赖任何密钥文件。
@pytest.fixture
def offline_keys(monkeypatch):
    priv = Ed25519PrivateKey.generate()
    crypto_mod.set_verify_public_key(priv.public_key().public_bytes_raw())
    yield priv
    crypto_mod.set_verify_public_key(None)


def _sign(priv, machine_code):
    """测试用离线码签名器（真实签名只由卖家统一发码器 reedcode_unified.py 完成）。"""
    payload = "%s|%d" % (machine_code, int(time.time()))
    sig = priv.sign(payload.encode("utf-8"))          # 64 字节 Ed25519 签名
    return payload + "." + base64.urlsafe_b64encode(sig).decode("ascii")


@pytest.fixture
def tmp_license(tmp_path, monkeypatch):
    """把授权状态文件指到临时路径 + 拦截三个中台调用（完全离线）。"""
    p = tmp_path / "license.json"
    monkeypatch.setattr(lic, "LICENSE_FILE", p)
    monkeypatch.delenv("TFD_LICENSE_FILE", raising=False)
    # 默认：中台不可达（返回 None 表示离线宽限），各测试按需覆盖
    monkeypatch.setattr(lic, "_safe_trial_sync", lambda machine, claim=False: None)
    monkeypatch.setattr(lic, "_safe_heartbeat", lambda code, machine: None)
    return p


def test_first_trial_allowed(tmp_license):
    g = lic.require_fix_entitlement()
    assert g["allowed"] is True
    assert g["reason"] == "trial"


def test_after_fix_trial_exhausted(tmp_license):
    assert lic.record_fix_used() is True
    g = lic.require_fix_entitlement()
    assert g["allowed"] is False
    assert g["reason"] == "trial_exhausted"
    assert "激活码" in g["message"]


def test_activate_success_then_allowed(tmp_license, monkeypatch):
    def fake_verify(code):
        return {"ok": True, "type": "lifetime", "status": "active"}
    monkeypatch.setattr(lic, "_check_mid_platform", fake_verify)
    r = lic.activate("TEST-CODE-123")
    assert r["ok"] is True
    assert r["kind"] == "lifetime"
    g = lic.require_fix_entitlement()
    assert g["allowed"] is True
    assert g["reason"] == "activated"


def test_activate_fail_then_no_change(tmp_license, monkeypatch):
    def fake_verify(code):
        return {"ok": False, "error": "invalid_code"}
    monkeypatch.setattr(lic, "_check_mid_platform", fake_verify)
    r = lic.activate("BAD-CODE")
    assert r["ok"] is False
    # 激活失败不应污染试用额度
    g = lic.require_fix_entitlement()
    assert g["reason"] == "trial"
    assert g["allowed"] is True


def test_activate_revoked_message(tmp_license, monkeypatch):
    monkeypatch.setattr(lic, "_check_mid_platform",
                        lambda code: {"ok": False, "error": "revoked"})
    r = lic.activate("REVOKED-CODE")
    assert r["ok"] is False
    assert "吊销" in r["message"]


def test_mid_platform_unreachable_does_not_crash(tmp_license, monkeypatch):
    def fake_verify(code):
        raise ConnectionError("离线")
    monkeypatch.setattr(lic, "_check_mid_platform", fake_verify)
    r = lic.activate("ANY-CODE")
    assert r["ok"] is False
    assert "离线" in r["message"]


def test_status_shape(tmp_license):
    s = lic.status()
    assert s["activated"] is False
    assert s["trial_fix_used"] is False
    assert s["trial_limit"] == 1
    assert isinstance(s["machine"], str) and len(s["machine"]) > 0


# ---------------------------------------------------------------------------
# 防重置（FR6）
# ---------------------------------------------------------------------------

def test_corrupt_state_fails_closed(tmp_license):
    """状态文件损坏 → 按「试用已用」处理，不给白嫖窗口。"""
    tmp_license.write_text("{ this is not json", encoding="utf-8")
    g = lic.require_fix_entitlement()
    assert g["allowed"] is False
    assert g["reason"] == "trial_exhausted"


def _sealed_state(**fields):
    """经 ``save_state`` 落盘一份「合法状态文件」（带完整性封印）。

    新安全模型下，手写 JSON（无封印）会被判为未激活态，所以凡是
    「模拟一个正常存在的激活状态」的测试都必须走这条路。
    """
    st = lic._default_state()
    st.update(fields)
    lic.save_state(st)
    return st


def test_copied_state_from_other_machine_not_activated(tmp_license):
    """拷贝别的机器的 license.json（激活态）过来 → 不认激活。"""
    _sealed_state(activated=True, code="SOME-CODE", kind="lifetime",
                  trial_fix_used=True, machine="ffffffffffffffff",
                  last_heartbeat=0)
    g = lic.require_fix_entitlement()
    assert g["allowed"] is False
    assert g["reason"] == "trial_exhausted"


def test_machine_mismatch_rebinds_when_online(tmp_license, monkeypatch):
    """机器指纹变化（改机名）且在线心跳 active → 自动重绑，不误锁用户。"""
    _sealed_state(activated=True, code="MY-CODE", kind="lifetime",
                  trial_fix_used=True, machine="ffffffffffffffff",
                  last_heartbeat=0)
    monkeypatch.setattr(lic, "_safe_heartbeat",
                        lambda code, machine: {"ok": True, "status": "active"})
    g = lic.require_fix_entitlement()
    assert g["allowed"] is True
    assert g["reason"] == "activated"


def test_server_revoked_locks(tmp_license, monkeypatch):
    """中台明确返回 revoked → 锁死（心跳超期时才查）。"""
    _sealed_state(activated=True, code="MY-CODE", kind="lifetime",
                  trial_fix_used=True, machine=lic._machine_fingerprint(),
                  last_heartbeat=0)        # 远超 3 天宽限
    monkeypatch.setattr(lic, "_safe_heartbeat",
                        lambda code, machine: {"ok": True, "status": "revoked"})
    g = lic.require_fix_entitlement()
    assert g["allowed"] is False
    assert g["reason"] == "revoked"


def test_server_trial_used_locks(tmp_license, monkeypatch):
    """本地被删/丢失，但中台记得该机器已用过试用 → 依然锁。"""
    monkeypatch.setattr(lic, "_safe_trial_sync",
                        lambda machine, claim=False: {"ok": True, "trial_used": True})
    g = lic.require_fix_entitlement()
    assert g["allowed"] is False
    assert g["reason"] == "trial_exhausted"


def test_record_fix_used_never_raises(tmp_license, monkeypatch):
    """写盘失败（如家目录只读）时不得抛异常，否则会「已生成但未计次」。"""
    def boom(state):
        raise OSError("read-only file system")
    monkeypatch.setattr(lic, "save_state", boom)
    assert lic.record_fix_used() is False


def test_no_network_endpoints_stale():
    """回归：中台真实路由必须与本地共识一致（防止再退回不存在的 /api/auth/*）。"""
    assert lic._P_ACTIVATE == "/api/activate"
    assert lic._P_HEARTBEAT == "/api/heartbeat"
    assert lic._P_TRIAL == "/api/trial"
    assert "/api/auth" not in lic._P_ACTIVATE + lic._P_HEARTBEAT


# ---------------------------------------------------------------------------
# 兜底方案：离线备用码（Ed25519 非对称；与统一发码器 reedcode_unified.py 共用算法；与导师版密钥隔离）
# ---------------------------------------------------------------------------

def test_offline_code_roundtrip(offline_keys):
    """sign -> verify 闭环：本机机器码生成可校验通过。"""
    mc = lic._machine_fingerprint()
    code = _sign(offline_keys, mc)
    assert code.count(".") == 1, "离线码格式应为 payload.signature"
    assert lic.verify_offline_code(code, mc) is True


def test_offline_code_rejects_wrong_machine_and_tamper(offline_keys):
    """离线码绑定机器码：错机器 / 篡改签名均拒。"""
    mc = lic._machine_fingerprint()
    code = _sign(offline_keys, mc)
    assert lic.verify_offline_code(code, "WRONG-MACHINE") is False
    payload, _, sig = code.rpartition(".")
    tampered = payload + "." + ("A" if sig[-1] != "A" else "B")
    assert lic.verify_offline_code(tampered, mc) is False


def test_activate_offline_persists(tmp_license, offline_keys):
    """离线激活写盘后：状态含 offline=True 且持久化（回归：offline 曾漏进 _default_state
    导致 save/load 被剥离、标志丢失）。"""
    mc = lic._machine_fingerprint()
    code = _sign(offline_keys, mc)
    r = lic.activate_offline(code)
    assert r["ok"] is True
    assert r["kind"] == "lifetime"
    # 重新读取状态文件，确认 offline 标记确实落盘并被还原
    st = lic.load_state()
    assert st.get("offline") is True
    assert st.get("activated") is True
    assert st.get("kind") == "lifetime"


def test_activate_offline_rejects_other_machine(tmp_license, offline_keys):
    """用别的机器码生成的离线码，在本机激活必须失败。"""
    code = _sign(offline_keys, "OTHER-MACHINE-CODE")
    r = lic.activate_offline(code)
    assert r["ok"] is False
    assert "机器码" in r["message"]


def test_offline_activation_grants_fix_without_network(tmp_license, monkeypatch, offline_keys):
    """离线激活后：fix 授权放行，且全程不触网（心跳/试用同步均被拦截）。"""
    mc = lic._machine_fingerprint()
    code = _sign(offline_keys, mc)
    assert lic.activate_offline(code)["ok"] is True
    # 即便中台全部不可达，离线授权也应直接放行
    monkeypatch.setattr(lic, "_safe_heartbeat", lambda code, machine: None)
    g = lic.require_fix_entitlement()
    assert g["allowed"] is True
    assert g["reason"] == "activated"
    s = lic.status()
    assert s["offline"] is True


# ---------------------------------------------------------------------------
# 本地状态完整性：封印（防「改两行 JSON 白嫖」）
# ---------------------------------------------------------------------------

def test_handwritten_json_cannot_unlock_offline(tmp_license, monkeypatch):
    """真风险回归：手写 {"activated":true,"offline":true} → 必须**不能**白嫖 fix。

    这是新安全模型的核心断言：授权裁决不再盲信本地布尔字段，而是重新验签离线码。
    """
    monkeypatch.setattr(lic, "_safe_heartbeat", lambda code, machine: None)
    monkeypatch.setattr(lic, "_safe_trial_sync", lambda machine, claim=False: None)

    # ① 完全手写的 JSON（无封印）
    tmp_license.write_text(json.dumps(
        {"activated": True, "offline": True, "trial_fix_used": False}),
        encoding="utf-8")
    g = lic.require_fix_entitlement()
    assert g["allowed"] is False, "手写 JSON 竟然放行了 fix"

    # ② 结论：不仅不放行，而且明确要求「重新激活」——自称已激活却拿不出有效码，
    #    既不能放行、也不该用「试用已用完」这种误导性文案（那是给正常试用用户的）。
    assert g["reason"] == "needs_reactivation"


def test_unsealed_without_activation_stays_fail_closed(tmp_license, monkeypatch):
    """删掉 _sig 且未激活 → 按试用已用（fail-closed），与代码注释口径一致。"""
    monkeypatch.setattr(lic, "_safe_trial_sync", lambda machine, claim=False: None)
    tmp_license.write_text(json.dumps({
        "activated": False, "trial_fix_used": True, "offline": False,
        "machine": lic._machine_fingerprint(),
    }), encoding="utf-8")
    st = lic.load_state()
    assert st["_state_reason"] == "unsealed"
    assert st["_state_ok"] is False
    g = lic.require_fix_entitlement()
    assert g["allowed"] is False
    assert g["reason"] == "trial_exhausted"


# ---------------------------------------------------------------------------
# 旧版状态文件（v2.0.1 及更早：save_state 不写 _sig）的可验证迁移
# 背景：升级后被判 unsealed → 已付费用户被告知"试用已用完"，是明确的用户事故。
# ---------------------------------------------------------------------------

def _legacy_state(**fields):
    """写一份**无封印**的状态文件，模拟 v2.0.1 的落盘格式。"""
    st = {k: v for k, v in lic._default_state().items()}
    st.update(fields)
    st.pop("_sig", None)
    return st


def test_legacy_offline_state_migrates_without_network(tmp_license, monkeypatch, offline_keys):
    """旧版离线激活用户：升级后自动迁移（本地验签即可，全程不联网）。"""
    mc = lic._machine_fingerprint()
    code = _sign(offline_keys, mc)
    tmp_license.write_text(json.dumps(_legacy_state(
        activated=True, code=code, kind="lifetime", offline=True,
        trial_fix_used=True, machine=mc, last_heartbeat=0)), encoding="utf-8")

    def _boom(*a, **k):
        raise AssertionError("离线迁移不该联网")
    monkeypatch.setattr(lic, "_safe_heartbeat", _boom)
    monkeypatch.setattr(lic, "_safe_trial_sync", _boom)

    g = lic.require_fix_entitlement()
    assert g["allowed"] is True, "旧版离线付费用户升级后被误锁"
    assert g["reason"] == "activated"
    # 迁移后必须已重封印：再次读取是 ok，且再跑一次裁决仍然放行
    assert lic.load_state()["_state_reason"] == "ok"
    assert lic.require_fix_entitlement()["allowed"] is True


def test_legacy_online_state_migrates_via_heartbeat(tmp_license, monkeypatch):
    """旧版在线激活用户：升级后靠一次中台心跳自证，随后重封印。"""
    tmp_license.write_text(json.dumps(_legacy_state(
        activated=True, code="ONLINE-CODE", kind="lifetime", offline=False,
        trial_fix_used=True, machine=lic._machine_fingerprint(),
        last_heartbeat=0)), encoding="utf-8")
    monkeypatch.setattr(lic, "_safe_heartbeat",
                        lambda code, machine: {"ok": True, "status": "active"})
    g = lic.require_fix_entitlement()
    assert g["allowed"] is True
    assert g["reason"] == "activated"
    assert lic.load_state()["_state_reason"] == "ok"


def test_legacy_state_with_bogus_code_is_rejected(tmp_license, monkeypatch):
    """伪造的旧版激活状态（无有效码）→ 不迁移、不放行（删 _sig 不是后门）。"""
    monkeypatch.setattr(lic, "_safe_heartbeat", lambda code, machine: None)
    monkeypatch.setattr(lic, "_safe_trial_sync", lambda machine, claim=False: None)
    tmp_license.write_text(json.dumps(_legacy_state(
        activated=True, code="FAKE-CODE", kind="lifetime", offline=True,
        trial_fix_used=False, machine=lic._machine_fingerprint(),
        last_heartbeat=0)), encoding="utf-8")
    g = lic.require_fix_entitlement()
    assert g["allowed"] is False
    assert g["reason"] == "needs_reactivation"


def test_legacy_revoked_online_state_is_not_migrated(tmp_license, monkeypatch):
    """旧版码已被中台吊销 → 不予迁移（迁移不是绕过吊销的后门）。"""
    tmp_license.write_text(json.dumps(_legacy_state(
        activated=True, code="REVOKED-CODE", kind="lifetime", offline=False,
        trial_fix_used=True, machine=lic._machine_fingerprint(),
        last_heartbeat=0)), encoding="utf-8")
    monkeypatch.setattr(lic, "_safe_heartbeat",
                        lambda code, machine: {"ok": True, "status": "revoked"})
    g = lic.require_fix_entitlement()
    assert g["allowed"] is False
    assert g["reason"] == "needs_reactivation"


def test_state_seal_detects_tampering(tmp_license, monkeypatch):
    """带封印的文件被改字段 → 封印失配 → 视为篡改，不放行。"""
    _sealed_state(activated=False, trial_fix_used=True, offline=False)
    raw = json.loads(tmp_license.read_text(encoding="utf-8"))
    assert "_sig" in raw, "save_state 必须写封印"
    raw["trial_fix_used"] = False       # 试图把「已用」改回「未用」
    tmp_license.write_text(json.dumps(raw), encoding="utf-8")
    monkeypatch.setattr(lic, "_safe_trial_sync", lambda machine, claim=False: None)

    assert lic.load_state()["_state_reason"] == "tampered"
    assert lic.require_fix_entitlement()["allowed"] is False


def test_activated_state_without_seal_is_rejected(tmp_license):
    """旧格式（无 _sig）的「已激活」状态 → 不认（否则删掉 _sig 就是后门）。"""
    tmp_license.write_text(json.dumps({
        "activated": True, "code": "ANY", "offline": True,
        "trial_fix_used": True, "machine": lic._machine_fingerprint(),
        "last_heartbeat": 0,
    }), encoding="utf-8")
    st = lic.load_state()
    assert st["_state_reason"] == "unsealed"
    assert st["activated"] is False


def test_offline_branch_requires_valid_code(tmp_license, monkeypatch):
    """封印有效但 code 无效（伪造）→ 绝不按「已激活」放行。"""
    monkeypatch.setattr(lic, "_safe_heartbeat", lambda code, machine: None)
    monkeypatch.setattr(lic, "_safe_trial_sync", lambda machine, claim=False: None)

    # ① 试用已用 → 直接拒（不能白嫖无限修正）
    _sealed_state(activated=True, offline=True, code="1234|1700000000|deadbeef",
                  kind="lifetime", trial_fix_used=True,
                  machine=lic._machine_fingerprint())
    g = lic.require_fix_entitlement()
    assert g["allowed"] is False
    assert g["reason"] != "activated"
    assert lic.load_state()["activated"] is False   # 状态被纠正为未激活

    # ② 试用未用 → 降级为一次免费试用（对「指纹已变的合法用户」保持宽容），
    #    但绝不能拿到 activated 级别的无限授权。
    _sealed_state(activated=True, offline=True, code="1234|1700000000|deadbeef",
                  kind="lifetime", trial_fix_used=False,
                  machine=lic._machine_fingerprint())
    g2 = lic.require_fix_entitlement()
    assert g2["reason"] == "trial"
    assert g2["remaining_trial"] == 1



def test_verify_offline_code_rejects_non_ascii_without_raising(offline_keys):
    """真 bug 回归：各种畸形 / 非 ASCII 输入不得让 verify 抛异常。

    在 GUI 里回调异常无控制台时完全静默 —— 用户看到「点了没反应」。
    """
    mc = lic._machine_fingerprint()
    for bad in ["", ".", mc, mc + "|", mc + ".", "机器码.签名",
                "机器码.中文签名", None, 12345]:
        assert lic.verify_offline_code(bad, mc) is False   # 不得抛异常
    assert lic.verify_offline_code(_sign(offline_keys, mc), mc) is True


def test_machine_fingerprint_is_stable_16_hex():
    """指纹：16 位小写 hex、同机稳定（多次调用一致）、不含敏感明文。"""
    fp1 = lic._machine_fingerprint()
    fp2 = lic._machine_fingerprint()
    assert fp1 == fp2
    assert len(fp1) == 16
    assert all(c in "0123456789abcdef" for c in fp1)
    import platform as _pf
    assert _pf.node() not in fp1     # 主机名不得以明文出现在指纹里


# ---------------------------------------------------------------------------
# 小阶/单位元点拒绝（纵深加固回归）
# ---------------------------------------------------------------------------

def test_rejects_small_order_points(tmp_license):
    """签名里的 R 或公钥 A 若为小阶点（含单位元）必须被拒。

    小阶点会让 [S]B == R + [k]A 出现可延展性；ed25519_verify._decompress 显式拒绝
    （[8]P == O 判据）。此用例把这道加固钉住，防日后被"优化"掉。
    """
    identity = b"\x01" + b"\x00" * 31          # 单位元（阶 1）

    # ① R 为单位元：构造 64 字节签名（R=单位元, S=1）
    mc = lic._machine_fingerprint()
    payload = "%s|%d" % (mc, 1700000000)
    code = payload + "." + base64.urlsafe_b64encode(
        identity + (1).to_bytes(32, "little")).decode("ascii")
    assert lic.verify_offline_code(code, mc) is False

    # ② 公钥 A 为单位元：即便签名结构正常也必须拒
    crypto_mod.set_verify_public_key(identity)
    try:
        assert crypto_mod.verify_offline_code(code, mc) is False
    finally:
        crypto_mod.set_verify_public_key(None)


# ---------------------------------------------------------------------------
# 封印密钥升级兼容（Codex 2026-09-12 P0-3）
# ---------------------------------------------------------------------------
def test_legacy_rsa_seal_is_still_accepted(tmp_path, monkeypatch):
    """v2.1.0 及更早用「RSA 公钥 PEM 的 utf-8 字节」当 HMAC 封印密钥。

    换成 Ed25519 公钥字节后，老用户本地那份**带 _sig 的** license.json 验签必然失败。
    若不兼容，load_state 会判 "tampered" —— 那条分支既丢掉 code 与试用计数、又不走
    「可验证迁移」，升级后直接被提示「免费试用已用完，请输入激活码」= 明确的用户事故。
    """
    import hashlib
    import hmac

    monkeypatch.setenv("TFD_LICENSE_FILE", str(tmp_path / "license.json"))
    payload = {
        "activated": True,
        "code": "PAID-CODE-123",
        "kind": "perpetual",
        "expires_at": None,
        "trial_fix_used": True,
        "revoked": False,
        "offline": True,
        "machine": lic._machine_fingerprint(),
        "last_heartbeat": 0,
    }
    blob = json.dumps(payload, ensure_ascii=False, sort_keys=True,
                      separators=(",", ":")).encode("utf-8")
    legacy = lic._LEGACY_SEAL_KEYS[0]
    legacy_sig = hmac.new(legacy, blob, hashlib.sha256).hexdigest()[:24]

    (tmp_path / "license.json").write_text(
        json.dumps(dict(payload, _sig=legacy_sig)), encoding="utf-8")
    st = lic.load_state()
    assert st["_state_reason"] == "ok", st["_state_reason"]
    assert st["activated"] is True
    assert st["code"] == "PAID-CODE-123"
    assert st["trial_fix_used"] is True

    # 兼容 ≠ 放松：改一个字段后必须仍然判篡改（老密钥也救不了伪造）
    (tmp_path / "license.json").write_text(
        json.dumps(dict(payload, code="FORGED", _sig=legacy_sig)),
        encoding="utf-8")
    assert lic.load_state()["_state_reason"] == "tampered"

    # 新密钥封印照常生效
    good = lic._state_sign(payload)
    (tmp_path / "license.json").write_text(
        json.dumps(dict(payload, _sig=good)), encoding="utf-8")
    assert lic.load_state()["_state_reason"] == "ok"

# ---------------------------------------------------------------------------
# 中台业务错误信封（非 2xx 也是"有效响应"）— 2026-09-12
# ---------------------------------------------------------------------------
def test_post_reads_body_of_non_2xx(monkeypatch):
    """中台用 HTTP 404 承载业务错误（``{"ok":false,"error":"invalid_code"}``）。

    ``urlopen`` 对非 2xx 抛 ``HTTPError``；若不当成响应读回来，用户输错一个码
    看到的是「网络不可用，可稍后重试」——把客户端错误说成客户网络问题。
    这里锁住：非 2xx 的 JSON 响应体必须被解析并返回。
    """
    import io
    import urllib.error

    body = json.dumps({"ok": False, "error": "invalid_code"}).encode("utf-8")

    def fake_urlopen(req, timeout=None):
        raise urllib.error.HTTPError(req.full_url, 404, "Not Found", {}, io.BytesIO(body))

    monkeypatch.setattr(lic.urllib.request, "urlopen", fake_urlopen)
    monkeypatch.setenv("TFD_FORCE_OFFLINE", "0")
    res = lic._post("/api/activate", {"product": "global", "code": "X", "machine_code": "M"})
    assert res == {"ok": False, "error": "invalid_code"}

    # 进而 activate() 必须给出"激活码无效"，而不是网络错误
    out = lic.activate("X")
    assert out["ok"] is False
    assert "无效" in out["message"], out["message"]


def test_post_reraises_when_error_body_is_not_json(monkeypatch):
    """真正的网络故障（网关返回 HTML/空体）仍须抛异常，由调用方走宽限。"""
    import io
    import urllib.error

    def fake_urlopen(req, timeout=None):
        raise urllib.error.HTTPError(req.full_url, 502, "Bad Gateway", {},
                                     io.BytesIO(b"<html>bad gateway</html>"))

    monkeypatch.setattr(lic.urllib.request, "urlopen", fake_urlopen)
    monkeypatch.setenv("TFD_FORCE_OFFLINE", "0")
    with pytest.raises(Exception):
        lic._post("/api/activate", {"product": "global", "code": "X", "machine_code": "M"})

# ---------------------------------------------------------------------------
# 中台业务错误信封（非 2xx 也是"有效响应"）— 2026-09-12
# ---------------------------------------------------------------------------
def test_post_reads_body_of_non_2xx(monkeypatch):
    """中台用 HTTP 404 承载业务错误（``{"ok":false,"error":"invalid_code"}``）。

    ``urlopen`` 对非 2xx 抛 ``HTTPError``；若不当成响应读回来，用户输错一个码
    看到的是「网络不可用，可稍后重试」——把客户端错误说成客户网络问题。
    这里锁住：非 2xx 的 JSON 响应体必须被解析并返回。
    """
    import io
    import urllib.error

    body = json.dumps({"ok": False, "error": "invalid_code"}).encode("utf-8")

    def fake_urlopen(req, timeout=None):
        raise urllib.error.HTTPError(req.full_url, 404, "Not Found", {}, io.BytesIO(body))

    monkeypatch.setattr(lic.urllib.request, "urlopen", fake_urlopen)
    monkeypatch.setenv("TFD_FORCE_OFFLINE", "0")
    res = lic._post("/api/activate", {"product": "global", "code": "X", "machine_code": "M"})
    assert res == {"ok": False, "error": "invalid_code"}

    # 进而 activate() 必须给出"激活码无效"，而不是网络错误
    out = lic.activate("X")
    assert out["ok"] is False
    assert "无效" in out["message"], out["message"]


def test_post_reraises_when_error_body_is_not_json(monkeypatch):
    """真正的网络故障（网关返回 HTML/空体）仍须抛异常，由调用方走宽限。"""
    import io
    import urllib.error

    def fake_urlopen(req, timeout=None):
        raise urllib.error.HTTPError(req.full_url, 502, "Bad Gateway", {},
                                     io.BytesIO(b"<html>bad gateway</html>"))

    monkeypatch.setattr(lic.urllib.request, "urlopen", fake_urlopen)
    monkeypatch.setenv("TFD_FORCE_OFFLINE", "0")
    with pytest.raises(Exception):
        lic._post("/api/activate", {"product": "global", "code": "X", "machine_code": "M"})
