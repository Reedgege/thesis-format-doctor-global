"""授权门禁（海外版 Thesis Format Doctor Global，隔离模块，不抄国内版源码）。

复用现有 tfd-auth 中台（Cloudflare Workers + D1 `tfd_auth`），海外版用独立
`product = "global"`（中台 `ensureProduct` 会自动建产品行），后台生成码时渠道选 `global`。
- 任何密钥 / Admin Key / Cloudflare Token 不进代码、不入仓、不入 memory。

模型 B（产品铁律）：
- 查（check）始终免费，无授权门禁。
- 修正（fix）：首次免费试用（TRIAL_FIX_LIMIT = 1），第 2 次起需激活码（lifetime 终身无限）。
- 全程零水印。

中台真实接口契约（与国内版共用同一套 Worker，已核对源码）：
- ``POST /api/activate``   {product, code, machine_code}
      → {ok:true, type, expires_at, status, uses_total, uses_used}
      | {ok:false, error: "invalid_code"|"revoked"|"expired"|"already_used"|"missing_fields"}
- ``POST /api/heartbeat``  {product, code, machine_code}
      → {ok:true, status: "active"|"revoked"|"expired", expires_at}
- ``POST /api/trial``      {product, machine_code, claim}
      → {ok:true, trial_used: bool, first: bool}
      **试用额度端点：需在共享中台部署后才生效**，见 `server_patch/README.md`。
      未部署时该端点返回 404，本模块按「服务器不可达」处理（离线宽限，本地状态兜底）。

防重置设计（FR6）：
- 本地状态机器绑定：状态文件里的 machine 与当前机器指纹不一致时，不认激活、可在线重绑；
- 状态文件损坏 / 非法 → **fail-closed**（按「试用已用」处理）；
- 状态文件原子写（tmp + os.replace），写一半崩溃不会退化成「未用过试用」；
- 在线时以中台为准核对试用/激活状态（服务器明确 exhausted/revoked/expired 才锁）；
- 网络不可达时按离线宽限，本地状态兜底（残余风险：离线删状态文件可重置，已在文档标注）。
- ``record_fix_used`` 绝不抛异常，避免「修正稿已生成但未计次」。

设计：
- 中台调用带浏览器 UA，避免 Cloudflare 1010 拦截。
- 心跳 3 天（HEARTBEAT_DAYS），仅当服务器返回 revoked/expired 才锁。
- 单元测试 monkeypatch 掉三个中台调用与状态路径，离线可跑。

不变量：
- 本模块不引入任何 docx / 格式引擎依赖；仅与中台和本地状态交互。
"""

from __future__ import annotations

import hashlib
import hmac
import json
import os
import platform
import sys
import time
import urllib.request
from pathlib import Path

# ---- 中台端点（生产配置；测试通过 monkeypatch 三个 _* 调用函数避开）----
API_BASE = "https://api.reedskill.com"
PRODUCT = "global"        # 海外版产品 id（中台 ensureProduct 自动建行）
TRIAL_FIX_LIMIT = 1       # 修正：1 次免费试用（产品模型 B：查免费、修正 1 次免费、之后付费）
HEARTBEAT_DAYS = 3        # 心跳宽限；超过仅当服务器返回 revoked/expired 才锁

_P_ACTIVATE = "/api/activate"
_P_HEARTBEAT = "/api/heartbeat"
_P_TRIAL = "/api/trial"

# 本地状态文件路径。
# 优先级：环境变量 TFD_LICENSE_FILE > 模块常量 LICENSE_FILE > 用户家目录下默认。
# Windows 下 Path.home() 不识别 HOME 环境变量，所以必须显式支持 TFD_LICENSE_FILE
# 才能在测试 / CI 中隔离。
LICENSE_FILE = Path.home() / ".thesis-format-doctor-global" / "license.json"

_USER_AGENT = "Mozilla/5.0 ThesisFormatDoctorGlobal/2.0"


def _license_file() -> Path:
    """实际状态文件路径（每次读 env，便于运行时切换）。"""
    env = os.environ.get("TFD_LICENSE_FILE")
    if env:
        return Path(env)
    return LICENSE_FILE


# ---------------------------------------------------------------------------
# 机器指纹 / 本地状态
# ---------------------------------------------------------------------------

def _raw_machine_signals() -> list:
    """采集「机器标识」信号（尽力而为：任一信号取不到都不影响其它信号）。

    为什么不再只用主机名：``platform.node()`` **完全由用户控制**——
    改个主机名就换指纹，既可能误伤合法用户（自己的码突然失效），
    也可能被用来蹭别人流传出去的码。所以优先使用「装系统时生成、改主机名不变」
    的硬件/系统标识，主机名只作为最后的兜底信号。

    三平台各自的稳定标识：
    - Windows: 注册表 ``HKLM\\SOFTWARE\\Microsoft\\Cryptography\\MachineGuid``（重装系统才变）
    - macOS:   ``IOPlatformUUID``（ioreg 读取）
    - Linux:   ``/etc/machine-id``（systemd 首次启动生成）

    全部失败也不抛异常：诚实降级到主机名，行为与旧版一致。
    """
    system = platform.system()
    sig = [system, platform.machine() or ""]

    if system == "Windows":
        try:
            import winreg  # noqa: PLC0415  Windows 专有标准库

            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,
                                r"SOFTWARE\Microsoft\Cryptography") as key:
                sig.append(str(winreg.QueryValueEx(key, "MachineGuid")[0]))
        except Exception:
            pass
    elif system == "Darwin":
        try:
            import re
            import subprocess

            out = subprocess.run(
                ["ioreg", "-rd1", "-c", "IOPlatformExpertDevice"],
                capture_output=True, text=True, timeout=3,
            ).stdout
            m = re.search(r'"IOPlatformUUID"\s*=\s*"([^"]+)"', out or "")
            if m:
                sig.append(m.group(1))
        except Exception:
            pass
    else:
        for p in ("/etc/machine-id", "/var/lib/dbus/machine-id"):
            try:
                value = Path(p).read_text(encoding="utf-8").strip()
            except Exception:
                continue
            if value:
                sig.append(value)
                break

    # 兜底信号：主机名（旧版唯一信号，保留以保证「取不到系统标识」时仍可用）
    sig.append(platform.node() or "")
    return sig


def _machine_fingerprint() -> str:
    """机器指纹 = SHA256(系统|架构|系统标识|主机名)[:16]。

    绑定粒度取舍：本产品是**离线**激活（不联网核对），指纹必须能由客户口头/截图
    报给卖家，所以取 16 位 hex（64 bit，碰撞概率可忽略），牺牲码长换易用性。

    注意：指纹变化（换机 / 重装系统）后离线码失效，需卖家重新发码——
    这是离线绑定的固有代价，已在文档与客服话术中说明。
    """
    raw = "|".join(_raw_machine_signals()).encode("utf-8", "replace")
    return hashlib.sha256(raw).hexdigest()[:16]


# 状态文件完整性封印（HMAC-SHA256，24 位 hex）。
# 封印密钥不再是什么「隐藏的对称秘密」——它直接取**内嵌公钥**的字节。公钥本就该
# 公开，所以这个封印只是「防小白手改 JSON 白嫖」的低成本护栏，不提供密码学强度
# （决心逆向者照样能重算）。真正能伪造激活码的**私钥**只在卖家本机（private_key.pem），
# 不在仓库、不在安装包，因此公开仓库无任何泄露风险。详见 src/license/crypto.py。
from .crypto import PUBLIC_KEY_PEM as _SEAL_KEY_PEM


def _state_sign(payload: dict) -> str:
    """状态文件完整性封印（HMAC-SHA256，24 位 hex）。

    目的**不是**「防死逆向」——公钥在客户端里，能逆向就能重算封印；目的是挡住
    「打开 license.json 把 activated 改成 true」这种**零成本白嫖**：
    手改后封印必然对不上，状态按被篡改处理（fail-closed）。
    真正的激活码伪造被 RSA 私钥挡住（私钥不在客户端），与此封印无关。
    """
    blob = json.dumps(payload, ensure_ascii=False, sort_keys=True,
                      separators=(",", ":")).encode("utf-8")
    return hmac.new(_SEAL_KEY_PEM.encode("utf-8"), blob, hashlib.sha256).hexdigest()[:24]


def _default_state() -> dict:
    return {
        "activated": False,
        "code": None,
        "kind": None,
        "expires_at": None,
        "trial_fix_used": False,
        "revoked": False,
        "offline": False,
        "machine": _machine_fingerprint(),
        "last_heartbeat": 0,
    }


def load_state() -> dict:
    """读本地状态。

    返回 dict 额外带两个**内部标记**（不会写回文件）：
    - ``_state_ok``: 状态文件是否可用（存在、JSON 合法、封印有效）；
    - ``_state_reason``: "ok" / "missing" / "corrupt" / "tampered" / "unsealed"。

    失败方向一律 **fail-closed**：宁可让用户重新激活，也不给"改文件白嫖"留窗口。
    """
    f = _license_file()
    base = _default_state()
    if not f.exists():
        base["_state_ok"] = False
        base["_state_reason"] = "missing"
        return base
    try:
        raw = json.loads(f.read_text(encoding="utf-8"))
        if not isinstance(raw, dict):
            raise ValueError("license state is not a JSON object")
    except Exception:
        base["_state_ok"] = False
        base["_state_reason"] = "corrupt"
        return base

    keys = set(base.keys())
    known = {k: v for k, v in raw.items() if k in keys}
    sig = raw.get("_sig")

    if not isinstance(sig, str) or not sig:
        # 没有封印 —— 只有两种可能：
        #   ① 旧版（v2.0.1 及更早）存的状态文件本来就不写 _sig；
        #   ② 有人把 _sig 那一行删了，想让试用/激活「重置」。
        # 本地无法区分，所以一律按 unsealed 处理（fail-closed，见 ③）。
        # 同时把线索留给 require_fix_entitlement() 做**可验证迁移**：只有
        # 「离线码能本地验签」或「在线码被中台确认为 active」才认 —— 伪造者拿不到
        # 有效码，迁移必然失败。于是：已付费用户升级无感，删 _sig 依旧白嫖不到。
        base.update(known)
        base["_state_ok"] = False
        base["_state_reason"] = "unsealed"
        if base.get("activated"):
            # 自称已激活但无法验证：先不认（否则删 _sig 就是万能后门），
            # 原始线索留给迁移分支去自证。
            base["activated"] = False
            base["_legacy_activated"] = True
            base["_legacy_code"] = known.get("code")
            base["_legacy_offline"] = bool(known.get("offline"))
        return base

    if not hmac.compare_digest(_state_sign(known), sig):
        base["activated"] = False
        base["_state_ok"] = False
        base["_state_reason"] = "tampered"
        return base

    base.update(known)
    base["_state_ok"] = True
    base["_state_reason"] = "ok"
    return base


def save_state(state: dict):
    """原子写状态文件（只落已知字段 + 完整性封印；内部 _state_* 标记不写盘）。

    原子性：先写 ``*.tmp`` 再 ``os.replace``，写一半崩溃不会退化成「未用过试用」。
    """
    f = _license_file()
    f.parent.mkdir(parents=True, exist_ok=True)
    keys = set(_default_state().keys())
    # 内部字段（_state_ok/_state_reason/_sig/_legacy_*）本就不写盘，也不算「未登记字段」。
    unknown = {k for k in state.keys() if k not in keys and not k.startswith("_")}
    payload = {k: v for k, v in state.items() if k in keys}
    payload["_sig"] = _state_sign(payload)
    tmp = f.with_name(f.name + ".tmp")
    tmp.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
    os.replace(str(tmp), str(f))
    # 未知字段会被静默丢弃——这曾导致 offline 标记丢失的真 bug（v2.0.1），
    # 所以这里不再沉默：调用方传了未登记字段时至少留下可诊断的痕迹。
    if unknown:
        try:
            print("[license] 警告：状态里存在未登记字段，已忽略：%s" % sorted(unknown),
                  file=sys.stderr)
        except Exception:
            pass


# ---------------------------------------------------------------------------
# 中台交互（带浏览器 UA，避免 1010；测试 monkeypatch 下面三个函数即可完全离线）
# ---------------------------------------------------------------------------

def _forced_offline() -> bool:
    """显式离线开关（环境变量 ``TFD_FORCE_OFFLINE=1``）。

    用途：
      ① 无网/内网环境下，用户与自动化不必每次操作都干等超时（默认 3–5 秒）；
      ② 自动化测试隔离中台，让回归不依赖服务器状态。

    **不绕过门禁**：关掉网络路径后仍走本地状态的既有裁决（未激活 = 1 次试用、
    离线授权直接放行），效果等同于拔网线，因此不引入新的白嫖口子。
    """
    return os.environ.get("TFD_FORCE_OFFLINE", "").strip().lower() in ("1", "true", "yes", "y")


def _post(path: str, payload: dict, timeout: int = 5) -> dict:
    """POST JSON 到中台。失败（网络/超时/非 JSON）抛异常，由调用方决定宽限策略。"""
    if _forced_offline():
        raise RuntimeError("offline mode (TFD_FORCE_OFFLINE)")
    url = API_BASE.rstrip("/") + path
    body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    req = urllib.request.Request(
        url, data=body, method="POST",
        headers={"Content-Type": "application/json", "User-Agent": _USER_AGENT},
    )
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        return json.loads(resp.read().decode("utf-8"))


def _check_mid_platform(code: str) -> dict:
    """向中台验证激活码（首次激活绑机）。"""
    return _post(_P_ACTIVATE, {
        "product": PRODUCT, "code": code, "machine_code": _machine_fingerprint(),
    })


def _heartbeat_check(code: str, machine: str) -> dict:
    """心跳 / 状态查询（只读，不扣次数、不绑机）。"""
    return _post(_P_HEARTBEAT, {
        "product": PRODUCT, "code": code, "machine_code": machine,
    }, timeout=4)


def _trial_sync(machine: str, claim: bool = False) -> dict:
    """试用额度登记 / 查询（中台侧一次性、按机器绑定）。

    claim=False → 只查该机器是否已用过试用；
    claim=True  → 首次领取（幂等：已领过返回 first=False）。

    端点未部署（404）会抛异常 → 由 _safe_* 包成「不可达」，走本地兜底。
    """
    return _post(_P_TRIAL, {
        "product": PRODUCT, "machine_code": machine, "claim": bool(claim),
    }, timeout=3)


def _safe_heartbeat(code, machine) -> dict | None:
    if not code:
        return None
    try:
        return _heartbeat_check(code, machine)
    except Exception:
        return None


def _safe_trial_sync(machine: str, claim: bool = False) -> dict | None:
    try:
        return _trial_sync(machine, claim)
    except Exception:
        return None


# ---------------------------------------------------------------------------
# 公开 API
# ---------------------------------------------------------------------------

_ACTIVATE_ERR = {
    "invalid_code": "激活码无效或不属于本产品",
    "revoked": "该激活码已被吊销（退款锁死）",
    "expired": "该激活码已过期",
    "already_used": "该激活码已绑定其他机器（永久/周/月卡限一机）",
    "missing_fields": "请求字段缺失（产品或激活码为空）",
}


def activate(code: str) -> dict:
    """输入激活码，向中台验证并写入本地状态。返回 {ok, kind, message}。"""
    code = (code or "").strip()
    if not code:
        return {"ok": False, "message": "激活失败：请先输入激活码。"}
    try:
        res = _check_mid_platform(code)
    except Exception as e:
        return {"ok": False, "message": f"激活失败：{e}（网络不可用时可稍后重试）"}
    if not res.get("ok"):
        err = res.get("error", "")
        return {"ok": False,
                "message": f"激活失败：{_ACTIVATE_ERR.get(err, err or '中台未通过')}"}
    state = load_state()
    state["activated"] = True
    state["code"] = code
    state["kind"] = res.get("type") or res.get("kind") or "lifetime"
    state["expires_at"] = res.get("expires_at")
    state["revoked"] = False
    state["machine"] = _machine_fingerprint()
    state["last_heartbeat"] = int(time.time())
    try:
        save_state(state)
    except Exception as e:
        return {"ok": False, "message": f"激活失败：无法写入本地状态（{e}）"}
    _safe_trial_sync(_machine_fingerprint(), claim=True)
    return {"ok": True, "kind": state["kind"], "message": "激活成功"}


# ---------------------------------------------------------------------------
# 兜底方案：离线备用码（RSA 非对称签名，见 src/license/crypto.py）
# ---------------------------------------------------------------------------
# 算法与 tools/reedcode*.py 完全一致（machine_code|时间戳 → RSA 签名）。
# 与导师版密钥隔离（各自独立 RSA 密钥对），保证跨产品码不互通。客户把本机机器码
# 发给卖家，卖家用 tools/reedcode_gui.py（持本机私钥）生成离线码，客户在激活页选
# 「离线激活」粘贴即可。私钥只在卖家本机，公开仓库 / 反编译客户端都拿不到。
from .crypto import verify_offline_code


def activate_offline(code: str) -> dict:
    """离线激活：校验离线码（绑定本机机器码）→ 写入本地状态（终身、离线、不联网）。

    返回 {ok, kind, message}。离线激活后授权为终身且永久离线，不再触网（隐私：离线版不回传）。
    """
    code = (code or "").strip()
    if not code:
        return {"ok": False, "message": "离线激活码为空"}
    fp = _machine_fingerprint()
    if not verify_offline_code(code, fp):
        return {"ok": False, "message":
                "离线激活码无效，或不属于本机（请确认是用本机的机器码生成的码）。"}
    state = load_state()
    state["activated"] = True
    state["code"] = code
    state["kind"] = "lifetime"
    state["expires_at"] = None
    state["revoked"] = False
    state["offline"] = True
    state["machine"] = fp
    state["last_heartbeat"] = int(time.time())
    try:
        save_state(state)
    except Exception as e:
        return {"ok": False, "message": f"离线激活失败：无法写入本地状态（{e}）"}
    return {"ok": True, "kind": "lifetime",
            "message": "离线激活成功（终身版，无需联网）。"}


def require_fix_entitlement() -> dict:
    """修正前的授权裁决（CLI/GUI 在调用 fixer 前先调）。

    返回 {allowed, reason, message, remaining_trial}：
      - allowed True / reason "trial"           未激活且试用未用
      - allowed True / reason "activated"       已激活（含自动重绑机器）
      - allowed False / reason "trial_exhausted" 试用已用完
      - allowed False / reason "revoked"        服务器明确吊销/过期

    注意：本函数可能写入本地状态（把服务器判定的「试用已用」/「吊销」落到本地，
    属 fail-closed 方向），但绝不抛出异常。
    """
    try:
        return _require_fix_entitlement()
    except Exception as e:   # 授权模块自身异常不应阻断用户，按离线宽限放行试用
        return {"allowed": True, "reason": "trial",
                "message": f"授权状态读取异常（{e}），按首次免费试运行。",
                "remaining_trial": 1}


def _require_fix_entitlement() -> dict:
    state = load_state()
    fp = _machine_fingerprint()

    # ① 机器绑定：状态文件来自别的机器 → 不认激活（在线可凭激活码重新绑定）
    if state.get("_state_ok") and state.get("machine") and state["machine"] != fp:
        hb = _safe_heartbeat(state.get("code"), fp)
        if hb and hb.get("ok") and hb.get("status") == "active":
            state["machine"] = fp
            state["last_heartbeat"] = int(time.time())
            save_state(state)
        else:
            state["activated"] = False

    # ①-b 离线授权：已离线激活 → 直接放行，绝不联网（隐私：离线版不回传）。
    #      但**绝不盲信本地布尔字段**：手写 {"activated":true,"offline":true} 就能
    #      永久白嫖。所以这里重新验签——没有有效离线码就不放行（fail-closed）。
    if state.get("activated") and state.get("offline"):
        if verify_offline_code(state.get("code") or "", fp):
            return {"allowed": True, "reason": "activated",
                    "message": f"已离线激活（{state.get('kind') or 'lifetime'}）",
                    "remaining_trial": None}
        # 状态被改动 / 指纹已变（换机、重装）→ 降级，走下面的试用与激活逻辑
        state["activated"] = False
        state["offline"] = False
        try:
            save_state(state)
        except Exception:
            pass

    # ①-c 旧版状态文件（无封印）的**可验证迁移** —— 保护 v2.0.1 及更早版本的
    #      已付费用户：那批版本的 save_state 不写 _sig，升级到本版后会被判
    #      unsealed → 未激活，弹窗还会误导成「试用已用完」（明确的用户事故）。
    #      这里只认**能自证**的身份，两条路：
    #        ① 离线码：本地 HMAC 验签（不联网，密码学上等价于 ①-b）；
    #        ② 在线码：问一次中台，服务器确认本机 active。
    #      伪造者拿不到有效码 → 两条路都失败 → 依旧 fail-closed（不给"删 _sig 即
    #      白嫖"留口子），但给**专门文案**而不是"试用已用完"。
    if state.get("_state_reason") == "unsealed" and state.get("_legacy_activated"):
        legacy_code = state.get("_legacy_code") or ""
        legacy_offline = bool(state.get("_legacy_offline"))
        migrated = False
        if legacy_offline and verify_offline_code(legacy_code, fp):
            migrated = True
        elif legacy_code:
            hb = _safe_heartbeat(legacy_code, fp)
            migrated = bool(hb and hb.get("ok") and hb.get("status") == "active")
        if migrated:
            state["activated"] = True
            state["code"] = legacy_code
            state["offline"] = legacy_offline
            state["machine"] = fp
            state["last_heartbeat"] = int(time.time())
            try:
                save_state(state)          # 重封印 → 从此走上新格式
            except Exception:
                pass
            kind = state.get("kind") or "lifetime"
            return {"allowed": True, "reason": "activated",
                    "message": (f"已离线激活（{kind}）" if legacy_offline
                                else f"已激活（{kind}）"),
                    "remaining_trial": None}
        return {"allowed": False, "reason": "needs_reactivation",
                "message": "检测到旧版授权文件（本次升级前激活的）。"
                           "请重新输入一次激活码即可恢复：离线码直接粘贴，在线码需联网。",
                "remaining_trial": 0}

    # ② 已激活：心跳 + 吊销/过期检查（服务器明确 revoked/expired 才锁；离线宽限）
    if state.get("activated"):
        age = time.time() - float(state.get("last_heartbeat") or 0)
        if age > HEARTBEAT_DAYS * 86400:
            hb = _safe_heartbeat(state.get("code"), fp)
            if hb and hb.get("status") in ("revoked", "expired"):
                state["activated"] = False
                state["revoked"] = True
                save_state(state)
            elif hb:
                state["last_heartbeat"] = int(time.time())
                save_state(state)
        if state.get("activated"):
            return {"allowed": True, "reason": "activated",
                    "message": f"已激活（{state.get('kind') or 'lifetime'}）",
                    "remaining_trial": None}
        return {"allowed": False, "reason": "revoked",
                "message": "授权已被吊销或已过期，请重新输入激活码。",
                "remaining_trial": 0}

    # ③ 未激活：试用额度（中台优先、本地兜底、损坏/被篡改 fail-closed）
    trial_used = bool(state.get("trial_fix_used"))
    if state.get("_state_reason") in ("corrupt", "tampered", "unsealed"):
        # 损坏、被改动、或「有文件但没封印」→ 一律按已用处理。
        # 关键：unsealed 也必须 fail-closed，否则「删掉 _sig 那一行」就等于
        # 免费重置试用（删一次白嫖一次）。全新用户走的是 missing，不受影响。
        trial_used = True

    srv = _safe_trial_sync(fp, claim=False)
    if srv is not None and srv.get("trial_used"):
        trial_used = True
        if not state.get("trial_fix_used"):
            try:
                state["trial_fix_used"] = True
                save_state(state)
            except Exception:
                pass

    if trial_used:
        return {"allowed": False, "reason": "trial_exhausted",
                "message": "免费试用已用完，请输入激活码解锁无限修正。",
                "remaining_trial": 0}
    return {"allowed": True, "reason": "trial",
            "message": f"首次免费试用（共 {TRIAL_FIX_LIMIT} 次）",
            "remaining_trial": 1}


def record_fix_used() -> bool:
    """修正成功后调用：标记试用已用 + 心跳 + 中台登记。

    **绝不抛异常**（否则会出现「修正稿已生成但未计次」，用户可反复白嫖）。
    返回是否成功落盘。
    """
    ok = True
    try:
        state = load_state()
        if not state.get("activated"):
            state["trial_fix_used"] = True
        state["last_heartbeat"] = int(time.time())
        save_state(state)
    except Exception:
        ok = False
    _safe_trial_sync(_machine_fingerprint(), claim=True)
    return ok


def post_fix_message(gate: dict, counted: bool = True) -> str:
    """修正**成功之后**该展示的状态文案。

    与 ``gate["message"]`` 的区别：gate 是修正**前**的裁决，此时试用额度尚未扣减，
    直接展示会出现「刚用完免费试用，却还显示首次免费试用」的错位。
    """
    if not counted:
        return "试用计数写入失败（不影响本次修正结果，但请留意）。"
    if gate.get("reason") == "activated":
        return f"已激活（{gate.get('kind') or 'lifetime'}），可继续无限修正。"
    return (f"免费试用额度已用尽（{TRIAL_FIX_LIMIT}/{TRIAL_FIX_LIMIT}），"
            f"下次修正需输入激活码。")


def status() -> dict:
    """仅返回当前授权状态（供 GUI 展示）。"""
    s = load_state()
    return {
        "activated": bool(s.get("activated")),
        "kind": s.get("kind"),
        "expires_at": s.get("expires_at"),
        "trial_fix_used": bool(s.get("trial_fix_used")),
        "revoked": bool(s.get("revoked")),
        "offline": bool(s.get("offline")),
        "state_ok": bool(s.get("_state_ok")),
        "machine": s.get("machine") or _machine_fingerprint(),
        "trial_limit": TRIAL_FIX_LIMIT,
    }
