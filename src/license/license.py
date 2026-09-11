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
import json
import os
import platform
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

def _machine_fingerprint() -> str:
    """机器指纹（不依赖 MAC/序列号，避免权限与跨平台问题）。

    以 hostname（platform.node）为主：容器 / 改机名会变，属已知限制；
    变化后在线可凭激活码自动重新绑定（见 require_fix_entitlement）。
    """
    node = platform.node() or ""
    raw = f"{platform.system()}|{node}|{platform.machine()}".encode()
    return hashlib.sha256(raw).hexdigest()[:16]


def _default_state() -> dict:
    return {
        "activated": False,
        "code": None,
        "kind": None,
        "expires_at": None,
        "trial_fix_used": False,
        "revoked": False,
        "machine": _machine_fingerprint(),
        "last_heartbeat": 0,
    }


def load_state() -> dict:
    """读本地状态。

    返回 dict 额外带两个**内部标记**（不会写回文件）：
    - ``_state_ok``: 状态文件是否可用（存在且 JSON 合法）；
    - ``_state_reason``: "ok" / "missing" / "corrupt"。
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
    base.update({k: v for k, v in raw.items() if k in base})
    base["_state_ok"] = True
    base["_state_reason"] = "ok"
    return base


def save_state(state: dict):
    """原子写状态文件（只落已知字段；内部 _state_* 标记不写盘）。"""
    f = _license_file()
    f.parent.mkdir(parents=True, exist_ok=True)
    keys = set(_default_state().keys())
    payload = {k: v for k, v in state.items() if k in keys}
    tmp = f.with_name(f.name + ".tmp")
    tmp.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
    os.replace(str(tmp), str(f))


# ---------------------------------------------------------------------------
# 中台交互（带浏览器 UA，避免 1010；测试 monkeypatch 下面三个函数即可完全离线）
# ---------------------------------------------------------------------------

def _post(path: str, payload: dict, timeout: int = 5) -> dict:
    """POST JSON 到中台。失败（网络/超时/非 JSON）抛异常，由调用方决定宽限策略。"""
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

    # ③ 未激活：试用额度（中台优先、本地兜底、损坏 fail-closed）
    trial_used = bool(state.get("trial_fix_used"))
    if state.get("_state_reason") == "corrupt":
        trial_used = True          # 状态损坏 → 按已用处理（不给白嫖窗口）

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
        "state_ok": bool(s.get("_state_ok")),
        "machine": s.get("machine") or _machine_fingerprint(),
        "trial_limit": TRIAL_FIX_LIMIT,
    }
