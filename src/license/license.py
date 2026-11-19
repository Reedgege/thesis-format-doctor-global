"""授权门禁（海外版 PaperFormat Pro，隔离模块，不抄国内版源码）。

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
      （type ∈ times/quarterly/annual/lifetime；expires_at 为毫秒；times 卡回 uses_total/uses_used）
- ``POST /api/consume``    {product, code, machine_code, dry_run, uses_used?}
      → 次卡计次（中台为准）：dry_run=true 仅预检；false 真扣，uses_used 为绝对已用次数（服务端取 max，幂等）
      → {ok:true, uses_total, uses_used} | {ok:false, error:"exhausted"|...}
- ``POST /api/heartbeat``  {product, code, machine_code}
      → {ok:true, status: "active"|"revoked"|"expired", expires_at,
         latest_version?, download_url?, min_version?}
      （latest_version 等版本字段为可选；缺失时客户端按「无更新」降级处理）

中台真实接口契约（与国内版共用同一套 Worker，已核对源码）：
- ``GET /api/version``      （无需授权，公共端点，供试用 / 离线用户取更新提示）
      → {latest_version, download_url, min_version?}
      **该端点需在中台部署后生效**；未部署时返回非 JSON（网关 HTML）或 404，
      本模块按「服务器不可达」处理（降级为无更新，不阻塞启动）。
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
- 心跳：启动 + 每次修正都向中台核对吊销/过期（方向 A，撤销即时生效）；仅服务器明确 revoked/expired 才锁，离线宽限保留。
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
import shutil
import socket
import sys
import time
import urllib.error
import urllib.request

# 版本号单一来源：User-Agent 里的版本号随 VERSION 文件自动同步，不写死（F4）。
from ..versioninfo import app_version
from pathlib import Path

# ---- 中台端点（生产配置；测试通过 monkeypatch 三个 _* 调用函数避开）----
API_BASE = "https://api.reedskill.com"
PRODUCT = "global"        # 海外版产品 id（中台 ensureProduct 自动建行）
TRIAL_FIX_LIMIT = 1       # 修正：1 次免费试用（产品模型 B：查免费、修正 1 次免费、之后付费）

# 授权类型（kind）取值：与中台 /api/activate 回的 type 字段对齐。
# 单位约定：expires_at 一律**毫秒**（Unix epoch ms），与国内版 + 共用 Worker 一致；
# 切勿直接转发 Paddle 的秒级时间戳，否则期限差 1000 倍。
KIND_TIMES = "times"          # 次卡：按次计费，uses_total/uses_used 有效
KIND_QUARTERLY = "quarterly"  # 季卡：一次付费 90 天，expires_at = 激活时刻 + 90d(ms)
KIND_ANNUAL = "annual"        # 年卡：订阅循环，expires_at 由心跳写回（续费自动延长）
KIND_LIFETIME = "lifetime"    # 终身：无期限

_P_ACTIVATE = "/api/activate"
_P_HEARTBEAT = "/api/heartbeat"
_P_TRIAL = "/api/trial"
_P_CONSUME = "/api/consume"   # 次卡计次（中台为准，dry_run 预检 / 真扣）

# 本地状态文件路径。
# 优先级：环境变量 TFD_LICENSE_FILE > 模块常量 LICENSE_FILE > 用户家目录下默认。
# Windows 下 Path.home() 不识别 HOME 环境变量，所以必须显式支持 TFD_LICENSE_FILE
# 才能在测试 / CI 中隔离。
LICENSE_FILE = Path.home() / ".paperformatpro" / "license.json"

# v2.3.10 重命名：旧隐藏目录 .thesis-format-doctor-global 仍在已装用户的机器上，
# 一次性迁移回退——读不到新路径时尝试旧路径，避免老用户升级后激活/试用状态丢失。
LICENSE_FILE_LEGACY = Path.home() / ".thesis-format-doctor-global" / "license.json"

_USER_AGENT = f"Mozilla/5.0 PaperFormatPro/{app_version()}"

# 启动一次性核验通过后的「内存可信标志」：置位后，本次会话内修正不再发心跳。
# 用户决策：启动核对一次即可，有效状态就不再打扰网络（见 _require_fix_entitlement ②）。
_launch_verified_active = False


def _license_file() -> Path:
    """实际状态文件路径（每次读 env，便于运行时切换）。

    v2.3.10 重命名：新目录 ``.paperformatpro`` 不存在但旧目录
    ``.thesis-format-doctor-global`` 的 license.json 还在时，一次性复制到新路径，
    后续读写都走新路径 —— 老用户升级后激活 / 试用状态不丢。
    """
    env = os.environ.get("TFD_LICENSE_FILE")
    if env:
        return Path(env)
    if not LICENSE_FILE.exists() and LICENSE_FILE_LEGACY.exists():
        try:
            LICENSE_FILE.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(LICENSE_FILE_LEGACY, LICENSE_FILE)
        except Exception:
            return LICENSE_FILE_LEGACY
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
# 封印密钥直接取**内嵌公钥**的字节（32 字节，base64 存于 crypto.PUBLIC_KEY）。
# 公钥本就该公开，所以这个封印只是「防小白手改 JSON 白嫖」的低成本护栏，不提供密码学强度
# （决心逆向者照样能重算）。真正能伪造激活码的**Ed25519 私钥**只在卖家本机
# （_signing_keys/overseas_ed25519_private.pem），不在仓库、不在安装包。详见 crypto.py。
from .crypto import PUBLIC_KEY as _SEAL_KEY

# 历史封印密钥：v2.1.0 及更早用「RSA 公钥 PEM 的 utf-8 字节」当 HMAC 密钥。
# **必须继续接受**，否则老用户本地那份**带 _sig 的** license.json 验签必然失败 →
# load_state 判 "tampered" → 那条分支既丢掉 code 与试用计数、又不走下面的
# "可验证迁移"（那条分支是专门为保护老付费用户写的）→ 升级后直接被提示
# 「免费试用已用完，请输入激活码」，这是明确的用户事故（Codex 2026-09-12 P0-3）。
#
# 安全性：公钥本就公开，多接受一把**公开**密钥不引入任何新的伪造面 ——
# 能读到源码的人本来就能用新密钥重算封印（封印只是防"手改 JSON 白嫖"的低成本护栏）。
# 值取自 `git show HEAD:src/license/crypto.py`（换算法前的最后一个提交）。
_LEGACY_SEAL_KEYS = (
    b"""-----BEGIN RSA PUBLIC KEY-----
MIIBCgKCAQEAhFJwSG+B2lh5+nf/aB9sJQmOxbUMlyxHj1nqlEkAzJfRXfNPi9mm
ABVzJSV0umzZ4d/goU3FejObhfsjmKEkOWFtA4iBhVuB9XrbDDF2vcLHsqDFmpX7
nxw6DzxCQzPGMMSlTqD/k2mUZtg8MLtnUoD/k09r3paC7ZstnEl6MQSwTCqq6+Oq
tGtoTZwVlkQiCo3LgFMIXUvQ5qw5UQbr6kZBHaI/Xl4U3gTWxKyG9o6C2SqrvH3g
15KSkIX8ZbCKJr9mhK5Su14xt3dtDji+HdMrE3aNUmbub6Ud5R+fgfz/l2bJ3QGE
QdwW1zXfHfE9BriXOsms5tCvEK1p3uQ3CwIDAQAB
-----END RSA PUBLIC KEY-----
""",
)


def _seal_blob(payload: dict) -> bytes:
    return json.dumps(payload, ensure_ascii=False, sort_keys=True,
                      separators=(",", ":")).encode("utf-8")


def _state_sign(payload: dict) -> str:
    """状态文件完整性封印（HMAC-SHA256，24 位 hex）。

    目的**不是**「防死逆向」——公钥在客户端里，能逆向就能重算封印；目的是挡住
    「打开 license.json 把 activated 改成 true」这种**零成本白嫖**：
    手改后封印必然对不上，状态按被篡改处理（fail-closed）。
    真正的激活码伪造被 Ed25519 私钥挡住（私钥不在客户端），与此封印无关。
    """
    return hmac.new(_SEAL_KEY, _seal_blob(payload), hashlib.sha256).hexdigest()[:24]


def _seal_matches(known: dict, sig: str) -> bool:
    """封印校验：先试当前密钥（Ed25519 公钥字节），再试历史密钥（RSA PEM）。

    兼容历史密钥是**升级兼容性**要求，不是安全放松：任何异常一律返回 False。
    """
    try:
        if hmac.compare_digest(_state_sign(known), sig):
            return True
    except Exception:
        return False
    blob = _seal_blob(known)
    for key in _LEGACY_SEAL_KEYS:
        try:
            if hmac.compare_digest(
                    hmac.new(key, blob, hashlib.sha256).hexdigest()[:24], sig):
                return True
        except Exception:
            continue
    return False


def _default_state() -> dict:
    return {
        "activated": False,
        "code": None,
        "kind": None,
        # expires_at：到期时间戳（毫秒，Unix epoch ms）。季卡/年卡有效；终身/试用为 None。
        "expires_at": None,
        "trial_fix_used": False,
        "revoked": False,
        "offline": False,
        "machine": _machine_fingerprint(),
        "last_heartbeat": 0,
        # —— 次卡（按次计费）专有字段 ——
        # uses_total: 总次数（次卡=2；非次卡为 None 表示「无限/不适用」）
        # uses_used: 已用次数（本地乐观推进；中台 /api/consume 为准）
        # owed: 本地欠账户数（扣次失败/离线时记一笔，下次联网 pay_owed 补扣）
        "uses_total": None,
        "uses_used": 0,
        "owed": 0,
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

    if not _seal_matches(known, sig):
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
            print("[license] Warning: unregistered fields found in state, ignored: %s" % sorted(unknown),
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


def _cacert_candidates() -> list:
    """自带 CA 证书包的候选路径（源码运行 + Nuitka 打包后的几种布局）。

    打包布局与 ``ui/iconpath.py`` 一致：整目录 assets/ 被拷到 exe 同级，
    所以产物里是 ``<exe 目录>/assets/cacert.pem``；源码运行时是 ``<repo>/assets``。
    """
    here = os.path.dirname(os.path.abspath(__file__))       # <repo>/src/license
    repo = os.path.normpath(os.path.join(here, "..", ".."))
    cands = [os.path.join(repo, "assets", "cacert.pem")]
    for raw in (getattr(sys, "argv", None) or [""])[:1] + [getattr(sys, "executable", "")]:
        try:
            if raw:
                base = os.path.dirname(os.path.abspath(raw))
                cands.append(os.path.join(base, "assets", "cacert.pem"))
        except Exception:
            continue
    cands.append(os.path.join(os.getcwd(), "assets", "cacert.pem"))
    out = []
    for c in cands:
        c = os.path.normpath(c)
        if c not in out:
            out.append(c)
    return out


def _ssl_context():
    """用随程序打包的 CA 证书建 SSL context。

    修复老 Linux 在线激活报 ``SSL: CERTIFICATE_VERIFY_FAILED —
    unable to get local issuer certificate``：中台 api.reedskill.com 用
    Google GTS Root R4（2021 新根），老发行版 ca-certificates 不含它；Windows
    自动更新根库、macOS 11+ 钥匙串已收录，只有 Linux 会中招。自带证书后
    三平台行为一致。找不到自带证书则回退系统默认，**绝不降低验证强度**
    （不用 _create_unverified_context）。"""
    try:
        import ssl
        for path in _cacert_candidates():
            if os.path.isfile(path):
                return ssl.create_default_context(cafile=path)
    except Exception:
        pass
    return None


def _open_url(req, timeout):
    """统一网络出口：能用自带 CA 就用，否则回退 urlopen 系统默认。

    抽成单一函数，激活 / 心跳 / 版本检查共用同一条证书路径；测试只需
    monkeypatch 本函数即可离线隔离，不必关心 opener 细节。"""
    ctx = _ssl_context()
    if ctx is None:
        return urllib.request.urlopen(req, timeout=timeout)
    opener = urllib.request.build_opener(urllib.request.HTTPSHandler(context=ctx))
    return opener.open(req, timeout=timeout)


def _post(path: str, payload: dict, timeout: int = 5) -> dict:
    """POST JSON 到中台，**业务错误也当响应体读回来**（非 2xx 不算网络故障）。

    中台用 HTTP 4xx/5xx 承载业务结果，例如输错码时返回
    ``{"ok":false,"error":"invalid_code"}`` 配 **HTTP 404**。``urlopen`` 对非 2xx 抛
    ``HTTPError``，若不把它当响应读取，客户端只能拿到 "HTTP Error 404: Not Found"，
    于是用户输错一个码会看到「网络不可用，可稍后重试」——既误导排查、又让客户以为
    是自家网络问题（海外版漏了这层，2026-09-12 补；导师版/学生版的 ``_http_json``
    一直是这么处理的）。真正的网络故障（连不上/超时/响应不是 JSON）仍照旧抛异常，
    由调用方决定宽限策略。
    """
    if _forced_offline():
        raise RuntimeError("offline mode (TFD_FORCE_OFFLINE)")
    url = API_BASE.rstrip("/") + path
    body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    req = urllib.request.Request(
        url, data=body, method="POST",
        headers={"Content-Type": "application/json", "User-Agent": _USER_AGENT},
    )
    # socket 默认超时兜底：urlopen(timeout=) 覆盖不到所有悬挂路径（DNS 解析阶段最典型）。
    prev = socket.getdefaulttimeout()
    socket.setdefaulttimeout(timeout)
    try:
        with _open_url(req, timeout) as resp:
            return json.loads(resp.read().decode("utf-8", "ignore"))
    except urllib.error.HTTPError as e:
        # 非 2xx：先按响应体解析（中台的业务错误信封都在这里）；解析不出来才算网络故障。
        try:
            return json.loads(e.read().decode("utf-8", "ignore"))
        except Exception:
            raise
    finally:
        socket.setdefaulttimeout(prev)


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

    端点未部署 → 响应不是 JSON（如网关 HTML 404）→ 抛异常 → 由 ``_safe_*`` 包成
    「不可达」，走本地兜底。端点存在但返回业务错误（JSON 信封）则会正常读回，
    调用方必须显式要求 ``ok=True`` 才采信（见 ``server_trial_used``）。
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


def _sync_from_heartbeat(state: dict, hb: dict) -> None:
    """心跳 active 时，把服务端下发的过期时间 / 计次写回本地（不落盘，调用方负责 save）。

    - expires_at（毫秒）：年卡续费后服务端下发新值 → 本地自动延长；
    - times 次卡：以服务端 uses_total/uses_used 为准，结清欠账（防多机不同步）。
    """
    if hb.get("expires_at"):
        state["expires_at"] = hb["expires_at"]
    if state.get("kind") == KIND_TIMES:
        if hb.get("uses_total") is not None:
            state["uses_total"] = hb["uses_total"]
        if hb.get("uses_used") is not None:
            state["uses_used"] = hb["uses_used"]
            if hb.get("uses_used", 0) >= (hb.get("uses_total") or 0):
                state["owed"] = 0


def _apply_heartbeat_revocation(state: dict, fp: str) -> bool:
    """对在线激活码做一次心跳吊销核对。返回 True 表示已被锁（revoked/expired）。

    方向 A：撤销即时生效——后台一停用，下次启动或下次点修正即锁。
    fail-soft：网络不可达（_safe_heartbeat 返回 None）不锁、不报错（离线宽限）。
    服务器可达且仍 active 时刷新本地 last_heartbeat，并写回服务端下发的
    expires_at / 计次（年卡续费自动延长、次卡多机同步）。
    """
    hb = _safe_heartbeat(state.get("code"), fp)
    if hb and hb.get("status") in ("revoked", "expired"):
        state["activated"] = False
        state["revoked"] = True
        try:
            save_state(state)
        except Exception:
            pass
        return True
    if hb:  # 可达且 active：刷新戳 + 写回服务端下发字段
        state["last_heartbeat"] = int(time.time())
        _sync_from_heartbeat(state, hb)
        try:
            save_state(state)
        except Exception:
            pass
    return False


def _safe_trial_sync(machine: str, claim: bool = False) -> dict | None:
    try:
        return _trial_sync(machine, claim)
    except Exception:
        return None


# ---------------------------------------------------------------------------
# 次卡（按次计费）计次：中台 /api/consume 为准，本地 owed 账本兜底防离线白嫖
# ---------------------------------------------------------------------------
MAX_OWED = 5  # 本地欠账上限：超过则视为异常（仍 fail-closed，待联网补扣失败才锁）


def _consume_use(dry_run: bool = False) -> dict | None:
    """向中台登记 / 预检一次次卡消耗。返回服务端响应或 None（网络不可达）。

    dry_run=True  → 只预检不真扣（修正前门禁用）；
    dry_run=False → 真扣：上报绝对已用次数（幂等，服务端取 max），
                   避免补扣 / 多机场景重复计数。
    """
    state = load_state()
    if not state.get("code"):
        return None
    payload = {
        "product": PRODUCT,
        "code": state.get("code"),
        "machine_code": _machine_fingerprint(),
        "dry_run": bool(dry_run),
    }
    if not dry_run:
        payload["uses_used"] = state.get("uses_used") or 0
    try:
        return _post(_P_CONSUME, payload, timeout=5)
    except Exception:
        return None


def pay_owed() -> None:
    """联网恢复后，把本地 owed 欠账户数一次性补扣到中台（次卡）。绝不抛异常。

    发送绝对 uses_used（含所有乐观推进 + 欠账），服务端取 max → 幂等；
    成功则本地 owed 清零、uses_used 以服务端为准。失败保留欠账，下次再补。
    """
    try:
        state = load_state()
        if state.get("kind") != KIND_TIMES or not state.get("code"):
            return
        if (state.get("owed") or 0) <= 0:
            return
        res = _consume_use(dry_run=False)
        if res and res.get("ok"):
            state = load_state()
            if res.get("uses_used") is not None:
                state["uses_used"] = res["uses_used"]
            state["owed"] = 0
            try:
                save_state(state)
            except Exception:
                pass
    except Exception:
        return


# ---------------------------------------------------------------------------
# 版本公告（公共端点，无需激活码；未部署则降级为「无更新」）
# ---------------------------------------------------------------------------
_P_VERSION = "/api/version"


def _parse_version_tuple(v: str) -> tuple:
    """把 ``"2.3.58"`` 解析成 ``(2, 3, 58)`` 用于比较；非法段按 0 处理。"""
    out = []
    for part in (v or "").split("."):
        try:
            out.append(int(part))
        except Exception:
            out.append(0)
    return tuple(out) if out else (0,)


def _current_version_tuple() -> tuple:
    try:
        from .. import versioninfo
        return _parse_version_tuple(versioninfo.app_version())
    except Exception:
        return (0,)


def _check_latest_version() -> dict | None:
    """公共版本端点（无需授权）；失败软降级为 None。

    与心跳不同，它不需要激活码，因此试用 / 离线用户也能拿到更新提示。
    端点未部署或网络不可达时返回 None（上层按「无更新」处理，绝不阻塞启动）。
    """
    if _forced_offline():
        return None
    url = API_BASE.rstrip("/") + _P_VERSION
    req = urllib.request.Request(
        url, method="GET", headers={"User-Agent": _USER_AGENT})
    prev = socket.getdefaulttimeout()
    socket.setdefaulttimeout(4)
    try:
        with _open_url(req, 4) as resp:
            return json.loads(resp.read().decode("utf-8", "ignore"))
    except Exception:
        return None
    finally:
        socket.setdefaulttimeout(prev)


def _version_info_from(hb: dict | None = None, public: dict | None = None) -> dict:
    """合并心跳 / 公共端点的版本字段；缺失则 update_available=False。"""
    src: dict = {}
    if public:
        src.update(public)
    if hb:
        src.update(hb)            # 心跳字段优先
    latest = src.get("latest_version") or src.get("version")
    url = src.get("download_url")
    min_v = src.get("min_version")
    update_available = False
    if latest:
        try:
            update_available = _parse_version_tuple(latest) > _current_version_tuple()
        except Exception:
            update_available = False
    return {
        "latest_version": latest,
        "download_url": url,
        "min_version": min_v,
        "update_available": update_available,
    }


# ---------------------------------------------------------------------------
# 公开 API
# ---------------------------------------------------------------------------

_ACTIVATE_ERR = {
    "invalid_code": "Activation code is invalid or does not belong to this product",
    "revoked": "This activation code has been revoked (refund-locked)",
    "expired": "This activation code has expired",
    "already_used": "This activation code is already bound to another machine (lifetime/weekly/monthly plans are limited to one machine)",
    "missing_fields": "Request fields missing (product or activation code is empty)",
}


def activate(code: str) -> dict:
    """输入激活码，向中台验证并写入本地状态。返回 {ok, kind, message}。"""
    code = (code or "").strip()
    if not code:
        return {"ok": False, "message": "Activation failed: please enter an activation code first."}
    try:
        res = _check_mid_platform(code)
    except Exception as e:
        return {"ok": False, "message": f"Activation failed: {e} (retry later if the network is unavailable)"}
    if not res.get("ok"):
        err = res.get("error", "")
        return {"ok": False,
                "message": f"Activation failed: {_ACTIVATE_ERR.get(err, err or 'Server rejected the request')}"}
    state = load_state()
    state["activated"] = True
    state["code"] = code
    state["kind"] = res.get("type") or res.get("kind") or KIND_LIFETIME
    # 到期时间（毫秒，Unix epoch ms）：以中台下发的 expires_at 为准；
    # 过渡期服务端若未下发，季/年卡按 kind 兜底推导，避免本地永远不过期
    # （年卡续费后由心跳写回新 expires_at，见 refresh_license_on_launch）。
    if state["kind"] in (KIND_QUARTERLY, KIND_ANNUAL):
        srv_exp = res.get("expires_at")
        if srv_exp:
            state["expires_at"] = srv_exp
        else:
            dur_ms = (90 if state["kind"] == KIND_QUARTERLY else 365) * 24 * 3600 * 1000
            state["expires_at"] = int(time.time() * 1000) + dur_ms
    else:
        # 终身 / 次卡：无绝对到期（次卡靠 uses_used 计次，不靠时间）
        state["expires_at"] = res.get("expires_at")
    # 次卡：按次计费字段，以中台为准；兜底 2 次 / 0 已用
    if state["kind"] == KIND_TIMES:
        state["uses_total"] = res.get("uses_total") if res.get("uses_total") is not None else 2
        state["uses_used"] = res.get("uses_used") or 0
    state["revoked"] = False
    state["machine"] = _machine_fingerprint()
    state["last_heartbeat"] = int(time.time())
    try:
        save_state(state)
    except Exception as e:
        return {"ok": False, "message": f"Activation failed: cannot write local state ({e})"}
    _safe_trial_sync(_machine_fingerprint(), claim=True)
    return {"ok": True, "kind": state["kind"], "message": "Activation successful"}


# ---------------------------------------------------------------------------
# 兜底方案：离线备用码（Ed25519 非对称签名，见 src/license/crypto.py + ed25519_verify.py）
# ---------------------------------------------------------------------------
# 算法与统一发码器 reedcode_unified.py 完全一致（machine_code|时间戳 → Ed25519 签名）。
# 与导师版 / 学生版密钥隔离（各自独立 Ed25519 密钥对），保证跨产品码不互通。
# 客户把本机机器码发给卖家，卖家用统一发码器（持本机私钥）生成离线码，客户在激活页
# 选「离线激活」粘贴即可。私钥只在卖家本机（_signing_keys/），公开仓库 / 反编译客户端都拿不到。
from .crypto import verify_offline_code


def activate_offline(code: str) -> dict:
    """离线激活：校验离线码（绑定本机机器码）→ 写入本地状态（终身、离线、不联网）。

    返回 {ok, kind, message}。离线激活后授权为终身且永久离线，不再触网（隐私：离线版不回传）。
    """
    code = (code or "").strip()
    if not code:
        return {"ok": False, "message": "Offline activation code is empty"}
    fp = _machine_fingerprint()
    if not verify_offline_code(code, fp):
        return {"ok": False, "message":
                "Offline activation code is invalid, or does not belong to this machine (make sure it was generated from this machine's code)."}
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
        return {"ok": False, "message": f"Offline activation failed: cannot write local state ({e})"}
    return {"ok": True, "kind": "lifetime",
            "message": "Offline activation successful (lifetime, no internet required)."}


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
    except Exception as e:   # noqa: BLE001
        return _entitlement_on_error(e)


def _entitlement_on_error(e: Exception) -> dict:
    """授权模块自身异常时的兜底裁决 —— **绝不能一律放行**（收费漏洞修复）。

    旧实现是"任何异常都返回 allowed=True + 一份新的试用额度"，注释写着"不阻断用户"。
    但这条分支实际是个**白嫖口子**：只要错误是持久的（授权文件所在目录不可写、磁盘满、
    状态文件内容不可能解析），**每一次**修正都会重新白送一次试用 —— 等于无限免费修正；
    篡改者还能主动制造这种异常来绕过门禁。

    新策略（与项目既定的「扣减失败也拦（防篡改）」一致，fail-closed）：

      · 本地状态明确**已激活** → 放行（不能让一次写入失败挡住已付费用户）；
      · 未激活 → 想给试用**必须先落盘记次**（trial_fix_used=True），落不进去就一律
        拦下 —— 宁可拦错一次，也不留"制造异常即白嫖"的口子；
      · 已用过试用 / 状态损坏或被篡改 → 照常拦（与正常路径的裁决保持一致）。
    """
    try:
        st = load_state()
    except Exception:
        return {"allowed": False, "reason": "state_error",
                "message": "License state cannot be read; fixing is temporarily unavailable. Please retry or contact support.",
                "remaining_trial": 0}
    if st.get("activated"):
        return {"allowed": True, "reason": "activated",
                "message": f"Activated ({st.get('kind') or 'lifetime'})",
                "remaining_trial": None}
    if (st.get("trial_fix_used")
            or st.get("_state_reason") in ("corrupt", "tampered", "unsealed")):
        return {"allowed": False, "reason": "trial_exhausted",
                "message": "Free trial has been used up. Please enter an activation code to unlock unlimited fixes.",
                "remaining_trial": 0}
    try:
        st["trial_fix_used"] = True
        save_state(st)          # 落盘成功，这一次试用才算真的给出去了
    except Exception:
        return {"allowed": False, "reason": "state_error",
                "message": "License state cannot be written; fixing is temporarily unavailable. Please retry or contact support.",
                "remaining_trial": 0}
    return {"allowed": True, "reason": "trial",
            "message": f"License state read error ({e}); allowing one free trial fix.",
            "remaining_trial": 0}


def _local_expiry_reason(state: dict) -> str | None:
    """本地有效期 / 计次强制（fail-closed）。返回锁定 reason 或 None。

    - lifetime：无期限 → None
    - times（次卡）：uses_used >= uses_total → "uses_exhausted"（离线也拦，
      与中台 /api/consume 为准，本地 owed 账本兜底防离线白嫖）
    - quarterly / annual：expires_at（毫秒）已过 → "expired"
      （年卡续费后心跳写回新 expires_at，故联网用户不会误锁）
    """
    kind = state.get("kind")
    if kind == KIND_TIMES:
        total = state.get("uses_total")
        used = state.get("uses_used") or 0
        if total is not None and used >= total:
            return "uses_exhausted"
        return None
    if kind in (KIND_QUARTERLY, KIND_ANNUAL):
        exp = state.get("expires_at")
        if exp and int(time.time() * 1000) > exp:
            return "expired"
        return None
    return None


_EXPIRY_MSG = {
    "expired": "Your license has expired. Please renew to continue fixing.",
    "uses_exhausted": "Your fix pack is used up. Please purchase a new pack to continue fixing.",
}


def _times_precheck_reason(state: dict) -> str | None:
    """次卡修正前服务端配额预检（dry_run）。返回 "uses_exhausted" 或 None。

    - 网络不可达（res None）→ 离线宽限，不拦（本地 uses_used 已拦满额）；
    - 服务端明确不 ok（如 exhausted / 多机已用完）→ 拦；
    - 其余 → 放行。
    """
    if state.get("kind") != KIND_TIMES:
        return None
    res = _consume_use(dry_run=True)
    if res is None:
        return None
    if not res.get("ok"):
        return "uses_exhausted"
    return None


def _require_fix_entitlement() -> dict:
    global _launch_verified_active
    state = load_state()
    fp = _machine_fingerprint()

    # ① 机器绑定：状态文件来自别的机器 → 不认激活（在线可凭激活码重新绑定）
    if state.get("_state_ok") and state.get("machine") and state["machine"] != fp:
        hb = _safe_heartbeat(state.get("code"), fp)
        if hb and hb.get("ok") and hb.get("status") == "active":
            state["machine"] = fp
            state["last_heartbeat"] = int(time.time())
            # 写盘失败不能让整个裁决炸出去（旧实现会一路抛到 require_fix_entitlement
            # 的兜底分支、被当成"异常即放行" → 免费无限修正）
            try:
                save_state(state)
            except Exception:
                pass
        else:
            state["activated"] = False

    # ①-b 离线授权：已离线激活 → 直接放行，绝不联网（隐私：离线版不回传）。
    #      但**绝不盲信本地布尔字段**：手写 {"activated":true,"offline":true} 就能
    #      永久白嫖。所以这里重新验签——没有有效离线码就不放行（fail-closed）。
    if state.get("activated") and state.get("offline"):
        if verify_offline_code(state.get("code") or "", fp):
            return {"allowed": True, "reason": "activated",
                    "message": f"Offline activated ({state.get('kind') or 'lifetime'})",
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
    #        ① 离线码：本地 Ed25519 验签（纯标准库，不联网，密码学上等价于 ①-b）；
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
                    "message": (f"Offline activated ({kind})" if legacy_offline
                                else f"Activated ({kind})"),
                    "remaining_trial": None}
        return {"allowed": False, "reason": "needs_reactivation",
                "message": "A legacy license file was detected (activated before this upgrade). "
                           "Please re-enter your activation code once to restore it: paste an offline code directly, or connect online for an online code.",
                "remaining_trial": 0}

    # ② 已激活：启动已一次性核验过（内存可信标志）→ 直接放行，修正时不再触网
    #    （用户决策：启动核对一次即可，有效状态就不再打扰网络）。
    #    启动未核验（如离线启动 / 核验线程还没回）→ 本次补一次心跳；
    #    服务器明确 revoked/expired 才锁；网络不可达（_safe_heartbeat 返回 None）
    #    走离线宽限、不锁。离线激活码已在 ①-b 提前返回，不会到这里。
    if state.get("activated"):
        # ② 本地有效期 / 计次强制（fail-closed，离线也拦）
        _exp_reason = _local_expiry_reason(state)
        if _exp_reason:
            return {"allowed": False, "reason": _exp_reason,
                    "message": _EXPIRY_MSG.get(_exp_reason, "License no longer valid."),
                    "remaining_trial": 0}
        # 次卡：修正前服务端配额预检（dry_run，离线宽限不拦）
        _q_reason = _times_precheck_reason(state)
        if _q_reason:
            return {"allowed": False, "reason": _q_reason,
                    "message": _EXPIRY_MSG.get(_q_reason, "License no longer valid."),
                    "remaining_trial": 0}
        if _launch_verified_active:
            return {"allowed": True, "reason": "activated",
                    "message": f"Activated ({state.get('kind') or 'lifetime'})",
                    "remaining_trial": None}
        if _apply_heartbeat_revocation(state, fp):
            return {"allowed": False, "reason": "revoked",
                    "message": "Your license has been revoked or expired. Please re-enter your activation code.",
                    "remaining_trial": 0}
        # 心跳已核对（active，已带回可能刷新的 expires_at）→ 本地有效期 / 计次强制
        _launch_verified_active = True
        return {"allowed": True, "reason": "activated",
                "message": f"Activated ({state.get('kind') or 'lifetime'})",
                "remaining_trial": None}

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
                "message": "Free trial has been used up. Please enter an activation code to unlock unlimited fixes.",
                "remaining_trial": 0}
    return {"allowed": True, "reason": "trial",
            "message": f"First free trial ({TRIAL_FIX_LIMIT} free use available)",
            "remaining_trial": 1}


def refresh_license_on_launch() -> dict:
    """App 启动调用（后台线程，非阻塞）：一次核对授权 + 版本（方向 A + 启动一次性验证）。

    返回 {checked, activated, revoked, locked, offline, network_unreachable,
          version:{latest_version, download_url, min_version, update_available}}。
    - 离线激活码（Ed25519）：offline=True，不锁（离线宽限）；版本信息来自公共端点。
    - 在线激活码：心跳核对吊销；active → 置位内存可信标志（之后修正免心跳）；带回版本信息。
    - 网络不可达：network_unreachable=True，离线宽限不锁（GUI 据此给「未能联网」提示）。
    - 未激活（试用）：不做授权核对（无码可核），但公共版本端点仍给更新提示。
    """
    global _launch_verified_active
    state = load_state()
    version = _version_info_from(public=_check_latest_version())  # 公共端点，人人可用

    if not state.get("activated"):
        return {"checked": False, "activated": False,
                "revoked": bool(state.get("revoked")),
                "locked": False, "offline": False,
                "network_unreachable": False, "version": version}

    if state.get("offline"):
        # 离线授权：永不触网（隐私），本地可信、不锁；版本信息来自公共端点（通了就用）。
        return {"checked": False, "activated": True,
                "revoked": False, "locked": False, "offline": True,
                "network_unreachable": False, "version": version}

    hb = _safe_heartbeat(state.get("code"), _machine_fingerprint())
    if hb is None:
        # 网络不可达：离线宽限，不锁；版本信息已尝试（可能也无）。
        _launch_verified_active = False
        return {"checked": False, "activated": True,
                "revoked": False, "locked": False, "offline": True,
                "network_unreachable": True, "version": version}
    if hb.get("status") in ("revoked", "expired"):
        state["activated"] = False
        state["revoked"] = True
        try:
            save_state(state)
        except Exception:
            pass
        _launch_verified_active = False
        return {"checked": True, "activated": False, "revoked": True,
                "locked": True, "offline": False,
                "network_unreachable": False,
                "version": _version_info_from(hb, version)}
    # active：写回服务端下发的过期时间 / 计次（年卡续费自动延长、次卡多机同步）
    state["last_heartbeat"] = int(time.time())
    _sync_from_heartbeat(state, hb)
    try:
        save_state(state)
    except Exception:
        pass
    _launch_verified_active = True
    pay_owed()  # 联网恢复：次卡欠账补扣（绝不抛异常）
    return {"checked": True, "activated": True, "revoked": False,
            "locked": False, "offline": False,
            "network_unreachable": False,
            "version": _version_info_from(hb, version)}


def record_fix_used() -> bool:
    """修正成功后调用：试用计次 / 次卡扣次 + 欠账兜底。绝不抛异常。返回是否落盘成功。

    - 未激活：试用标记 + 中台登记（原逻辑）；
    - 次卡（times）：乐观推进本地 uses_used，再上报中台真扣；上报成功以服务端为准并
      结清 owed 欠账，失败/离线记 owed，下次联网由 pay_owed() 补扣（中台为准，防离线白嫖）；
    - 其他已激活（lifetime / 季 / 年）：无计次，跳过。
    """
    ok = True
    state = None
    try:
        state = load_state()
        if not state.get("activated"):
            state["trial_fix_used"] = True
        elif state.get("kind") == KIND_TIMES:
            # 先乐观推进本地计数，再上报；上报成功以服务端为准并结清欠账
            state["uses_used"] = (state.get("uses_used") or 0) + 1
            res = _consume_use(dry_run=False)
            if res and res.get("ok"):
                if res.get("uses_used") is not None:
                    state["uses_used"] = res["uses_used"]
                state["owed"] = 0
            else:
                # 网络不可达 / 服务端异常：记欠账，下次联网补扣
                state["owed"] = (state.get("owed") or 0) + 1
        save_state(state)
    except Exception:
        ok = False
    if state is not None and not state.get("activated"):
        _safe_trial_sync(_machine_fingerprint(), claim=True)
    return ok


def post_fix_message(gate: dict, counted: bool = True) -> str:
    """修正**成功之后**该展示的状态文案。

    与 ``gate["message"]`` 的区别：gate 是修正**前**的裁决，此时试用额度尚未扣减，
    直接展示会出现「刚用完免费试用，却还显示首次免费试用」的错位。
    """
    if not counted:
        return "Trial count write failed (this fix still succeeded, but please note)."
    if gate.get("reason") == "activated":
        s = load_state()
        if s.get("kind") == KIND_TIMES:
            total = s.get("uses_total")
            used = s.get("uses_used") or 0
            left = (total - used) if total is not None else None
            if left is not None and left <= 0:
                return "Two-use pack used up. Please purchase a new pack to continue fixing."
            return f"Two-use pack · {left} fix(es) left."
        return f"Activated ({gate.get('kind') or 'lifetime'}), you can continue fixing without limits."
    return (f"Free trial quota exhausted ({TRIAL_FIX_LIMIT}/{TRIAL_FIX_LIMIT}); "
            f"next fix requires an activation code.")


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


# 授权类型对外展示标签（全英文，海外版零中文）。
_KIND_LABELS = {
    KIND_TIMES: "Two-use pack",
    KIND_QUARTERLY: "3-month pass",
    KIND_ANNUAL: "Annual plan",
    KIND_LIFETIME: "Lifetime license",
}


def license_summary() -> dict:
    """供 GUI 展示的授权摘要：类型标签、剩余天数 / 剩余次数、预警与耗尽标志。

    阈值（与国内版对齐）：
    - 季 / 年卡：剩余 ≤ 0 天 → exhausted；≤ 7 天 → warn（近到期软提醒）；
    - 次卡：剩余 ≤ 0 次 → exhausted；≤ 1 次 → warn（低额度加购提示）。
    """
    s = load_state()
    kind = s.get("kind")
    now_ms = int(time.time() * 1000)
    out = {
        "activated": bool(s.get("activated")),
        "kind": kind,
        "kind_label": _KIND_LABELS.get(kind, "Free trial"),
        "offline": bool(s.get("offline")),
        "revoked": bool(s.get("revoked")),
        "trial_fix_used": bool(s.get("trial_fix_used")),
        "expires_at": s.get("expires_at"),
        "days_left": None,
        "uses_total": s.get("uses_total"),
        "uses_used": s.get("uses_used") or 0,
        "uses_left": None,
        "owed": s.get("owed") or 0,
        "warn": False,       # 临近到期 / 仅剩 1 次
        "exhausted": False,  # 已到期 / 次数用尽
    }
    if kind in (KIND_QUARTERLY, KIND_ANNUAL):
        exp = s.get("expires_at")
        if exp:
            days = (exp - now_ms) // 86400000
            out["days_left"] = days
            if days <= 0:
                out["exhausted"] = True
            elif days <= 7:
                out["warn"] = True
    elif kind == KIND_TIMES:
        total = s.get("uses_total")
        used = s.get("uses_used") or 0
        if total is not None:
            left = total - used
            out["uses_left"] = left
            if left <= 0:
                out["exhausted"] = True
            elif left <= 1:
                out["warn"] = True
    return out
