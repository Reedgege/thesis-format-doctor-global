# -*- coding: utf-8 -*-
"""离线激活的非对称（Ed25519）验签 · 海外版
================================================

为什么是 Ed25519：
    旧方案离线码用**对称 HMAC**——验签密钥 == 签名密钥，必须塞进客户端，
    于是只能写成源码明文常量。仓库公开后任何人都能伪造无限离线激活码。

    现在改成 **Ed25519 非对称签名（RFC 8032）**：
      * 私钥：只在卖家本机（`_signing_keys/overseas_ed25519_private.pem`，
        由统一发码器 reedcode_unified.py 持有，绝不入库、绝不进安装包）。
        发码工具用它给客户的机器码签名。
      * 公钥：本文件里的 ``PUBLIC_KEY`` 常量（32 字节原始公钥，base64 编码存储），
        编译进客户端二进制。公开也无害——它只能「验签」，不能「签名」。
        哪怕把整个仓库 / 反编译客户端公开，没有私钥就签不出任何一个能过验签的离线码。

    关键安全/误报收益：客户端**只验签、不签名**，且验签是纯标准库实现
    （见 ed25519_verify.py，只用到 hashlib + 模运算，没有 decrypt / 私钥解析 /
    密钥生成 任何代码）。这彻底消除了旧 RSA 方案把整个 ``rsa`` 库（含 decrypt /
    生成密钥 等字符串）打进 exe 后撞上杀软「文件锁/勒索」启发式（HEUR:Ransom/
    LockFile.a）的问题——Ed25519 是签名算法，本就没有「解密」概念。

离线码格式（与统一发码器 reedcode_unified.py 共用）：
    ``<machine_code>|<issued_at_unix>.<base64urlsafe(Ed25519 签名 64 字节)>``
    验签时：① 用内嵌公钥验签名；② 校验前缀是本机机器码（绑定设备）。
"""
import base64

from .ed25519_verify import verify as _ed25519_verify

# 内嵌公钥（32 字节原始公钥，base64 存储）。由 gen_ed25519_keys.py 生成，
# 与 _signing_keys/overseas_ed25519_private.pem 配对。私钥绝不出现在此文件或安装包。
PUBLIC_KEY_B64 = "yPcze9RggoAEDF6tezfnRBbHbj0NQJHyCB3XGP+aAjg="
PUBLIC_KEY = base64.b64decode(PUBLIC_KEY_B64)

# 测试钩子：pytest 用临时公钥覆盖内嵌公钥，避免依赖卖家私钥（CI 上也没有私钥）。
_TEST_PUB = None


def set_verify_public_key(pub):
    """仅供测试：用临时 32 字节公钥覆盖内嵌公钥。传 None 恢复默认。"""
    global _TEST_PUB
    _TEST_PUB = pub


def _public_key() -> bytes:
    return _TEST_PUB if _TEST_PUB is not None else PUBLIC_KEY


def b64u(b: bytes) -> str:
    return base64.urlsafe_b64encode(b).decode("ascii")


def ub64u(s: str) -> bytes:
    """base64url 解码 —— 容忍缺失的 '=' 填充与标准 base64 的 +/ 字符。

    宽松解码可避免「发码器未补 padding 就整批离线码失效」这种静默故障。
    """
    s = s.strip().replace("+", "-").replace("/", "_")
    s += "=" * (-len(s) % 4)
    return base64.urlsafe_b64decode(s.encode("ascii"))


def verify_offline_code(code: str, machine_code: str) -> bool:
    """校验离线码：① 内嵌公钥验签通过；② 绑定本机机器码。

    **对任意畸形 / 伪造输入都必须返回 False，绝不抛异常**（在 GUI 回调里被调用，
    无控制台（打包版）下异常会被静默吞掉，用户只会看到「点了没反应」）。
    """
    if not isinstance(code, str) or not isinstance(machine_code, str):
        return False
    code = code.strip()
    if not code or "." not in code:
        return False
    try:
        payload, _, sig_b64 = code.rpartition(".")
        if not payload.startswith(machine_code + "|"):
            return False
        sig = ub64u(sig_b64)
        msg = payload.encode("utf-8")          # 被签名的消息 = "机器码|时间戳"
        return _ed25519_verify(_public_key(), msg, sig)
    except Exception:
        return False
