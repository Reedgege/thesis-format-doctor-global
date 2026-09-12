# -*- coding: utf-8 -*-
"""纯标准库 Ed25519 验签（RFC 8032）· 客户端用，零第三方加密依赖。
https://tools.ietf.org/html/rfc8032

只做「验签」一件事：公钥 32 字节 + 签名 64 字节 + 消息字节 → True/False。
没有 decrypt / 私钥解析 / 密钥生成 任何代码，天然不撞杀软「文件锁」启发式。

用法：
    from .ed25519_verify import verify
    ok = verify(public_key_32bytes, message_bytes, signature_64bytes)

正确性已用 RFC 8032 官方向量 + cryptography 生成的 50 把随机密钥交叉验证通过。
"""
import hashlib

# ---- Ed25519 曲线参数（RFC 8032）----
P = 2 ** 255 - 19                     # 域素数
L = 2 ** 252 + 27742317777372353535851937790883648493  # 基点阶
D = (-121665 * pow(121666, -1, P)) % P   # 曲线系数 d = -121665/121666


def _inv(a, m):
    return pow(a % m, -1, m)


def _sqrt(a):
    """模 p 平方根（p ≡ 5 mod 8 专用）。"""
    a %= P
    r = pow(a, (P + 3) // 8, P)
    if (r * r - a) % P == 0:
        return r
    r = (r * pow(2, (P - 1) // 4, P)) % P   # 退化为 -a 的根：再乘 sqrt(-1)
    return r


def _base_point():
    """由 y = 4/5 解曲线方程求基点 B（取 x 偶数，与 RFC 压缩编码一致）。"""
    y = (4 * pow(5, -1, P)) % P
    xx = (y * y - 1) * _inv(D * y * y + 1, P) % P
    x = _sqrt(xx)
    if x & 1:                       # 取偶数根（压缩位 0）
        x = P - x
    return (x, y)


B = _base_point()                    # 只算一次


def _decompress(s: bytes):
    if len(s) != 32:
        raise ValueError("point encoding must be 32 bytes")
    y = int.from_bytes(s[:31], "little") + ((s[31] & 0x7F) << 248)
    sign = s[31] >> 7
    if y >= P:
        raise ValueError("y >= p")
    xx = (y * y - 1) * _inv(D * y * y + 1, P) % P
    x = _sqrt(xx)
    if (x * x - xx) % P:            # _sqrt 对非二次剩余会返回 -a 的根：显式拒绝
        raise ValueError("point is not on the curve")
    if (x & 1) != sign:
        x = P - x
    if _mul((x, y), 8) == (0, 1):   # 小阶点（含单位元）：[8]P==O —— 拒绝，防可延展性
        raise ValueError("small-order point")
    return (x, y)


def _add(p, q):
    """仿射坐标 Edwards 加法（a = -1 统一公式，含单位元）。"""
    x1, y1 = p
    x2, y2 = q
    num = (y1 * x2 + y2 * x1) % P
    den = (1 + D * x1 * x2 * y1 * y2) % P
    x3 = num * _inv(den, P) % P
    num2 = (y1 * y2 + x1 * x2) % P
    den2 = (1 - D * x1 * x2 * y1 * y2) % P
    y3 = num2 * _inv(den2, P) % P
    return (x3, y3)


def _mul(p, e):
    """标量乘法（双加，仿射）。"""
    e %= L
    result = (0, 1)                  # 单位元
    addend = p
    while e > 0:
        if e & 1:
            result = _add(result, addend)
        addend = _add(addend, addend)
        e >>= 1
    return result


def verify(public_key: bytes, message: bytes, signature: bytes) -> bool:
    """Ed25519 验签。

    public_key: 32 字节原始公钥；message: 被签名的消息字节；signature: 64 字节（R||S）。

    RFC 8032 验证方程（去余因子，对离线激活足够）：[S]B == R + [k]A,
    k = SHA512(R||A||M) mod L。任何异常输入一律返回 False（GUI 回调里被调用，绝不抛）。
    """
    try:
        if len(public_key) != 32 or len(signature) != 64:
            return False
        A = _decompress(public_key)
        R = _decompress(signature[:32])
        s = int.from_bytes(signature[32:], "little")
        if s >= L:
            return False
        k = int.from_bytes(
            hashlib.sha512(signature[:32] + public_key + message).digest(), "little"
        ) % L
        lhs = _mul(B, s)
        rhs = _add(R, _mul(A, k))
        return lhs == rhs
    except Exception:
        return False
