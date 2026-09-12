# -*- coding: utf-8 -*-
"""离线激活的非对称（RSA）加解密 · 海外版
================================================

为什么是非对称：
    旧方案离线码用**对称 HMAC**——验签密钥 == 签名密钥，必须同时塞进客户端，
    于是它只能写成源码里的明文常量。仓库公开后，任何人 fork 都能拿到密钥、伪造
    无限离线激活码。

    现在改成 **RSA 非对称签名**：
      * 私钥：只在卖家本机（仓库根 `private_key.pem`，已 gitignore，绝不进安装包）。
        发码工具用它给客户的机器码签名。
      * 公钥：本文件里的 ``PUBLIC_KEY_PEM`` 常量，编译进客户端二进制。公开也无害——
        它只能「验签」，不能「签名」。所以哪怕把整个仓库 / 反编译客户端公开，
        没有私钥就签不出任何一个能通过验证的离线码。

    这样公开仓库里**不存在任何能伪造激活码的密钥**，彻底关闭泄露口子。

离线码格式（与 tools/reedcode*.py 共用）：
    ``<machine_code>|<issued_at_unix>.<base64(RSA-SHA256 签名)>``
    验签时：① 用内嵌公钥验签名；② 校验前缀是本机机器码（绑定设备）。
"""
import base64
import os
import time

import rsa

# 内嵌公钥（PKCS#1 PEM）。由 tools/gen_keys.py 生成，可入库的 public_key.pem 与之相同。
# 私钥绝不出现在此文件或任何安装包里。
PUBLIC_KEY_PEM = """-----BEGIN RSA PUBLIC KEY-----
MIIBCgKCAQEAhFJwSG+B2lh5+nf/aB9sJQmOxbUMlyxHj1nqlEkAzJfRXfNPi9mm
ABVzJSV0umzZ4d/goU3FejObhfsjmKEkOWFtA4iBhVuB9XrbDDF2vcLHsqDFmpX7
nxw6DzxCQzPGMMSlTqD/k2mUZtg8MLtnUoD/k09r3paC7ZstnEl6MQSwTCqq6+Oq
tGtoTZwVlkQiCo3LgFMIXUvQ5qw5UQbr6kZBHaI/Xl4U3gTWxKyG9o6C2SqrvH3g
15KSkIX8ZbCKJr9mhK5Su14xt3dtDji+HdMrE3aNUmbub6Ud5R+fgfz/l2bJ3QGE
QdwW1zXfHfE9BriXOsms5tCvEK1p3uQ3CwIDAQAB
-----END RSA PUBLIC KEY-----
"""

# 测试钩子：pytest 用临时密钥对覆盖公钥，避免依赖卖家私钥（CI 上也没有私钥）。
_TEST_PUB = None


def set_verify_public_key(pub):
    """仅供测试：用临时公钥覆盖内嵌公钥。传 None 恢复默认。"""
    global _TEST_PUB
    _TEST_PUB = pub


def _public_key():
    if _TEST_PUB is not None:
        return _TEST_PUB
    return rsa.PublicKey.load_pkcs1(PUBLIC_KEY_PEM.encode("utf-8"), "PEM")


def b64u(b: bytes) -> str:
    return base64.urlsafe_b64encode(b).decode("ascii")


def ub64u(s: str) -> bytes:
    return base64.urlsafe_b64decode(s.encode("ascii"))


def load_private_key(path: str = None) -> "rsa.PrivateKey":
    """加载卖家私钥。默认读取仓库根目录 private_key.pem（仅本机持有）。

    找不到时抛出清晰的错误，而不是让 GUI 静默崩溃。
    """
    if path is None:
        here = os.path.dirname(os.path.abspath(__file__))
        repo_root = os.path.abspath(os.path.join(here, "..", ".."))
        path = os.path.join(repo_root, "private_key.pem")
    if not os.path.isfile(path):
        raise FileNotFoundError(
            "未找到卖家私钥：%s\n卖家发码工具需要本机私钥（private_key.pem）。"
            "该文件只存于你的电脑，绝不入库、绝不进安装包。"
            "用 `python tools/gen_keys.py` 在你机器上生成一次即可。" % path)
    with open(path, "rb") as fh:
        return rsa.PrivateKey.load_pkcs1(fh.read(), "PEM")


def sign_offline(private_key: "rsa.PrivateKey", machine_code: str) -> str:
    """machine_code + 时间戳 → RSA 签名 → 离线码（与 verify_offline_code 兼容）。"""
    ts = int(time.time())
    payload = "%s|%d" % (machine_code, ts)
    sig = rsa.sign(payload.encode("utf-8"), private_key, "SHA-256")
    return payload + "." + b64u(sig)


def generate_offline_code(machine_code: str, private_key_path: str = None) -> str:
    """卖家发码工具用：用本机私钥给客户机器码签名生成离线码。"""
    return sign_offline(load_private_key(private_key_path), machine_code)


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
        # rsa.verify 验签失败会抛 rsa.pkcs1.VerificationError；成功返回哈希名（忽略）。
        rsa.verify(payload.encode("utf-8"), sig, _public_key())
        return True
    except Exception:
        return False
