# -*- coding: utf-8 -*-
"""生成离线激活用的 RSA 密钥对（每个产品独立一组）。

用法：
    python tools/gen_keys.py

产物（都在仓库根目录）：
    private_key.pem  —— 卖家私钥，**仅留本机**，已被 .gitignore 忽略，绝不入库、绝不进安装包
    public_key.pem   —— 公钥，可入库作参考；真正使用的是嵌入进 crypto.py 的 PUBLIC_KEY_PEM 常量

离线激活改成「非对称」后：
    * 私钥只卖家持有 → 任何人（含 fork 公开仓库）都无法伪造离线码
    * 公钥嵌进客户端 → 只能验签，不能签名；公开无害
"""
import os

import rsa

HERE = os.path.dirname(os.path.abspath(__file__))
REPO_ROOT = os.path.abspath(os.path.join(HERE, ".."))


def main():
    print("生成 RSA-2048 密钥对 ...")
    pub, priv = rsa.newkeys(2048)
    priv_pem = priv.save_pkcs1("PEM").decode("utf-8")
    pub_pem = pub.save_pkcs1("PEM").decode("utf-8")

    priv_path = os.path.join(REPO_ROOT, "private_key.pem")
    pub_path = os.path.join(REPO_ROOT, "public_key.pem")
    with open(priv_path, "w", encoding="utf-8") as f:
        f.write(priv_pem)
    with open(pub_path, "w", encoding="utf-8") as f:
        f.write(pub_pem)

    print("已写入：")
    print("  %s  (私钥，务必仅留本机，已 gitignore)" % priv_path)
    print("  %s  (公钥，可入库参考)" % pub_path)
    print("下一步：把 public_key.pem 的内容原样粘进 src/license/crypto.py 的 PUBLIC_KEY_PEM 常量。")


if __name__ == "__main__":
    main()
