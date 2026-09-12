# -*- coding: utf-8 -*-
"""
reedcode · Thesis Format Doctor Global（海外版）离线发码工具（兜底用）
======================================================================
仅在「中台跑路 / 宕机 / 客户无法联网」时给客户重激活使用。
用法：
  1) 双击本程序，按提示粘贴客户机器码，回车即得离线码；
  2) 或命令行：reedcode.exe "<客户机器码>"  直接输出离线码。
把离线码发给客户，他在激活页选「离线激活」粘贴即可。
自动记入 reedcode_log.txt 台账。

安全：离线码用 **RSA 非对称签名**（私钥只在你本机 private_key.pem，绝不进仓库/安装包；
公钥嵌进客户端，只能验签不能签名）。详见 src/license/crypto.py。
"""
import os
import sys
import time

# 让 tools/ 能 import 到 src/（无需在仓库根跑命令）
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.abspath(os.path.join(HERE, "..")))

from src.license.crypto import generate_offline_code


def main():
    # 日志写到 exe 所在目录，避免临时目录被清后找不到台账
    HERE = os.path.dirname(os.path.abspath(sys.argv[0]))
    if len(sys.argv) > 1:
        target = " ".join(sys.argv[1:]).strip()
    else:
        target = input("粘贴客户的机器码，回车生成离线码：").strip()
    if not target:
        print("未提供机器码，已退出。")
        sys.exit(1)
    try:
        code = generate_offline_code(target)
    except Exception as e:
        print("\n生成离线码失败：%s" % e)
        print("请确认本机存在 private_key.pem（卖家私钥），且未误删。")
        sys.exit(1)
    print("\n生成的离线备用码（复制发给客户）：")
    print(code)
    try:
        log_path = os.path.join(HERE, "reedcode_log.txt")
        with open(log_path, "a", encoding="utf-8") as f:
            f.write("%s | 客户机器码 %s | 离线码 %s\n" % (
                time.strftime("%Y-%m-%d %H:%M"), target, code))
        print("\n已记入台账 reedcode_log.txt（%s）" % log_path)
    except Exception:
        pass
    if len(sys.argv) == 1:
        input("\n按回车退出...")


if __name__ == "__main__":
    main()
