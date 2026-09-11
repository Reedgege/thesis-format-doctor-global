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
安全：离线码用对称签名（密钥与 src/license/license.py 共用，明文常量）。
"""
import os
import sys
import time
import hmac
import hashlib


# 与 src/license/license.py 的 _OFFLINE_KEY 完全一致；与国内版密钥不同，防互用。
_OFFLINE_KEY = b"tfdglobal|kami|offline|2026|sign|v1"


def _offline_sign(payload):
    return hmac.new(_OFFLINE_KEY, payload.encode("utf-8"), hashlib.sha256).hexdigest()[:24]


def generate_offline_code(machine_code):
    """machine_code + 时间戳 + 签名 -> 离线码（与主程序 verify 兼容）。"""
    ts = int(time.time())
    payload = "%s|%d" % (machine_code, ts)
    return payload + "|" + _offline_sign(payload)


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
    code = generate_offline_code(target)
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
