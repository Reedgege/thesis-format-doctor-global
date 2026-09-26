# -*- coding: utf-8 -*-
"""构建后自检：运行时资源真的进了产物吗？（海外版）
==================================================================
用法：``python tools/verify_bundle.py [dist_dir]``（默认 ``dist``）

为什么需要它
------------
``if os.path.isfile(资源) else 静默跳过`` 这类守卫会把"资源没进包"变成**完全无声的
功能缺失**。姊妹项目（导师版 tfd-mentor）就真栽过：``build.py`` 漏了数据目录 →
打包版「二维码不显示、水印图片丢失」却能一路绿灯发出去。

海外版同样存在这类路径依赖：
  * ``src/data/icon.png``                      窗口图标（``src/ui/iconpath.py`` 第一候选）
  * ``assets/icon.ico`` / ``assets/icon.png``  Windows ``iconbitmap`` 兜底 / exe 同级图标
  * ``assets/icon.icns``                       macOS 图标
  * ``src/data/ai_questionnaire_template.yaml`` CLI ``export-schema`` 与 GUI 导出模板
  * ``VERSION``                                运行时版本号（``src/versioninfo.py``）

  .. note:: v2.3.4 起界面背景改为整窗纯色，不再依赖 ``backdrop.png``，故该资源已从
           强制清单移除（纯装饰、缺失也不再影响功能）。

按**文件名**匹配（跨平台都成立：Linux 文件夹、Windows 文件夹、macOS ``.app`` 内部布局
各不相同，写死相对路径反而会误报）。缺任何一个 → 退出码 1，让 CI 红灯。
"""
import os
import sys

# Windows runner 控制台默认 cp1252：print 中文（全角冒号等）会 UnicodeEncodeError，
# 反而让自检自己把构建搞红。与 build.py 同款兜底：强制 UTF-8、编不出的字符替换。
for _stream in ("stdout", "stderr"):
    try:
        getattr(sys, _stream).reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

REQUIRED_FILES = (
    "icon.ico",
    "icon.png",
    "icon.icns",
    "ai_questionnaire_template.yaml",
    "VERSION",
)


def main(dist="dist"):
    if not os.path.isdir(dist):
        print("找不到产物目录：%s" % dist)
        return 1
    found = set()
    for _root, _dirs, files in os.walk(dist):
        found.update(files)
    missing = [n for n in REQUIRED_FILES if n not in found]
    if missing:
        print("产物缺少运行时资源：%s" % ", ".join(missing))
        print("（检查 build.py 的 --include-data-dir / --include-package-data 是否生效）")
        return 1
    print("OK: %d runtime assets bundled into %s" % (len(REQUIRED_FILES), dist))
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1] if len(sys.argv) > 1 else "dist"))
