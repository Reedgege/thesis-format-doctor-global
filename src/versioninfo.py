# -*- coding: utf-8 -*-
"""版本号的单一读取入口（不依赖 tkinter，CLI 与 GUI 共用）。

发版只改仓库根的 ``VERSION`` 文件一处；打包时 ``build.py`` 用
``--include-data-file=VERSION=VERSION`` 把它放到产物根目录，于是运行时
在「exe 同级」也能读到 —— 不必把版本号硬编码进代码（国内版就是硬编码在
``gui.py`` 的 ``APP_VERSION``，每次发版都得记得同步两处，漏过）。

查找顺序：
  1. ``TFD_VERSION`` 环境变量（CI / 测试可注入，优先级最高）；
  2. 源码树内的 ``<repo>/VERSION``（开发态直接跑源码时命中）；
  3. 可执行文件同级的 ``VERSION``（打包后的正式形态）。

全部落空时返回 ``"dev"`` —— 版本号只是展示用，读不到也绝不能让程序起不来。
"""

from __future__ import annotations

import os
import sys

_FALLBACK = "dev"
_HERE = os.path.dirname(os.path.abspath(__file__))


def _candidate_paths() -> list:
    """按优先级列出 VERSION 文件的候选路径。"""
    cands = []
    # ① 源码树：src/versioninfo.py -> <repo>/VERSION
    try:
        cands.append(os.path.normpath(os.path.join(_HERE, "..", "..", "VERSION")))
    except Exception:
        pass
    # ② 打包产物：exe / 入口脚本所在目录
    for raw in (getattr(sys, "argv", None) or [""])[:1] + [getattr(sys, "executable", "")]:
        try:
            if raw:
                cands.append(os.path.join(os.path.dirname(os.path.abspath(raw)), "VERSION"))
        except Exception:
            continue
    # ③ 当前工作目录（用户在仓库根运行时的兜底）
    try:
        cands.append(os.path.join(os.getcwd(), "VERSION"))
    except Exception:
        pass
    return cands


def app_version() -> str:
    """返回版本号字符串（如 ``"2.0.2"``）；读不到时返回 ``"dev"``。"""
    env = os.environ.get("TFD_VERSION", "").strip()
    if env:
        return env
    for p in _candidate_paths():
        try:
            if os.path.isfile(p):
                with open(p, "r", encoding="utf-8") as fh:
                    v = fh.read().strip()
                if v:
                    return v.lstrip("vV")      # 兼容写成 "v2.0.2" 的情况
        except Exception:
            continue
    return _FALLBACK


__version__ = app_version()
