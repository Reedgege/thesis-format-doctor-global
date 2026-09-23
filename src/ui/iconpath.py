"""窗口图标的路径解析（纯路径逻辑，**不依赖 tkinter**）。

单独成模块有两个好处：

* 可以在没有 Tk 的环境（本机 venv、CI 的 lint 步骤）里单元测试路径解析；
* ``app.py`` 只需要「给我一个存在的 PNG 路径」，不关心打包布局。

打包后的目录布局随平台/构建方式变化，所以这里把几种可能都列出来按序探测：

===========  =======================================================
运行方式      图标可能的位置
===========  =======================================================
源码运行      ``src/data/icon.png``（相对本模块）
Nuitka/Win    ``<exe 目录>/assets/icon.png``（build.py 的 include-data-dir）
Nuitka/macOS  ``<.app/Contents/MacOS>/assets/icon.png``
包数据被编译  ``<exe 目录>/src/data/icon.png``（include-package-data）
===========  =======================================================

全部找不到就返回 ``None``——图标只是锦上添花，绝不能影响启动。
"""

from __future__ import annotations

import os
import sys


def _norm(path: str) -> str:
    try:
        return os.path.normpath(os.path.abspath(path))
    except Exception:
        return path


def _runtime_bases() -> list:
    """打包后「可执行文件所在的目录」候选（源码运行时基本用不上）。"""
    bases = []
    argv = getattr(sys, "argv", None) or []
    raw_list = [argv[0] if argv else "", getattr(sys, "executable", "") or ""]
    for raw in raw_list:
        try:
            if raw:
                bases.append(os.path.dirname(_norm(raw)))
        except Exception:
            continue
    try:
        bases.append(os.getcwd())
    except Exception:
        pass
    # 去重且保序
    out = []
    for b in bases:
        if b and b not in out:
            out.append(b)
    return out


def icon_candidates() -> list:
    """按「最可能命中」的顺序返回图标候选路径（含不存在的）。"""
    here = os.path.dirname(_norm(__file__))
    cands = [
        os.path.join(here, "..", "data", "icon.png"),   # 源码运行：src/ui -> src/data
        os.path.join(here, "data", "icon.png"),
    ]
    for base in _runtime_bases():
        cands += [
            os.path.join(base, "assets", "icon.png"),   # build.py 拷到 exe 同级
            os.path.join(base, "icon.png"),
            os.path.join(base, "src", "data", "icon.png"),
        ]
    out = []
    for c in cands:
        c = _norm(c)
        if c not in out:
            out.append(c)
    return out


def find_icon() -> "str | None":
    """返回第一个真实存在的图标路径；都没有则返回 None（调用方静默跳过）。"""
    for path in icon_candidates():
        try:
            if os.path.isfile(path):
                return path
        except Exception:
            continue
    return None


def _data_candidates(filename: str) -> list:
    """通用数据资源候选路径（与图标同款布局假设，按「最可能命中」排序）。"""
    here = os.path.dirname(_norm(__file__))
    cands = [
        os.path.join(here, "..", "data", filename),          # 源码运行：src/ui -> src/data
        os.path.join(here, "data", filename),
        os.path.join(here, "..", "..", "assets", filename),  # 源码运行：<repo>/assets
    ]
    for base in _runtime_bases():
        cands += [
            os.path.join(base, "assets", filename),          # build.py 拷到 exe 同级
            os.path.join(base, filename),
            os.path.join(base, "src", "data", filename),     # include-package-data
        ]
    out = []
    for c in cands:
        c = _norm(c)
        if c not in out:
            out.append(c)
    return out


def ico_candidates() -> list:
    """按「最可能命中」的顺序返回 Windows 图标（``.ico``）候选路径。

    仅用于 Windows 的 ``iconbitmap`` 兜底：个别环境的 Tcl 没编进 PNG 解码器
    （Tk 安装不完整），此时「PNG + iconphoto」会失败，而 ``iconbitmap`` 读原生
    ``.ico`` 仍然可用 —— 窗口图标不会因为一个平台的 Tcl 缺件就整个丢掉。
    """
    here = os.path.dirname(_norm(__file__))
    cands = [
        os.path.join(here, "..", "..", "assets", "icon.ico"),   # 源码运行：<repo>/assets
        os.path.join(here, "data", "icon.ico"),
    ]
    for base in _runtime_bases():
        cands += [
            os.path.join(base, "assets", "icon.ico"),           # build.py 拷到 exe 同级
            os.path.join(base, "icon.ico"),
        ]
    out = []
    for c in cands:
        c = _norm(c)
        if c not in out:
            out.append(c)
    return out


def find_icon_ico() -> "str | None":
    """返回第一个真实存在的 ``.ico``；都没有则 None（Windows 兜底用）。"""
    for path in ico_candidates():
        try:
            if os.path.isfile(path):
                return path
        except Exception:
            continue
    return None
