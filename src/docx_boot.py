# -*- coding: utf-8 -*-
"""修正打包后 python-docx 的「相对 __file__」——否则模板文件定位不到。

问题现象（2.3.122 实测）
------------------------
macOS / Linux 发行版里「页码修正」完全失效，而且**不报错**：fix 返回成功、
报告说页码已改，实际文档里连页眉都没创建；同一份代码在 Windows 上是好的，
在 macOS 上用源码跑也正常。

根因
----
python-docx 读自带的模板文件时，路径是从模块的 ``__file__`` 推出来的：

    path = os.path.join(os.path.split(__file__)[0], "..", "templates",
                        "default-header.xml")

Nuitka 打包后，``__file__`` 在不同平台给的形态不同：

    Windows（产物里实测）：  docx\\parts\\hdrftr.py   → 运行时解析成绝对路径，正常
    macOS / Linux（实测）：  docx/parts/hdrftr.py    → 保持相对路径

相对路径是按**进程当前工作目录**解析的。客户双击启动、或从任何别的目录调用时，
cwd 根本不是程序目录，于是 ``open()`` 抛 FileNotFoundError；而调用方
（``fixer._ensure_page_number_in_part``）用 ``except Exception`` 兜底，
异常被彻底吞掉 —— 表现就是「修了但没改，还告诉你改好了」。

受影响的都是「首次创建部件」这类操作（都要读模板）：
    · 页眉 / 页脚（页码）      → _default_header_xml / _default_footer_xml
    · 批注                     → _default_comments_xml
    · 设置 / 样式              → _default_settings_xml / _default_styles_xml

修法
----
启动时把 docx 各模块的相对 ``__file__`` 就地改成绝对路径（基准目录 = 可执行文件
所在目录，也就是 Nuitka dist 根）。只处理确实能定位到同级 ``templates`` 目录的
模块 —— 定位不到就不动，绝不硬塞一个错的路径。

源码运行时 ``__file__`` 本来就是绝对路径，本模块直接跳过，对开发环境无影响。
"""

from __future__ import annotations

import os
import sys

_applied = False

# 会去读 docx/templates 下文件的模块。必须确保它们已进入 sys.modules，
# 否则 apply() 遍历时还没导入，等后面按需导入时又是相对路径了。
_CRITICAL_SUBMODULES = (
    "docx.parts.hdrftr",     # 页眉 / 页脚：default-header.xml / default-footer.xml
    "docx.parts.comments",   # 批注：default-comments.xml
    "docx.parts.settings",   # 文档设置：default-settings.xml
    "docx.parts.styles",     # 样式：default-styles.xml
    "docx.parts.document",   # 正文部件
    "docx.parts.numbering",
    "docx.opc.packuri",
    "docx.section",
)


def _program_dir() -> str:
    """可执行文件所在目录（Nuitka dist 根）；源码运行时返回空串。"""
    exe = getattr(sys, "argv", None)
    if not exe:
        return ""
    try:
        first = exe[0]
    except Exception:
        return ""
    if not first:
        return ""
    # 源码运行：sys.argv[0] 是 .py 脚本，此时不需要任何修正。
    # "-" 是 `python -` / 交互式的写法，同理不当作程序目录。
    base_name = os.path.basename(first)
    if first.endswith(".py") or first.endswith(".pyc") or base_name in ("", "-"):
        return ""
    return os.path.dirname(os.path.abspath(first))


def _templates_dir_for(module_file: str) -> str:
    """给定 docx 某模块的 __file__，返回它对应的 templates 目录（不保证存在）。"""
    here = os.path.dirname(module_file)
    return os.path.normpath(os.path.join(here, "..", "templates"))


def apply(base_dir: str | None = None) -> int:
    """就地修正，返回处理过的模块数（0 表示无需处理）。幂等。

    base_dir 仅供测试注入：源码运行时不需要修正（__file__ 本就是绝对路径），
    但自动化测试要模拟「打包后相对 __file__」的场景，得能指定基准目录。
    """
    global _applied
    # 显式传入 base_dir 的是测试路径，不受幂等标志限制 ——
    # 否则 src/__init__.py 在 import 时已用「源码模式」跑过一次并置位，
    # 测试再想模拟打包环境就永远进不来。
    explicit = base_dir is not None
    if _applied and not explicit:
        return 0

    base = base_dir or _program_dir()
    if not base:
        _applied = True
        return 0

    # 先把会用模板的子模块拉进来，保证它们也在下面的遍历范围内
    import importlib

    for sub in _CRITICAL_SUBMODULES:
        try:
            importlib.import_module(sub)
        except Exception:
            pass

    fixed = 0
    for name, mod in list(sys.modules.items()):
        if not name or not name.startswith("docx"):
            continue
        try:
            cur = getattr(mod, "__file__", None)
        except Exception:
            continue
        if not cur or os.path.isabs(cur):
            continue
        # 打包后是相对 dist 根的路径；拼回绝对路径并确认 templates 真的在那儿，
        # 免得把一个同样错的路径塞进去，反而多一个难查的失败点。
        cand = os.path.normpath(os.path.join(base, cur))
        tpl = _templates_dir_for(cand)
        if not os.path.isdir(tpl):
            continue
        try:
            mod.__file__ = cand
        except Exception:
            continue
        fixed += 1

    _applied = True
    return fixed


def report() -> dict:
    """给自检用：当前 docx 模块里还有几个是相对路径、templates 定位得到吗。"""
    info = {"program_dir": _program_dir(), "relative": [], "templates_ok": None}
    for name, mod in list(sys.modules.items()):
        if not name or not name.startswith("docx"):
            continue
        cur = getattr(mod, "__file__", None)
        if cur and not os.path.isabs(cur):
            info["relative"].append((name, cur))
    try:
        from docx.parts.hdrftr import HeaderPart
        info["templates_ok"] = len(HeaderPart._default_header_xml())
    except Exception as e:      # noqa: BLE001
        info["templates_ok"] = "error: %r" % (e,)
    return info
