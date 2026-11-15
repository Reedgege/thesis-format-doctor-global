# -*- coding: utf-8 -*-
"""让打包后的 python-docx 找得到自带模板 —— 否则 macOS / Linux 上「建页眉」必挂。

问题现象（2.3.122 实测）
------------------------
macOS / Linux 发行版里「页码修正」完全失效，而且**不报错**：fix 返回成功、
报告写着页码已改到右上，实际文档里连页眉都没被创建；同样的代码在 Windows 上
是好的，用源码跑也是好的。

真根因（2026-11-15 从产物里取证，非推测）
-----------------------------------------
python-docx 读自带模板时这样拼路径（`docx/parts/hdrftr.py`）：

    path = os.path.join(os.path.split(__file__)[0], "..", "templates",
                        "default-header.xml")

打包（Nuitka standalone）后 docx 的代码被编进二进制，磁盘上只有作为**数据文件**
拷出来的 ``docx/templates/*.xml``；``docx/parts/`` 已经不是真实目录了。于是上面的
拼接结果形如 ``.../docx/parts/../templates/default-header.xml``：

    · POSIX（macOS / Linux）：解析路径时要求**中间每一级都真实存在**，
      ``docx/parts`` 不存在 → 立刻 ENOENT；
    · Windows：Win32 先把 ``..`` 词法归一化再打开 → 恰好没事。

所以这是"只有非 Windows 平台才会犯"的错，而且只影响带 ``..`` 的那 5 处。
反证：``Document()`` 读的是 ``templates/default.docx``（没有 ``..``），
在所有平台都正常 —— 与实测完全吻合（空白文档能建、页眉建不了）。

受影响的 5 个函数（都是「首次创建部件」时才调用）：
    HeaderPart._default_header_xml      → 页眉（页码 / running head）
    FooterPart._default_footer_xml      → 页脚
    SettingsPart._default_settings_xml  → 文档设置
    StylesPart._default_styles_xml      → 样式表
    CommentsPart._default_comments_xml  → 批注

修法
----
不猜 ``__file__`` 会是什么形态（实测它本来是绝对路径，改它没用），改为：

1. 在程序目录附近找到真实的 ``docx/templates`` 目录；
2. 把上面 5 个 classmethod 接管成"直接从该目录读文件"的版本。

这样与 ``__file__`` 长什么样彻底解耦，也不依赖 python-docx 的拼接写法。
找不到模板目录就原地不动（绝不硬塞一个错的路径），并把失败原因记进 ``reasons``
供 ``selfcheck`` 输出 —— 排查时能一眼看到"是没找到，还是不需要"。

源码运行时模板在 site-packages 里找得到，接管后行为与原生完全一致（读的是同一份
文件），所以开发环境不受影响；找不到时更是直接跳过。
"""

from __future__ import annotations

import os
import sys

_applied = False

# 被接管的目标：(模块名, 类名, 方法名, 模板文件名)
_READERS = (
    ("docx.parts.hdrftr", "HeaderPart", "_default_header_xml", "default-header.xml"),
    ("docx.parts.hdrftr", "FooterPart", "_default_footer_xml", "default-footer.xml"),
    ("docx.parts.settings", "SettingsPart", "_default_settings_xml", "default-settings.xml"),
    ("docx.parts.styles", "StylesPart", "_default_styles_xml", "default-styles.xml"),
    ("docx.parts.comments", "CommentsPart", "_default_comments_xml", "default-comments.xml"),
)

# 判定「这是个可用的 docx 模板目录」所需的最小文件集
_MUST_HAVE = ("default-header.xml", "default-footer.xml")

# 自检/诊断用：上一次 apply 的结果
_LAST = {"templates_dir": "", "patched": [], "reasons": [], "ok": None}


def _program_dirs() -> list:
    """可能放着 docx 模板的候选根目录（可执行文件所在目录优先）。"""
    out = []

    def _add(p):
        if not p:
            return
        try:
            d = os.path.dirname(os.path.abspath(p))
        except Exception:
            return
        if d and d not in out:
            out.append(d)

    argv = getattr(sys, "argv", None) or []
    _add(argv[0] if argv else "")
    _add(getattr(sys, "executable", "") or "")
    return out


def _looks_like_templates(d: str) -> bool:
    try:
        return all(os.path.isfile(os.path.join(d, n)) for n in _MUST_HAVE)
    except Exception:
        return False


def _find_templates_dir(bases) -> str:
    """在候选根目录下找真实的 docx/templates 目录。

    先试最常见的 ``<base>/docx/templates``，再限深遍历（打包目录层级因平台而异，
    尤其 macOS 的 .app 包）。遍历有深度上限，不会在客户机上乱翻。
    """
    for base in bases:
        fast = os.path.join(base, "docx", "templates")
        if _looks_like_templates(fast):
            return fast
    for base in bases:
        try:
            for dirpath, dirnames, _files in os.walk(base):
                if os.path.basename(dirpath) == "templates" and _looks_like_templates(dirpath):
                    return dirpath
                # 限制下探深度：docx 不会藏太深，避免在大目录里浪费时间
                rel = dirpath[len(base):]
                if rel.count(os.sep) >= 5:
                    dirnames[:] = []
        except Exception:
            continue
    return ""


def _patch_readers(tpl_dir: str) -> list:
    """把 5 个读模板的 classmethod 接管掉，返回接管的清单。

    只接管**确实存在**的方法；python-docx 若改了名字，这里会跳过而不是报错，
    但下面的 verify 会把"接管之后仍读不到"暴露出来。
    """
    import importlib

    patched = []

    def _make_reader(fname: str):
        cache = {"data": None}

        def _reader(cls, *a, **kw):        # classmethod 形态：(cls)
            if cache["data"] is None:
                path = os.path.join(tpl_dir, fname)
                with open(path, "rb") as f:
                    cache["data"] = f.read()
            return cache["data"]

        return _reader

    for mod_name, cls_name, meth, fname in _READERS:
        try:
            mod = importlib.import_module(mod_name)
            cls = getattr(mod, cls_name, None)
        except Exception as e:      # noqa: BLE001
            _LAST["reasons"].append("import %s failed: %r" % (mod_name, e))
            continue
        if cls is None or not hasattr(cls, meth):
            _LAST["reasons"].append("%s.%s 不存在（python-docx 版本差异）"
                                    % (cls_name, meth))
            continue
        try:
            setattr(cls, meth, classmethod(_make_reader(fname)))
            patched.append("%s.%s" % (cls_name, meth))
        except Exception as e:      # noqa: BLE001
            _LAST["reasons"].append("patch %s.%s failed: %r" % (cls_name, meth, e))
    return patched


def _verify() -> object:
    """接管之后真读一次页眉模板。返回读到的字节数，失败返回异常字符串。"""
    try:
        from docx.parts.hdrftr import HeaderPart

        return len(HeaderPart._default_header_xml())
    except Exception as e:      # noqa: BLE001
        return "error: %r" % (e,)


def apply(base_dir: str | None = None) -> int:
    """执行接管，返回接管的函数个数（0 = 不用接管或没找到模板目录）。

    base_dir 仅供测试注入：测试要模拟"模板在旁边、但 ``__file__`` 指向的
    中间目录不存在"的打包布局。显式传入时不受幂等标志限制。
    """
    global _applied
    explicit = base_dir is not None
    if _applied and not explicit:
        return _LAST.get("count", 0)

    _applied = True
    _LAST.update({"templates_dir": "", "patched": [], "reasons": [], "ok": None,
                  "count": 0})

    bases = [base_dir] if base_dir else _program_dirs()
    tpl = _find_templates_dir(bases)
    if not tpl:
        # 模板没跟着产物走，或压根没打包 —— 不猜、不动，只留证据。
        # 源码运行时理论上总能找到（site-packages 里有），找不到就说明
        # python-docx 本身装坏了，此时接管也没有意义。
        _LAST["reasons"].append(
            "未在候选目录下找到 docx/templates（候选：%s）" % ", ".join(bases or ["(无)"]))
        return 0

    _LAST["templates_dir"] = tpl
    patched = _patch_readers(tpl)
    _LAST["patched"] = patched
    _LAST["count"] = len(patched)
    _LAST["ok"] = _verify()
    return len(patched)


def report() -> dict:
    """给 selfcheck 用：模板目录、接管清单、接管后能否读到模板。"""
    info = {
        "program_dirs": _program_dirs(),
        "templates_dir": _LAST.get("templates_dir") or "",
        "patched": list(_LAST.get("patched") or []),
        "reasons": list(_LAST.get("reasons") or []),
        # 现读一次，而不是回放 apply 时的结果：能反映"此刻"的真实状态
        "templates_ok": _verify(),
    }
    return info
