# -*- coding: utf-8 -*-
"""回归锁：打包后 python-docx 必须还能读到自带模板。

被锁住的那个 bug（2.3.122，只在 macOS / Linux 上犯）
----------------------------------------------------
python-docx 拼模板路径用的是 ``dirname(__file__) + "/../templates/xxx.xml"``。
打包后 ``docx/parts/`` 不是真实目录（代码进了二进制，只有 templates 作为数据拷出来），
于是：

* POSIX：解析路径要求中间每一级都真实存在 → ``docx/parts`` 不存在 → ENOENT；
* Windows：Win32 先把 ``..`` 词法归一化 → 恰好没事。

表现是「页码修正静默失效」：fix 返回 ok、报告说改好了，文档里连页眉都没创建。

这里的测试不依赖具体平台：先构造一个"模板在旁边、但 ``__file__`` 指向的中间目录
不存在"的 dist 布局，然后断言 ``docx_boot.apply()`` 之后模板读得到、页眉真的写得进去。
"""

from __future__ import annotations

import os
import shutil
import sys
import zipfile

import pytest

from src import docx_boot


def _installed_docx_dir() -> str:
    import docx

    return os.path.dirname(os.path.abspath(docx.__file__))


@pytest.fixture()
def fake_dist(tmp_path):
    """造一个仿打包布局：dist/docx/templates/*.xml 存在，dist/docx/parts/ 不存在。"""
    src_tpl = os.path.join(_installed_docx_dir(), "templates")
    dst_tpl = tmp_path / "dist" / "docx" / "templates"
    dst_tpl.mkdir(parents=True)
    for name in ("default-header.xml", "default-footer.xml",
                 "default-settings.xml", "default-styles.xml",
                 "default-comments.xml", "default.docx"):
        s = os.path.join(src_tpl, name)
        if os.path.isfile(s):
            shutil.copyfile(s, dst_tpl / name)

    assert not (tmp_path / "dist" / "docx" / "parts").exists(), "parts 目录必须不存在"

    # 把 docx 各模块的 __file__ 指向"不存在的中间目录"，模拟打包后的形态
    import docx.parts.hdrftr as hdrftr
    import docx.parts.settings as settings_mod
    import docx.parts.styles as styles_mod

    mods = [hdrftr, settings_mod, styles_mod]
    try:
        import docx.parts.comments as comments_mod
    except Exception:      # 老版本 python-docx 没有 comments
        comments_mod = None
    if comments_mod is not None:
        mods.append(comments_mod)

    saved = {m: m.__file__ for m in mods}
    for m in mods:
        m.__file__ = str(tmp_path / "dist" / "docx" / "parts" / (m.__name__.split(".")[-1] + ".py"))

    # 还原：__file__ / 被接管的方法 / 幂等标志
    saved_attrs = {}
    for mod_name, cls_name, meth in (
        ("docx.parts.hdrftr", "HeaderPart", "_default_header_xml"),
        ("docx.parts.hdrftr", "FooterPart", "_default_footer_xml"),
        ("docx.parts.settings", "SettingsPart", "_default_settings_xml"),
        ("docx.parts.styles", "StylesPart", "_default_styles_xml"),
        ("docx.parts.comments", "CommentsPart", "_default_comments_xml"),
    ):
        try:
            mod = __import__(mod_name, fromlist=["_"])
            cls = getattr(mod, cls_name, None)
            if cls is not None and hasattr(cls, meth):
                saved_attrs[(cls, meth)] = cls.__dict__.get(meth)
        except Exception:
            continue

    yield str(tmp_path / "dist")

    for m, f in saved.items():
        m.__file__ = f
    for (cls, meth), orig in saved_attrs.items():
        if orig is None:
            try:
                delattr(cls, meth)
            except Exception:
                pass
        else:
            setattr(cls, meth, orig)
    docx_boot._applied = False


def test_header_template_unreadable_then_fixed(fake_dist):
    """核心断言：接管之后，页眉模板必须读得到（不论 __file__ 长什么样）。"""
    from docx.parts.hdrftr import HeaderPart

    # 修前：中间目录不存在，POSIX 上必然读失败。Windows 会归一化，所以只有非 Windows 才断言。
    if sys.platform != "win32":
        with pytest.raises(FileNotFoundError):
            HeaderPart._default_header_xml()

    n = docx_boot.apply(base_dir=fake_dist)
    assert n >= 1, "应该至少接管 1 个读模板的方法：%r" % (docx_boot.report(),)

    data = HeaderPart._default_header_xml()
    assert isinstance(data, bytes) and len(data) > 500
    assert b"hdr" in data, "读到的内容不像页眉模板"

    rep = docx_boot.report()
    assert rep["templates_dir"].endswith(os.path.join("docx", "templates"))
    assert rep["templates_ok"] == len(data)


def test_page_number_can_be_written_after_boot(fake_dist):
    """端到端：接管之后页码真的能写进 header part。"""
    import docx

    docx_boot.apply(base_dir=fake_dist)

    from src.engine.fixer import _apply_page_number

    doc = docx.Document()
    doc.add_paragraph("Body text for the page-number test.")
    stats = _apply_page_number(doc, "top-right")
    assert stats.get("added") == 1
    assert stats.get("location") == "header"
    assert stats.get("alignment") == "right"

    out = os.path.join(os.path.dirname(fake_dist), "pn.docx")
    doc.save(out)
    with zipfile.ZipFile(out) as z:
        names = z.namelist()
    assert any(n == "word/header1.xml" for n in names), "页眉 part 没被创建：%s" % names


def test_boot_is_noop_when_templates_missing(tmp_path):
    """找不到模板目录时绝不硬塞路径，只留证据（否则会制造更难查的失败）。"""
    empty = tmp_path / "empty_dist"
    empty.mkdir()
    assert docx_boot.apply(base_dir=str(empty)) == 0
    rep = docx_boot.report()
    assert rep["templates_dir"] == ""
    assert rep["reasons"], "必须说明为什么没接管"
