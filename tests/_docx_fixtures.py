# -*- coding: utf-8 -*-
"""测试用 docx 夹具：批注锚点 / 批注注入 / 满格表。

为什么单独一个模块：海外「模板批注」相关的三个测试文件都要造「带 Word 批注的
docx」，而 python-docx 不原生支持写批注（要手工拼 comments.xml + Content-Type
Override + relationship）。三份复制粘贴必然漂移，故收在这里。

⚠️ ``inject_comments`` 的 XML 闭合标签是 ``</Types>`` / ``</Relationships>``
（**不带冒号**）。照抄国内旧脚本的 ``</Types:`` 是空操作——批注会写进 zip 却不
建立关系，Word 能打开但引擎按关系读不到（实测踩过）。
"""

from __future__ import annotations

import zipfile

from docx.oxml import OxmlElement
from docx.oxml.ns import qn


def mark_comment(paragraph, cid: int) -> None:
    """在段落上挂一条批注锚点（同一段可挂多条，Word 支持）。"""
    p = paragraph._p
    cs = OxmlElement("w:commentRangeStart")
    cs.set(qn("w:id"), str(cid))
    ppr = p.find(qn("w:pPr"))
    if ppr is not None:
        ppr.addnext(cs)
    else:
        p.insert(0, cs)
    ce = OxmlElement("w:commentRangeEnd")
    ce.set(qn("w:id"), str(cid))
    p.append(ce)
    r = OxmlElement("w:r")
    cref = OxmlElement("w:commentReference")
    cref.set(qn("w:id"), str(cid))
    r.append(cref)
    p.append(r)


def grid_table(doc, rows: int = 2, cols: int = 2):
    """插入满格表：四边 + 内部横竖线全有 —— 明确不是三线表。"""
    t = doc.add_table(rows=rows, cols=cols)
    for r in range(rows):
        for c in range(cols):
            t.cell(r, c).text = "x"
    borders = OxmlElement("w:tblBorders")
    for side in ("top", "bottom", "left", "right", "insideH", "insideV"):
        e = OxmlElement("w:" + side)
        e.set(qn("w:val"), "single")
        e.set(qn("w:sz"), "8")
        borders.append(e)
    t._tbl.tblPr.append(borders)
    return t


def inject_comments(src: str, out: str, comments) -> None:
    """把 ``[(id, text), ...]`` 写成 ``word/comments.xml`` 并建立关系。

    锚点必须**在保存前**用 ``mark_comment`` 挂好（写在这个函数里就晚了）。
    """
    body = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">',
    ]
    for cid, text in comments:
        esc = (text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))
        body.append(
            '<w:comment w:id="%d" w:author="Test" w:initials="T" '
            'w:date="2026-10-29T10:00:00Z"><w:p><w:r><w:t xml:space="preserve">'
            "%s</w:t></w:r></w:p></w:comment>" % (cid, esc))
    body.append("</w:comments>")

    with zipfile.ZipFile(src) as zin:
        items = {n: zin.read(n) for n in zin.namelist()}

    ct = items["[Content_Types].xml"].decode("utf-8")
    if "/word/comments.xml" not in ct:
        ct = ct.replace(
            "</Types>",
            '<Override PartName="/word/comments.xml" ContentType="application/'
            'vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/></Types>')
    items["[Content_Types].xml"] = ct.encode("utf-8")

    rels = items["word/_rels/document.xml.rels"].decode("utf-8")
    if "comments.xml" not in rels:
        rels = rels.replace(
            "</Relationships>",
            '<Relationship Id="rIdCmts900" Type="http://schemas.openxmlformats.org/'
            'officeDocument/2006/relationships/comments" Target="comments.xml"/></Relationships>')
    items["word/_rels/document.xml.rels"] = rels.encode("utf-8")
    items["word/comments.xml"] = "".join(body).encode("utf-8")

    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for n, data in items.items():
            zout.writestr(n, data)
