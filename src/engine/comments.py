"""Word 批注（comments.xml）模块 —— 海外版「一键修正」的改动说明层。

Thesis Format Doctor Global —— 用 python-docx 的 Part 模型在文档里写入批注，
让用户知道「哪段被改了、按什么规则改的、该怎么手动调」。

设计铁律（移植自国内版踩坑记录）：
- 每遍修正开头先 `strip_our_comments()` 清掉本工具上一遍写的批注（按作者标记判定），
  避免「旧气泡带下来」显得脏（国内版 Pit 4）。这样无论跑多少遍，批注始终是「当前这一遍」的。
- 批注作者固定为 COMMENT_AUTHOR，绝不改写用户 / 模板原有批注。
- 批注只加在「被改动的段」与「标题段」上；正文逐段不逐一加（太吵），改为文档首段一条
  汇总批注说明本遍全部页面 / 正文级改动。
- 全部英文（海外用户）。

本模块只依赖 python-docx 的底层 Part 模型，不依赖任何网络 / LLM。
"""

from __future__ import annotations

import datetime
from typing import Optional

from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.packuri import PackURI
from docx.opc.part import Part
from lxml import etree

# 本工具批注的统一作者标记；strip_our_comments 据此识别并清除本工具产物。
COMMENT_AUTHOR = "Thesis Format Doctor"

_W_COMMENTS_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_COMMENT_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"
_COMMENT_PARTNAME = "/word/comments.xml"


def _get_comments_part(document):
    """返回 (part, root_element)。若文档尚无 comments 部件则新建并注册。

    新建时同时注册 [Content_Types].xml 的 Override 与 document.xml.rels 的关系，
    保证 Word 能识别 comments.xml。
    """
    main = document.part
    part = None
    try:
        part = main.part_related_by(RT.COMMENTS)
    except Exception:
        part = None
    if part is None:
        blob = (
            b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            b'<w:comments xmlns:w="' + _W_COMMENTS_NS.encode("utf-8") + b'"/>'
        )
        part = Part(PackURI(_COMMENT_PARTNAME), _COMMENT_CT, blob, main.package)
        main.relate_to(part, RT.COMMENTS)
        try:
            main.package.content_types.add_override(_COMMENT_PARTNAME, _COMMENT_CT)
        except Exception:
            # 已存在 Override 或 content_types 模型略有差异时忽略，不阻断主流程
            pass
    root = parse_xml(part.blob)
    return part, root


def _reserialize(part, root) -> None:
    """把改动后的 comments 根元素写回 part 的 blob（plain Part 用 blob 序列化）。"""
    part._blob = etree.tostring(root, encoding="UTF-8", xml_declaration=True, standalone=True)


def add_comment(document, para, text: str) -> int:
    """在 para 旁加一条批注，返回批注 id。

    - 在 comments.xml 追加一个 w:comment（作者=COMMENT_AUTHOR）；
    - 在段落首尾插入 w:commentRangeStart / w:commentRangeEnd，并在段尾追加
      w:commentReference run。
    """
    part, root = _get_comments_part(document)
    ids = [int(c.get(qn("w:id"))) for c in root.findall(qn("w:comment"))]
    nid = (max(ids) + 1) if ids else 0

    c = OxmlElement("w:comment")
    c.set(qn("w:id"), str(nid))
    c.set(qn("w:author"), COMMENT_AUTHOR)
    c.set(qn("w:date"), datetime.datetime.now(datetime.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"))
    c.set(qn("w:initials"), "TFD")
    cp = OxmlElement("w:p")
    cr = OxmlElement("w:r")
    ct = OxmlElement("w:t")
    ct.set(qn("xml:space"), "preserve")
    ct.text = text
    cr.append(ct)
    cp.append(cr)
    c.append(cp)
    root.append(c)
    _reserialize(part, root)

    # 段落标记
    p_el = para._p
    cs = OxmlElement("w:commentRangeStart")
    cs.set(qn("w:id"), str(nid))
    ce = OxmlElement("w:commentRangeEnd")
    ce.set(qn("w:id"), str(nid))
    cref_r = OxmlElement("w:r")
    cref = OxmlElement("w:commentReference")
    cref.set(qn("w:id"), str(nid))
    cref_r.append(cref)

    pPr = p_el.find(qn("w:pPr"))
    if pPr is not None:
        pPr.addnext(cs)        # w:pPr 之后第一个位置
    else:
        p_el.insert(0, cs)
    p_el.append(ce)
    p_el.append(cref_r)
    return nid


def strip_our_comments(document) -> int:
    """清掉本工具上一遍写的批注（按作者标记），返回清除的批注数。

    同时移除 document.xml 里对应的 commentRangeStart / commentRangeEnd /
    commentReference 标记，避免残留孤立标记。幂等：无本工具批注时返回 0。
    """
    main = document.part
    part = None
    try:
        part = main.part_related_by(RT.COMMENTS)
    except Exception:
        return 0
    if part is None:
        return 0

    root = parse_xml(part.blob)
    our_ids = {
        c.get(qn("w:id"))
        for c in root.findall(qn("w:comment"))
        if c.get(qn("w:author")) == COMMENT_AUTHOR
    }
    if not our_ids:
        return 0

    for c in list(root):
        if c.get(qn("w:id")) in our_ids:
            root.remove(c)
    _reserialize(part, root)

    # 移除 document.xml 中的对应标记
    doc_root = main.element
    to_remove = []
    for el in doc_root.iter():
        tag = el.tag
        if tag in (qn("w:commentRangeStart"), qn("w:commentRangeEnd")):
            if el.get(qn("w:id")) in our_ids:
                to_remove.append(el)
        elif tag == qn("w:r"):
            cref = el.find(qn("w:commentReference"))
            if cref is not None and cref.get(qn("w:id")) in our_ids:
                to_remove.append(el)
    for el in to_remove:
        parent = el.getparent()
        if parent is not None:
            parent.remove(el)
    return len(our_ids)
