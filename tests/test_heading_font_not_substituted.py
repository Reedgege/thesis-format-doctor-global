"""标题 / 题名页的字体绝不被「正文兜底」替换（v2.3.123）。

为什么需要这份护栏
------------------
macOS 2.3.122 产物在真实 APA 学生样例上实测：客户原稿全文 Arial 10pt，点「修正」后
**题名页 6 段（题目/姓名/系/课程/导师/日期）与 Abstract、Introduction 标题被刷成
Times New Roman，而正文 15 段仍是 Arial** —— 同一份文档两种字体，客户拿到的稿子比改
之前更乱。

根因（`fixer._format_heading_one`）：``head_font = f.font_family or body_font``。
内建规范（APA / IEEE / MLA）只声明**正文**字体，对标题字体一个字没提
（``TitlePageFormat`` 里连 title_font 字段都没有），于是「规范正文字体」被当成
标题字体真的写进了 run；而正文按裁定 C-4（§2.19 字体↔字号配对不符只提示、不自动修）
一个字节都不动，两边就这么岔开了。

这条与国内 Windows 版（学生版 / 导师版 ``headings_fix._set_run_rpr``）的口径**相反**：
那边是「spec 含该字段才写、不含则不写（交还样式管辖）」，模板没给这一级字体就绝不
替客户编一个。现在海外版对齐成同一口径：``f.font_family`` 是唯一字体来源。

顺带修掉一个「检查不报、修正却改」的问题：检查器根本不判标题字体（heading format
只看居中/加粗/行距），修正器却偷偷换字体，违反 check→fix 单一真相源。
"""
from __future__ import annotations

import os

import pytest
from docx import Document
from docx.shared import Pt

from src.engine.fixer import fix_docx
from src.engine.questionnaire import target_from_dict, target_from_spec

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
APA_STUDENT_DOCX = os.path.join(ROOT, "tools", "crossplat", "fixtures",
                                "paper_apa_student.docx")


def _fonts_of(para) -> set:
    """段落里所有非空 run 的显式 ascii 字体（没写 rFonts 就是「交给样式」）。"""
    out = set()
    for r in para.runs:
        if r.text.strip():
            out.add(r.font.name or "(inherit)")
    return out


def _doc_with_headings(tmp_path, font="Arial"):
    """两个带 Heading 样式的标题 + 两段正文，标题 run 显式带 Arial。

    必须套 Heading 样式：层级识别靠样式，纯文本段会被当成「伪标题」推断层级，
    那种情况下 heading_formats 未必落到这两段上，测不到我们关心的事。
    """
    d = Document()
    for lvl, text in ((1, "Heading One"), (2, "Sub Heading Two")):
        p = d.add_paragraph()
        p.style = d.styles["Heading %d" % lvl]
        run = p.add_run(text)
        run.font.name = font
        run.font.size = Pt(11)
    # 正文 run 同样显式带字体：真实客户稿都带（Word 默认就会写 rFonts），
    # 而且「run 没写字体」是另一条既有逻辑（补规范字体），与本案「替换已有字体」无关。
    for text in ("Body paragraph under heading one.", "Another body paragraph."):
        p = d.add_paragraph()
        run = p.add_run(text)
        run.font.name = font
        run.font.size = Pt(11)
    path = tmp_path / "in.docx"
    d.save(str(path))
    return str(path)


# ------------------------------------------------------------------ 标题
def test_heading_font_untouched_when_nothing_specified(tmp_path):
    """规范只说「居中加粗」、没提字体 → 标题字体一个字节都不许动。

    旧实现会把它刷成规范正文字体（APA = Times New Roman），与正文混排。
    """
    src = _doc_with_headings(tmp_path)
    out = tmp_path / "out.docx"
    target = target_from_dict({
        "heading_levels": 2,
        "heading_formats": {
            "1": {"alignment": "center", "bold": True},
            "2": {"alignment": "left", "bold": True},
        },
    })
    fix_docx(str(src), str(out), target)
    got = [p for p in Document(str(out)).paragraphs if p.text.strip()]
    for p in got:
        assert "Times New Roman" not in _fonts_of(p), \
            "没规定字体却被换成了正文字体：%r -> %r" % (p.text, _fonts_of(p))


def test_heading_font_still_applied_when_explicitly_given(tmp_path):
    """显式给了本级字体（学校模板「一级黑体」）→ 必须真的写进去，不许因噎废食。"""
    src = _doc_with_headings(tmp_path)
    out = tmp_path / "out.docx"
    target = target_from_dict({
        "heading_levels": 2,
        "heading_formats": {
            "1": {"alignment": "center", "bold": True, "font_family": "SimHei"},
        },
    })
    fix_docx(str(src), str(out), target)
    h1 = [p for p in Document(str(out)).paragraphs
          if p.text.strip() == "Heading One"][0]
    assert h1.runs[0].font.name == "SimHei"


def test_heading_font_from_apa_spec_is_not_substituted(tmp_path):
    """内建 APA 规范下同样不许替换（这是线上真实踩到的路径）。"""
    src = _doc_with_headings(tmp_path)
    out = tmp_path / "out.docx"
    fix_docx(str(src), str(out), target_from_spec("APA-Student"))
    for p in [p for p in Document(str(out)).paragraphs if p.text.strip()]:
        assert "Times New Roman" not in _fonts_of(p)


# ------------------------------------------------------------- 题名页
@pytest.mark.skipif(not os.path.isfile(APA_STUDENT_DOCX),
                    reason="缺 crossplat 的 APA 学生样例（tools/crossplat/fixtures）")
def test_title_page_font_survives_body_font_fallback(tmp_path):
    """真实样例：全文 Arial，修正后题名页不许变成 Times New Roman。

    这是 2.3.122 的原始现场 —— 题名页被刷 TNR、正文仍是 Arial。
    """
    out = tmp_path / "out.docx"
    fix_docx(APA_STUDENT_DOCX, str(out), target_from_spec("APA-Student"))

    src_doc = Document(APA_STUDENT_DOCX)
    out_doc = Document(str(out))
    src_fonts = set().union(*[_fonts_of(p) for p in src_doc.paragraphs if p.text.strip()])
    # 原稿是单一 Arial；修正后只允许出现「同一种字体」或「交还样式」，
    # 不许冒出第二种具体字体（那就是被替换了）。
    out_fonts = set().union(*[_fonts_of(p) for p in out_doc.paragraphs if p.text.strip()])
    out_fonts.discard("(inherit)")
    assert out_fonts <= src_fonts, \
        "修正后出现了原稿没有的字体：%r（原稿只有 %r）" % (sorted(out_fonts), sorted(src_fonts))


@pytest.mark.skipif(not os.path.isfile(APA_STUDENT_DOCX),
                    reason="缺 crossplat 的 APA 学生样例（tools/crossplat/fixtures）")
def test_title_page_still_centered_after_font_guard(tmp_path):
    """字体不兜底 ≠ 什么都不做：题名页的居中仍然必须生效（别把功能改死）。"""
    out = tmp_path / "out.docx"
    fix_docx(APA_STUDENT_DOCX, str(out), target_from_spec("APA-Student"))
    doc = Document(str(out))
    first = [p for p in doc.paragraphs if p.text.strip()][0]
    assert "CENTER" in str(first.alignment).upper()
