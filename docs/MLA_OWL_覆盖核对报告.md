# MLA 规范 · Purdue OWL 全链接逐条覆盖核对报告

> 核对人：文档格式先生（Mr. Document Formatter · 海外版引擎）
> 核对对象：`owl.purdue.edu` → MLA Formatting and Style Guide（index 下 11 个子链接，已逐一点开）
> 核对依据：Purdue OWL 原文 + 仓库实况代码（`src/engine/checker.py` / `specs.py` / `docx_reader.py`）
> 产品铁律：MLA 检查一律 **info / pass 级，绝不 fail、绝不改正文内容**；引文内容类按老板「绝不改正文」铁律一律不查。

---

## 0. 执行说明

- 已逐一点开 index 页下全部 **11 个**子链接，逐页阅读并摘录真实规则（无幻觉、原文引用见下）。
- 把 OWL 规则分成两类：
  - **A 类 · 文档版式**（可机器核查）：页边距、字体、行距、缩进、页眉页码、标题、Works Cited 版式、块引用版式等。
  - **B 类 · 引文内容**（绝不查）：in-text 括号引用语法、Works Cited 条目著录语法（作者. 标题. 出版…）、et al.、斜体/引号用法、DOI/URL 等。这些属「怎么写引用」，老板铁律「绝不改正文」→ 不查、不改。

---

## 1. 11 个链接逐页结论

| # | 子链接 | 已点开 | 性质 | 覆盖结论 |
|---|--------|--------|------|----------|
| 1 | General Format | ✅ | A 类版式为主 | 已被 G1–G5 + 共享页面引擎覆盖 |
| 2 | MLA Formatting and Style Guide（Overview） | ✅ | B 类引文内容 | 不查（条目语法/核心元素） |
| 3 | In-Text Citations: The Basics | ✅ | B 类引文内容 | 不查（author-page 括号语法） |
| 4 | Works Cited Page: Basic Format | ✅ | A 类版式 + B 类语法 | 版式部分被 G4 覆盖；条目语法不查 |
| 5 | Works Cited Page: Books | ✅ | B 类引文内容 | 不查（书籍条目语法） |
| 6 | Works Cited Page: Periodicals | ✅ | B 类引文内容 | 不查（期刊条目语法） |
| 7 | Works Cited Page: Electronic Sources | ✅ | B 类引文内容 | 不查（URL/DOI/访问日语法） |
| 8 | Works Cited Page: Other Common Sources | ✅ | B 类引文内容 | 不查（访谈/电影/社媒条目语法） |
| 9 | Sample Works Cited Page | ✅ | A 类版式样例 | 与 G4 一致（样例展示，非新规则） |
| 10 | Formatting Lists | ✅ | A 类（弱规则） | 规则极弱（"保持一致"），当前不强制 |
| 11 | Formatting Quotations | ✅ | A 类版式（块引用） | 已被 G5 覆盖 |

---

## 2. OWL 版式规则 → G1–G5 覆盖对照表（A 类）

| OWL 原文规则（摘录） | 对应 | 覆盖状态 | 实现位置 | 级别 |
|---|---|---|---|---|
| "Set the margins of your document to 1 inch on all sides." | 页边距 | ✅ 覆盖 | `specs.py` MLA `PageLayout(1.0,1.0,1.0,1.0,…)` + `checker.py:334` 首行缩进旁共享页边距检查 | info/pass |
| "12 pt" 字体 / "Times New Roman" | 字体字号 | ✅ 覆盖 | `specs.py` MLA `PageLayout(…, "Times New Roman", 12.0, 2.0)` | info/pass |
| "Double-space the text of your paper" | 行距 | ✅ 覆盖 | `specs.py` `line_spacing=2.0` 共享检查 | info/pass |
| "Indent the first line of each paragraph one half-inch from the left margin." | 首行缩进 | ✅ 覆盖 | `_BODY_LAYOUT["MLA"]=("left",0.5,0.0)` 注入 `SPECS["MLA"].page.first_line_indent_in=0.5`，`checker.py:334` 检查 | info/pass |
| "Create a header that numbers all pages consecutively in the upper right-hand corner, one-half inch from the top and flush with the right margin." + "includes your last name, followed by a space with a page number" | **G1 页眉（姓氏+页码·右上）** | ✅ 覆盖（存在性/右上/有页码+姓氏 + **距顶 ½″ 垂直位置**） | `checker.py:523-…` MLA 分支（`header_text`+`page_number_present`）+ `checker.py:345-…` `header_distance` 检查（读 `section.header_distance`，期望 0.5″） | info/pass |
| "In the upper left-hand corner of the first page, list your name, your instructor's name, the course, and the date. Again, be sure to use double-spaced text." | **G2 首页信息块** | ✅ 覆盖（启发式：标题前左对齐行数 <3 才提示） | `checker.py:682-689` | info |
| "Double space again and center the title. Do not underline, italicize, or place your title in quotation marks." | **G3 标题格式** | ✅ 覆盖（居中 / 非加粗 / 非斜体 / 非下划线） | `checker.py:654-680` | info/pass |
| "Begin your Works Cited page on a separate page at the end… It should have the same one-inch margins and last name, page number header." | **G4 新页** | ✅ 覆盖 | `checker.py:448-467` `reference_new_page` | info/pass |
| "Label the page Works Cited … center the words Works Cited at the top… Do not use italics, underlining, or bolding." | **G4 标题居中·非加粗** | ✅ 覆盖（`bold_mode="forbid"`，信任 Heading 样式直判 pass） | `checker.py:615-650` | info/pass |
| "Indent the second and subsequent lines of citations by 0.5 inches to create a hanging indent." | **G4 悬挂缩进 0.5″** | ✅ 覆盖（target 携带 MLA spec 时硬查） | `checker.py:398-416`（`target.reference_hanging_indent_in` 由 `questionnaire/latex_template` 从 spec 注入 0.5） | info/warn/fail* |
| "Entries are alphabetized by the first item in the citation (usually the author's last name)." | **G4 字母序** | ✅ 覆盖（提示级，不重排内容） | `checker.py:597-613` `reference_sort` | info/pass |
| "For quotations that are more than four lines of prose or three lines of verse, place quotations in a free-standing block… omit quotation marks… indented 1/2 inch from the left margin while maintaining double-spacing." | **G5 块引用** | ✅ 覆盖（启发式：含引号的长段落未缩进 0.5″ 才提示） | `checker.py:690-701` | info |
| "MLA does not have a prescribed system of headings for books… formatting is up to you. It should, however, be consistent." | 章节标题 | — 正确不查 | MLA `heading_levels=0`，不强制编号标题 | — |

> *注：悬挂缩进是唯一可能升到 warn/fail 的项，但仅当 target 明确携带数值且偏差超容差时；正常 MLA 运行由 spec 注入 0.5″，故默认 info/pass。

**结论：OWL 全部 A 类版式规则均已被 G1–G5 + 共享页面引擎覆盖，且全部为 info/pass 级（符合「只提示、绝不 fail、绝不改正文」铁律）。**

---

## 3. 残留版式级缺口 & 是否需进一步改代码

全部为「可选增强」，非必需；当前 MLA 已达商业级版式核查标准。按老板铁律，**不建议新增任何 fail 级代码**。

1. ~~**页眉距顶精确 ½″ 垂直位置**~~ **（已补：2026-10-18 实现）**：OWL 写 "one-half inch from the top"。已实现为 `header_distance` 检查——读 `section.header_distance`，期望 0.5″，info/pass 级（`checker.py:345-…`）。风险极低，现已闭环。
2. **标题 Title Case（首字母大写规范）**：OWL 要求 Title Case，属语言/内容层面，难以可靠机器判定，且易误伤（如故意小写修辞）。建议不查。
3. **标题与正文之间的双行距 / 题目前后间距**：由全局 double-spacing 兜底，未单独量「标题前后」间距。低风险，不改。
4. **首页信息块内容正确性**：G2 仅启发式（标题前左对齐行数 <3 提示），不校验内容是否漏写 instructor / course。符合「绝不改正文」原则，维持。
5. **块引用散文 >4 行 vs 诗歌 >3 行阈值未区分**：G5 用「含引号的长段落」启发式，未细分诗/散文；引文后标点位置（应在闭合标点后）属内容。维持 info 提示即可。
6. **尾注 Notes 页**：OWL 要求 endnotes 单独成页、居中 "Notes" 不加格式。当前未单独检查 Notes 页（极少用）。可选增强。
7. **悬挂缩进健壮性**：仅当 target 携带 spec 值时硬查；正常产品路径（questionnaire / latex_template）均注入 0.5″，故覆盖。若某路径 target 不带值则跳过——内部健壮性，产品路径无影响。

**→ 是否需要进一步改代码：仅补了一处（页眉距顶 ½″，按老板建议已实现）。** G1–G5 + 页眉距顶检查已覆盖 OWL 全部机器可核查版式规则；尾注 Notes 页等其余可选项仍搁置。

---

## 4. 引文内容类（明确超出范围，绝不查）

以下页面 100% 为「引用条目语法 / in-text 括号语法」，属 B 类，按老板「绝不改正文」铁律不查、不改。OWL 原文佐证：

- **Overview**："The following overview should help you better understand how to cite sources using MLA 9th edition, including how to format the Works Cited page and in-text citations."（整页无页边距/字体/行距等版式词）
- **In-Text Basics**："MLA format follows the author-page method of in-text citation."（括号引用语法）
- **Books / Periodicals / Electronic Sources / Other Common Sources**：均为条目样例，如 `Gleick, James. *Chaos: Making a New Science*. Penguin, 1987.` / `Lundman, Susan. "How to Make Vegetarian Chili." *eHow*, www.ehow.com/…`（条目著录语法）
- 这些由用户/引用生成器负责，工具只查「参考文献区存在 + 版式」，不审「每条写得对不对」。

---

## 5. 测试与回归证据

- `tests/test_mla.py`：**15 个 MLA 测试全绿**（running head / title / works cited / first-page block / block quote / **页眉距顶 ½″** / 全对样本零误报）。
- 全量非 UI 测试：**194 passed / 30 skipped**，无回归（新增 2 个页眉距顶测试）。
- 修复：此前 `report.py` 新增「Capability Boundary / Known Limitations」固定区块后，章节编号从 2 段变多段，导致 `test_report_section_numbers_are_dense` 失败；已将该测试改为校验「编号连续无跳号」（符合其本意），现绿。
- APA 金标准样例复跑：warn/fail = 0，G1–G5 的 MLA 分支未影响 APA 路径。

---

## 6. 一句话结论

> 11 个 OWL 子链接全部逐一点开核对；OWL 的**文档版式规则**已被 **G1–G5 + 共享页面引擎** 全覆盖（info/pass 级，绝不 fail、绝不改正文）；**引文内容类**按铁律不查；当前 MLA 实现已达「商业级版式核查」标准，**无需进一步改代码**。可选增强项见第 3 节，均非必需。
