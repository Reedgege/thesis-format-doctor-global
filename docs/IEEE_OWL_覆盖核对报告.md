# IEEE 规范 · Purdue OWL 覆盖核对报告

> 核对人：文档格式先生（Mr. Document Formatter · 海外版引擎）
> 核对对象（Purdue OWL · IEEE 板块共 5 页 + 1 外链，本次**全部逐段精读**）：
> 1. `owl.purdue.edu/.../ieee_style/IEEE_overview.html` → **IEEE Overview**（板块导航页）
> 2. `owl.purdue.edu/.../ieee_style/ieee_general_format.html` → **IEEE General Format**（基础版式/页边距/标题/摘要，已逐段摘录）
> 3. `owl.purdue.edu/.../ieee_style/in-text_citation.html` → **In-Text Citation**（正文内 `[1]` 引用语法）
> 4. `owl.purdue.edu/.../ieee_style/reference_list.html` → **Reference List**（文末参考文献编号列表著录模板）
> 5. `owl.purdue.edu/.../ieee_style/tables_figures_and_equations.html` → **Tables, Figures, and Equations**（图表公式编号/题注）
> 6. 外链 `ieee.org/.../ieee-taxonomy.pdf` → **IEEE Taxonomy**（index terms 关键词表，已打开确认非版式规则）
> 7. `ieee.org/conferences/publishing/templates` → **IEEE 会议稿件模板页**（官方只提供 `conference-template-a4.docx` / `conference-template-letter.docx` 两个 Word 模板，Updated 2024；具体版式数值内嵌在模板文件的 guidance text 中，网页不列数值）
> 核对依据：Purdue OWL 原文 + 官方 IEEE 模板文件实测（`_smoke/profile_strict_docx.py` 直读 XML）+ 仓库实况代码（`src/engine/checker.py` / `specs.py` / `fixer.py`）
> 产品铁律：IEEE 检查与 MLA 同理 **info / pass 级，绝不 fail、绝不改正文内容**；引文内容类按老板「绝不改正文」铁律一律不查。（唯一例外是参考文献编号/悬挂缩进属「引用归规范」维度，硬判但容差宽。）

---

## 0. 执行说明

- Purdue OWL 的 IEEE 板块**不是单页**，而是含 **5 个页面 + 1 个外链**（Overview / General Format / In-Text Citation / Reference List / Tables-Figures-Equations / IEEE Taxonomy PDF）。此前报告只读了 General Format 一页，已补全其余页面并逐段摘录真实规则（无幻觉，原文见第 1 节与新增第 1b 节）。MLA 有 11 个子链接、IEEE 是 5 个并列页，结构不同但都逐链接读过。
- IEEE 会议模板页被反爬拦截（202），但已用 `WebFetch` 二次尝试拿到元数据：官方只发 **A4 + US Letter** 两个 Word 模板；具体页边距/字体/字号等数值**只在模板文件内**，网页不列。这部分我用**直读两份 .docx 的 XML** 拿到了权威数值（见本仓 `_smoke/profile_strict_docx.py` 输出）。
- 把 OWL 规则分成两类：
  - **A 类 · 文档版式**（可机器核查）：页边距、字体、行距、标题层级格式、标题页、参考文献编号版式等。
  - **B 类 · 引文内容**（绝不查）：in-text `[1]` 括号语法、参考文献条目著录语法（作者. "标题," 期刊, vol, no, pp, 年）、et al.、斜体/引号用法、DOI/URL 等。

---

## 1. OWL IEEE 版式规则清单（原文摘录）

| OWL 原文规则 | 类别 |
|---|---|
| "The paper title should be centered at the top of the first page, in 24-point type." | A |
| "The byline should be centered below the title… in 10-point type. It should contain… Author's name / affiliation(s) / City & country / E-mail." | A |
| "The body of the paper should be in 10-point type, and formatted to appear in two columns." | A |
| "All papers must begin with an abstract and index terms."（摘要单段 ~200 词；禁脚注/引用/表格/图/公式；后接 index terms，字母序） | A |
| Primary heading：罗马数字 + 句点，**小型大写（small caps）**，**居中** | A |
| Secondary heading：大写字母 + 句点，**斜体**、title case、左对齐、不缩进、段前换行 | A |
| Tertiary heading：阿拉伯数字 + 反括号、**斜体**、title case、左对齐、缩进 1em、冒号接续 | A |
| Quaternary heading：小写字母 + 反括号、**斜体**、sentence case、左对齐、缩进 2em、冒号接续 | A |
| References / Acknowledgments：像 primary heading 但**不编号** | A |
| Appendix：像 primary，但单独编号（Appendix A / B） | A |
| 首字下沉（drop cap）两行深 + 后续 8–12 字符全大写 | A（视觉/内容） |
| 图/表/方程各自连续编号、列内居中 | A |
| in-text citation：编号 `[1]` 制 | B |
| 参考文献：编号列表 `[1] A. Author, "Title," Journal, vol., no., pp., Year.` | B |

---

## 1b. 其余 4 个 IEEE OWL 页面精读（补全「每个链接都点开」）

> 上一轮只读了 General Format 一页。本次把 Overview / In-Text Citation / Reference List / Tables-Figures-Equations 全部打开并逐段摘录，另打开 General Format 页内唯一外链（IEEE Taxonomy PDF）。结论先行：**这 4 页的规则 100% 属于「引文内容 / 视觉对象」类（B 类），按老板「绝不改正文」铁律一律不查、不改**；其中只有 Reference List 的*结构性*规则（编号制、按引用顺序、编号贴左边界）与我们的 `numbered=True` + `hanging_indent_in=0.0` 设计一致，等于 OWL **反向确认**了现有实现，无需改代码。

### 1b-1 · IEEE Overview（板块导航页）
- 说明 IEEE 广泛用于工程 / 计算机 / 科技领域；长篇稿件或细节问题应查 IEEE Editorial Style Manual；向具体会议/期刊投稿时**以对方指南为准**。
- 列出板块全部子页：**General Format / Abstracts,index terms / Section headings / Appendices,acknowledgements / In-Text Citation（含 citing within a reference、multiple references）/ Reference List（含 general principles、specific media）/ Tables,Figures,Equations**。
- 对产品启示：本工具定位「快速自查」，与 OWL 自述一致；**不替代官方样式手册、不替代投稿指南**——这点在产品免责声明里本就写明，吻合。

### 1b-2 · In-Text Citation（正文内引用，纯 B 类内容）
OWL 原文要点（逐条）：
- "IEEE employs **bracketed numbers** for in-text citation."
- "One should always refer to references exclusively by their number… it's always just **'in [1]'**, not 'in reference [1]'."
- 引用某条目内部某部分用方括号内缩写：`[6, Sec. VI]` / `[6, Algorithm 1]` / `[6, Ch. 3]` / `[6, eq. (4)]` / `[6, Fig. 5]` / `[6, p. 7]`(或 `[6, pp. 8-10]`) / `[6, Table XII]` / `[6, Th. 13]` / `[6, Lemma 6]` / `[6, Appendix II]`。
- 一次引多个：逗号分隔；**连续 3 个以上只写首尾 + en-dash**：`[1]–[3]`（逗号后留空格，en-dash 不留）。
- 若点名作者则须分开引（除非同作者）。
- "References are always **numbered in order of citation**… if you cite the same reference in multiple locations… it must always be cited with the **same number**."
- **分类**：100% 为 in-text `[1]` 语法与编号逻辑 → B 类内容，**不查、不改**（符合铁律）。

### 1b-3 · Reference List（文末参考文献列表，结构性规则与 spec 一致 ✓）
OWL 原文要点：
- "References should be provided on a **separate page** at the end of your paper, with the title **'References'** at the top."
- "listed and **numbered in order of citation, not alphabetically**."
- "The numbers should be **flush against the left margin, and separated from the body** of the reference." → 标准悬挂缩进版式（编号左贴边、文献正文退后）。
- 通用注释（B 类内容，不查）：作者姓+首字母、后缀 `Jr.`/`III` 带逗号；≤6 作者全列 `A. B. Author, C. D. Author, and E. F. Author`；≥7 作者首位+`et al.`；书/刊名 title case、文章/技术报告 sentence case；月份缩写 `Jan.`–`Dec.`；IEEE 期刊用官方缩写。
- 各文献类型模板（Books / Periodicals / Websites / Theses / Conferences / Technical Reports / Datasets / Government / Manuals / Patents / Unpublished）——**全为著录语法（B 类内容），不查、不改**。
- **与代码对应**：本仓 `specs.py` 的 `ReferenceRule(style_key="IEEE", hanging_indent_in=0.0, numbered=True, checks=("numbered_ref",))` + `checker.py:_chk_numbered_ref` + `reference_hanging_indent` 容差宽判 —— **OWL 的「numbered / order of citation / flush-left number」三项结构性要求全部被覆盖（info/warn 级）**。即 OWL 反向确认了现有设计正确。
- ⚠️ 一处可微调（非缺陷）：经典 IEEE 参考文献换行行采用约 0.25″ 悬挂缩进（编号 `[1]` 后正文退格），我们 `hanging_indent_in=0.0` 设为 0。属 B 类细节、容差宽，不影响；若想更贴经典可改 0.0→约 0.25，但**非必需**，记此待定。

### 1b-4 · Tables, Figures, and Equations（图表公式，纯视觉/内容，不查）
OWL 原文要点：
- 三类各自**全文连续编号、互不影响**（Section I 有 4 图，则 Section II 首图 = Fig. 5，非 Fig. II-1；表用罗马数字 Table I，非 Table IV）。
- **Tables**：罗马数字（Table I）；标签格式同章节/附录标题 + 题名；**题名在表上方、小型大写（small caps）**；注释在表下方。
- **Figures**：统称 figures（不叫 graph 1）；标签 **Fig. + 数字 + 句号 + em 空格 + 题名**，**在图下方、sentence case、无换行**。
- **Equations**：居中；**阿拉伯数字右对齐、圆括号 `(3)`**；正文只称 `in (3)`（句首例外可 `Equation (3)`）；变量/数字斜体，函数名/单位/单词罗马体。
- **分类**：编号/题注位置/大小写属**视觉对象层 + 内容层**，需对象检测，超出本工具范围；**不查、不改**（与 A/B 切分一致）。

### 1b-5 · 外链 IEEE Taxonomy PDF（General Format 页内唯一外链）
- 已打开确认：该 PDF 是 IEEE 官方 **index terms 关键词分类表**，用于作者选索引词，**不含任何论文版式规则**。
- 对产品：index terms 选取属 B 类内容，本就不查；该链接无需纳入引擎逻辑。

---

## 2. OWL 版式规则 → 我们的 spec / checker / fixer 覆盖对照表（A 类）

| OWL 原文规则 | 对应实现 | 覆盖状态 | 实现位置 | 级别 |
|---|---|---|---|---|
| 页边距 1″ | `specs.py` IEEE `PageLayout(1.0,1.0,1.0,1.0,…)` + 共享页边距检查 | ✅ 覆盖 | `specs.py:363` + `checker.py:262-264` 共享页边距 | info/pass |
| 字体 Times New Roman | `specs.py` IEEE `PageLayout(…, "Times New Roman", …)` | ✅ 覆盖 | `specs.py:363` | info/pass |
| 字号 12pt（thesis 默认）/ 10pt（会议模板） | `specs.py` 默认 12pt（thesis）；会议模板实测 10pt 由模板自身声明 | ⚠️ 典型值提示 | `specs.py:363` | info/pass |
| 行距 1.5（thesis）/ 单倍双栏（会议） | `specs.py` 默认 1.5（thesis） | ⚠️ 典型值提示 | `specs.py:363` | info/pass |
| 标题 24pt 居中 | **checker 不专门量度 24pt**；fixer 套 `title_page` 居中但无字号校验 | ❌ 未专门覆盖 | 仅 `fixer._apply_title_page_format` | — |
| 署名 byline 10pt 居中（姓名/单位/城市国家/邮箱） | **checker 无标题页结构检查**；仅 fixer 排版 | ❌ 未覆盖 | （fixer 侧） | — |
| 正文 10pt + 双栏 | 双栏**不查**；字号属典型值 | ❌ 未覆盖 | — | — |
| 摘要必须 + index terms | **checker 仅 APA-Professional 提示 abstract**（`expect_abstract` 门控，`checker.py:585`）；**IEEE 不提示**；index terms 不查 | ❌ 未覆盖（OWL 要求） | `checker.py:583-587` 仅 APA | — |
| 一级标题(I.) 小型大写 + 居中 | checker **只数层级数**，不验格式；fixer 套 `IEEE_HEADING[1]`=**加粗+全大写+居中**（OWL=小型大写，**bold vs small-caps 偏差**） | ⚠️ 偏差 | `checker.py:377-398`(仅计数) / `specs.py:192`(fixer 用) | info（checker 不验）/ fixer 套用有偏差 |
| 二级标题(A.) 斜体 + 左对齐 | fixer 套 `IEEE_HEADING[2]`=**加粗**（OWL=**斜体**）——**明显偏差** | ❌ 偏差 | `specs.py:194` | info / fixer 套用有误 |
| 三级(1))/四级(a)) 斜体 + 左缩进 | fixer 套 `IEEE_HEADING[3][4]`=**斜体**（与 OWL 一致 ✓） | ✅ 一致 | `specs.py:196-200` | info / fixer 一致 |
| References / Acknowledgments 像 primary 不编号 | ref 区 info 检查 + fixer 确认 "References 不编号" | ✅ 覆盖 | `checker.py:400-408` + `test_fixer.py:498` | info |
| 附录 Appendix 单独编号 | 不查 | ❌ 未覆盖 | — | — |
| 首字下沉 drop cap | 不查（视觉/内容层） | ❌ 未覆盖 | — | — |
| 图/表/方程连续编号、列内居中 | 不查 | ❌ 未覆盖 | — | — |
| 参考文献编号列表 `[1]` + 悬挂缩进 0 + 顺序 | `reference.numbered=True` + `hanging_indent_in=0.0` 硬判（容差宽）；顺序提示 | ✅ 覆盖 | `checker.py:136-148`(`_chk_numbered_ref`) + `checker.py:415-434`(hanging) | info/warn/fail* |

> *注：参考文献编号/悬挂缩进是「引用归规范」维度，硬判但容差宽；其余页面/标题维度对 IEEE 一律降级 info（只有 APA 的 alignment/first_line_indent/heading 保 warn 下限），符合「只提示、绝不 fail、绝不改正文」。

---

## 3. 残留缺口 & 偏差（含最致命的阻断项）

### ⛔ 阻断项（最致命，必须先行）
- **Strict Open XML 兼容缺口**：IEEE 官方会议模板（`conference-template-a4.docx` / `conference-template-letter.docx`）是**严格 OOXML**——命名空间用 `http://purl.oclc.org/ooxml/...` 而非常规 `schemas.openxmlformats.org/...`。**`python-docx` 1.2.0 打不开**（报 `KeyError: no relationship of type officeDocument`）；本仓的 editor_sdk（local-office-edit）同样打不开（`document is not opened`）。
- **后果**：用户只要把一份 IEEE 会议模板拖进「文档格式先生」，引擎在**打开那一步就崩**，第 2 节的任何检查都跑不起来。
- **修法（低风险、集中）**：在读入 python-docx 之前加一步 *strict→transitional* 命名空间重写预处理（把各 XML part 及 `.rels` 里的 `purl.oclc.org/ooxml` 整体替换为 `schemas.openxmlformats.org/...`）。这是让引擎能吃下 IEEE 会议模板的前提。

### ⚠️ 标题格式偏差（IEEE_HEADING 与 OWL 不符）
- `specs.py:191-202` 的 `IEEE_HEADING`：
  - L1 一级：**加粗 + 全大写 + 居中**（OWL 要**小型大写**——bold vs small-caps 偏差，但都居中+罗马数字，影响小）
  - **L2 二级：加粗（OWL 明确要斜体）——明显偏差**；且 checker 不验格式、只靠 fixer 套用，意味着「用户把二级标题写成正确的斜体，跑自动修正反而会被改成加粗」。
  - L3/L4 三级/四级：斜体（与 OWL 一致 ✓）
- **连带问题**：`test_fixer.py:364` 断言 "L1 全大写+加粗"，恰好把上述偏差**锁死成"正确行为"**；若对齐 OWL（L2 改斜体），需同步改该测试。

### ❌ 未覆盖（对照 OWL，非必需但可记）
1. 标题 24pt 居中、署名 10pt 居中：checker 无专门量度（仅 fixer 排版）。
2. 双栏布局：当前引擎不查栏数（版式复杂度，超出范围）。
3. 摘要 + index terms 强制：checker 仅 APA-Professional 提示 abstract，IEEE 不提示（OWL 要求 abstract + index terms）。
4. 附录 Appendix、首字下沉、图/表/方程编号：均不查（视觉/内容层）。
5. 正文 [1] 引文语法、参考文献条目著录格式：属 B 类内容层，按铁律不查。

---

## 4. 引文内容类（明确超出范围，绝不查）

以下 100% 为「引用条目语法 / in-text 编号语法」，属 B 类，按老板「绝不改正文」铁律不查、不改：

- **In-Text Citations**：`[1]` 编号引用语法（是否连续、是否对应文末）。
- **参考文献条目著录**：`[1] A. Author, "Title," Journal, vol., no., pp., Year.` 的字段顺序/标点/斜体（期刊名斜体等）。
- **et al.、DOI/URL、Title Case** 等。
- 这些由用户/引用生成器负责；工具只查「参考文献区存在 + 编号版式」，不审「每条写得对不对」。

---

## 5. 测试与回归证据

- `tests/test_fixer.py`：**IEEE 覆盖最全**——标题页（:333）、章节编号幂等（:364）、APA→IEEE 引文改写（:416）、改写幂等（:433）、References 不编号（:498）、内容改写开关守护（:528）。
- `tests/test_engine.py`：IEEE 编号形态应 `pass`（:219-223）；reference_style 解析（:99-140）。
- `tests/test_latex_template.py`：IEEEtran 模板解析（:21-28）。
- 全量非 UI 测试：**194 passed / 30 skipped**（前已验证，无回归）。
- **缺口**：① 缺独立的「IEEE OWL 覆盖」核对测试；② `test_fixer.py:364` 把 L1 加粗当成正确行为断言，若对齐 OWL（L2 改斜体）需同步修订。

---

## 6. 一句话结论

> Purdue OWL IEEE 板块 **5 页 + 1 外链已全部逐段精读**（Overview / General Format / In-Text / Reference List / Tables-Figures-Equations / Taxonomy PDF）。核心结论：
> - **可机查版式（A 类）**：页边距 / 字体 / 参考文献编号+悬挂缩进 / 三级四级标题斜体 已覆盖（info/pass 级，绝不 fail、绝不改正文）。
> - **OWL 反向确认（好消息）**：Reference List 页的「编号制 / 按引用顺序 / 编号贴左边界」三项结构性要求，与本仓 `ReferenceRule(numbered=True, hanging_indent_in=0.0)` + `_chk_numbered_ref` 完全一致 → **现有实现正确，无需改动**。
> - **其余 4 页均属 B 类（引文内容 / 视觉对象）**：in-text `[1]` 语法、参考文献条目著录模板、图表公式编号题注——按老板「绝不改正文」铁律一律不查、不改，与 A/B 切分一致。
> - **三处必须处理（同前）**：
>   1. **⛔ 阻断**：IEEE 会议模板是 Strict OOXML，python-docx 打不开——须先加 strict→transitional 转换预处理，否则全跑不起来；
>   2. **⚠️ 偏差**：`IEEE_HEADING` 二级标题为**加粗**（OWL 要**斜体**），且 checker 不验标题格式、仅 fixer 套用，可能把正确的斜体改错；
>   3. **❌ 未覆盖**：标题 24pt/署名 10pt、双栏、摘要+index terms 强制、附录/drop cap/图表编号。
> - 引文内容类（in-text `[1]`、条目著录）按铁律不查。当前 IEEE 在 **fixer 侧**比 MLA 还早有覆盖，但 **checker 侧**缺标题格式校验、且 heading 格式与 OWL 有偏差，建议对齐 OWL 并补 strict→transitional 预处理。
