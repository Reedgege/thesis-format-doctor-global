# IEEE 三版本对比：会议（Conference）/ 期刊（Journal）/ 杂志（Magazine）

> 约束变化：用户先要求「只用已发材料、不搜索」——当时材料只有会议版，对比缺口已记录。随后用户选 **方案 B（解除搜索限制）**，于是本版拉取了官方期刊/杂志模板做真实对比。
> 全部数值来自**实际下载并直读 XML** 的官方模板 + Purdue OWL + IEEE Author Center 文档，非臆测。

## 0. 证据来源（可追溯）

| 版本 | 使用的官方材料 | 获取方式 |
|---|---|---|
| 会议 Conference | `conference-template-a4.docx` / `conference-template-letter.docx`（用户 @ 投放）+ OWL IEEE General Format + ieee.org/conferences 模板页 | 用户文件直读 |
| 期刊 Journal | `IEEE Access` 官方 Word 模板 `Access-Template-2024.docx`（ieeeaccess.ieee.org 官方下载）+ IEEE Author Center「Structure Your Article / First Footnote」文档 | 官方下载后直读 |
| 杂志 Magazine | `IEEE Electrification Magazine` 官方 Word 模板 `EM-ArticleTemplate-11-2021.docx`（ieee-pes.org 官方下载）+ IEEE Power & Energy Magazine Author's Guide | 官方下载后直读 |

> 说明：IEEE 期刊模板因官网「Template Selector」有 WAF 拦截，无法直接抓包；改用 IEEE Access（真实 IEEE 期刊）官方模板作为期刊代表。IEEE 不同期刊在「单栏/双栏」上有差异（见 §6 注），但作者信息放首页脚注、带 DOI/卷期等特征为期刊通用，且有 Author Center 文档佐证。

---

## 1. 会议版（Conference）具体规格 · 来自你发的两个 .docx

- **纸张**：A4 (8.27×11.69″) 或 US Letter (8.5×11″)，纵向
- **页边距**：上 0.75″ / 下 1″ / 左右 0.62″ / 页眉页脚 0.5″
- **栏数**：双栏（OWL 通用页明确要求 *two columns*）
- **标题区**：`papertitle` 24pt 居中；`papersubtitle` 14pt 居中；`Author` 11pt 居中（多行：姓名/单位/城市国家/邮箱或 ORCID）；`Affiliation` 居中
- **摘要/关键词**：`Abstract` 9pt 加粗 两端对齐；`Keywords` 斜体
- **正文**：`BodyText` 两端对齐，行距约 11.4pt（双栏下正文约 10pt）
- **标题层级（Heading1–5）**：H1 罗马数字+居中；H2 大写字母+左对齐+**斜体**；H3 阿拉伯数字+左缩进+**斜体**；H4 小写字母+左缩进+**斜体**；H5 用于 Acknowledgment/References
- **参考文献**：`references` 8pt 两端对齐，编号列表 `[1]`
- **页脚**：`XXX-X-XXXX-XXXX-X/XX/$XX.00 ©20XX IEEE`（会议版权串）
- **强制前置**：摘要 + index terms；**严禁**自行加页眉页脚/页码（出版方统一打会议名与版号）

---

## 2. 期刊版（Journal，以 IEEE Access 官方模板为代表）具体规格

- **纸张**：约 8.0×10.88″（Letter 派生、略裁切）；页边距 上 0.9″ / 下 0.72″ / 左右 0.51″ / 页眉 0.25″ / 页脚 0.44″
- **栏数**：双栏（Access 特例；IEEE 多数期刊初稿现用**单栏**，见 §6 注）
- **标题**：`PaperTitle` 22pt 粗体（Helvetica）；正文 `Text` 10pt 两端对齐、首行缩进 0.14″（Times 系）
- **作者信息位置（关键特征）**：**不在标题下方，而在首页左下角脚注区**——作者名行 + 机构编号（`1National Institute...` / `2Department...` / `3Electrical Engineering...`）+ 通讯作者 + 基金资助；脚注首段还含 *Date of publication… Digital Object Identifier 10.1109/ACCESS…*（DOI + 收稿/录用日期）
- **摘要/索引词**：`ABSTRACT`（150–250 词，自包含、无缩写/脚注/引用）+ `INDEX TERMS`（必填）
- **标题层级**：粗体、Helvetica；H1 居中 10pt（如 `INTRODUCTION`）、H2 斜体（如 `A. Abbreviations`）、H3 斜体……同样用 I./A./1) 编号体系但字号更小、字体为 Helvetica
- **参考文献**：`References` 8pt 两端对齐，编号列表 `[1]`
- **页眉/页脚**：页眉 = 作者名（`Author Name: Preparation of Papers for IEEE Access`）；页脚 = 卷期（`VOLUME XX, 2017`）
- **作者传记**：参考文献之后含 `AU_Bios`（每位作者小传）
- **公式**：允许（Word 公式编辑器 / MathType）

---

## 3. 杂志版（Magazine，以 IEEE Electrification Magazine 官方模板为代表）具体规格

- **纸张**：US Letter 8.5×11″；页边距 **上下左右均为 1.0″** / 页眉页脚 0.5″
- **栏数**：**单栏**（指南原文："single column as opposed to the double-column formatting found in Transactions papers"）
- **标题区**：`Title` 26pt；`Subtitle` 12pt 斜体；作者行 `By Author 1, Author 2, Author 3, etc.` 12pt 居中
- **正文**：`NormalWeb` **Times New Roman 12pt**
- **标题层级**：`Heading1` 14pt 粗体 / `Heading2` 13pt 粗体 / `Heading3` 粗体 / `Heading4` 粗体斜体……最多 9 级但指南鼓励 ≤2–3 级；**无罗马/字母编号体系**，就是普通粗体标题
- **摘要/关键词/参考文献**：**全部不允许**（指南原文："Articles cannot contain abstracts, keywords, or references"）
- **替代结构**：`For Further Reading`（≤6 条，自引≤2）+ `Acknowledgments` + `Biographies`（每位作者一行：姓名/邮箱/单位/国家）
- **图/表**：`Caption` 9pt 斜体、居中；图/表内嵌正文
- **页眉**：`IEEE Electrification Magazine`（杂志名）
- **篇幅/风格**：约 5000–6000 词；6–8 张图；**不含公式**；面向更广读者（一年级工程生/非技术读者能读懂 80%）

---

## 4. 相同之处（三种版本共通的 IEEE 约定）

- 主字体族均为 **Times New Roman 系**（会议 Times；期刊正文 Times + 标题 Helvetica；杂志 Times New Roman 12pt）
- **编号引用 `[1]` + 按引用顺序编号的参考文献列表**：会议与期刊共有（杂志用「For Further Reading」替代）
- 标准论文结构：标题 → 作者 → 章节 → 图/表 → 参考文献/延伸阅读
- **摘要 + index terms**：会议与期刊强制要求（杂志例外，明确禁止）
- 页眉/页脚带 **IEEE 品牌标识**（会议=版权串、期刊=作者名+卷期、杂志=杂志名）
- 标题层级均采用 `I. / A. / 1) / a)` 式编号体系（杂志除外，用普通粗体标题）

---

## 5. 不同之处（核心交付：三版本对照表）

| 维度 | 会议 Conference | 期刊 Journal（Access 例） | 杂志 Magazine（Electrification 例） |
|---|---|---|---|
| 栏数 | **双栏** | 双栏（Access）/ 多数期刊初稿单栏 | **单栏** |
| 纸张 | A4 或 Letter | Letter 派生 (8.0×10.88) | Letter (8.5×11) |
| 页边距 | T0.75 / B1.0 / L·R0.62 | T0.9 / B0.72 / L·R0.51 | **全 1.0** |
| 标题字号/字体 | 24pt 居中 (Times) | 22pt 粗体 (Helvetica) | 26pt (Times) |
| **作者信息位置** | 标题下**居中网格** | **首页左下脚注**（机构编号+基金+DOI+日期） | 标题下 `By Author…` 居中一行 |
| 页眉/页脚 | **严禁**自加（出版社统一打） | 页眉=作者名；页脚=卷期 | 页眉=杂志名 |
| **摘要** | 必填 | 必填 (150–250 词) | **不允许** |
| **关键词/index terms** | 必填 | 必填 | **不允许** |
| **参考文献** | 编号 `[1]` 列表 | 编号 `[1]` 列表 | **不允许**（改 For Further Reading ≤6） |
| **公式** | 允许 | 允许 | **不允许** |
| 作者传记 | 不典型 | 文末小传 | 文末一行小传（必填） |
| 标题层级风格 | 罗马/字母/数字 + H2–H4 斜体 | 粗体 Helvetica + 编号 | 普通粗体、无编号体系 |
| 篇幅/读者 | 会议录，专业读者 | 学术期刊，专业读者 | 科普向，广读者 |

---

## 6. 补充说明（避免误读）

1. **期刊栏数有差异**：IEEE Access 用双栏；但 IEEE 自 2019 起推的通用期刊 Word 模板（Template Selector 提供）多数要求**单栏初稿**，终稿由出版社排版。因此「期刊=单栏或双栏」因刊而异，但「作者信息在首页脚注 + 带 DOI/卷期 + 必含摘要/index terms」是期刊通用特征。
2. **杂志 ≠ 期刊**：杂志（Magazine，如 Spectrum / Electrification / Power & Energy）定位科普，故去摘要、去参考文献、去公式、单栏、强读者友好；期刊（Transactions/Journal）定位学术，保留摘要+编号参考文献+公式。
3. **对产品的含义（文档格式先生）**：当前引擎 `specs.py` 的 `IEEE` 画像（双栏、标题 24pt、作者居中、摘要+index terms、编号参考文献）最贴近**会议版**；若要支持期刊/杂志，需新增「作者信息=首页脚注」「杂志去摘要/去参考文献/单栏」等差异化规则——属于范围扩展，非当前默认。
4. 已下载的原始模板存于 `D:/AgentSpace/thesis-format-doctor-global/_smoke/`：`IEEE_magazine_template.docx`、`IEEE_journal_Access_template.docx`，供复核查证。
