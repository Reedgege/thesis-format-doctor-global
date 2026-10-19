# APA 7th 学生版 / 专业版 · 完整参数对照（可复制）

> 来源：`src/engine/specs.py`（`_apa_spec` / `_APA_STUDENT` / `_APA_PRO`）+ `questionnaire.target_from_spec`。
> 设计原则：**一套底座 + 差异单摘**。正文/标题/参考文献/页面规则两版完全共用同一套值；
> 真正的差异只有 5 个 tier 标记 + `running_head_max_chars`（学生 `None` / 专业 `50`）+ 报告说明文案。
> 其余所有 APA 专属检查（页码/参考文献另起一页/字体清单/字母序）两版**同值同查**。

---

## 一、参数总表（学生版 vs 专业版）

| # | 参数（代码字段） | 学生版值 | 专业版值 | 共用? | 引擎行为 |
|---|---|---|---|---|---|
| 1 | 上页边距 `margin_top_in` | 1.0″ | 1.0″ | ✅ | checker 查(info) / fixer 可改 |
| 2 | 下页边距 `margin_bottom_in` | 1.0″ | 1.0″ | ✅ | 同上 |
| 3 | 左页边距 `margin_left_in` | 1.0″ | 1.0″ | ✅ | 同上 |
| 4 | 右页边距 `margin_right_in` | 1.0″ | 1.0″ | ✅ | 同上 |
| 5 | 正文字体 `font_family` | Times New Roman | Times New Roman | ✅ | checker 查(info) / fixer 可改 |
| 6 | 正文字号 `font_size_pt` | 12.0 | 12.0 | ✅ | 同上 |
| 7 | 行距 `line_spacing` | 2.0（双倍） | 2.0（双倍） | ✅ | 同上 |
| 8 | 正文对齐 `body_alignment` | left | left | ✅ | 同上 |
| 9 | 首行缩进 `first_line_indent_in` | 0.5″ | 0.5″ | ✅ | 同上 |
| 10 | 段后间距 `space_after_pt` | 0.0 | 0.0 | ✅ | 同上 |
| 11 | 引用风格 `reference.style_key` | APA | APA | ✅ | **硬查**（引用归规范） |
| 12 | 参考文献悬挂缩进 `hanging_indent_in` | 0.5″ | 0.5″ | ✅ | 硬查 / fixer 可改 |
| 13 | 是否编号 `numbered` | False（著者-出版年） | False | ✅ | 硬查 |
| 14 | 引用确定性检查 `checks` | has_year_paren, author_year_order | 同左 | ✅ | 硬查 |
| 15 | 标题层级数 `heading_levels` | 5 | 5 | ✅ | fixer 按级格式化 |
| 16 | 5 级标题格式 `heading_formats` | APA_HEADING（见下） | 同左 | ✅ | fixer 施加 |
| 17 | 标题页基础 `title_page` | 标题居中+粗体+作者居中 | 同左 | ✅ | 仅说明，未自动生成字段 |
| 18 | 页码位置 `page_number_position` | "top-right" | "top-right" | ✅ | checker(info) 提示 |
| 19 | 参考文献另起一页 `reference_new_page` | True | True | ✅ | checker(info) 提示 |
| 20 | 字体允许清单 `font_allowlist` | 7 种（见下） | 7 种（见下） | ✅ | checker(info) 提示 |
| 21 | 参考文献字母序提示 `reference_alpha_sort_hint` | True | True | ✅ | checker(info) 提示（不改内容） |
| 22 | **页眉短标题上限 `running_head_max_chars`** | **None** | **50** | ❌ 差异 | 仅专业版生效：checker(info) 查长度 |
| 23 | **`running_head`（需页眉短标题）** | **False** | **True** | ❌ 差异 | 仅专业版：缺→info 提示"应有 running head"；学生版不查 |
| 24 | **`expect_abstract`（须含摘要+关键词）** | **False** | **True** | ❌ 差异 | 仅专业版：缺→info 提示 |
| 25 | **`expect_author_note`（须含作者注）** | **False** | **True** | ❌ 差异 | 仅专业版：缺→info 提示；学生版仅当多写时温和提示"非必需" |
| 26 | **`title_page_kind`** | "student" | "professional" | ❌ 差异 | 仅驱动报告文案（两套标题页字段结构暂未自动生成/校验） |
| 27 | **`notes`（前置材料说明）** | 学生版文案 | 专业版文案 | ❌ 差异 | 进报告展示 |

> 结论：**学生版与专业版真正不同的只有第 22–27 行这 6 个字段**；其余全部共用同一底座。

---

## 二、5 级标题格式（两版完全相同 · `APA_HEADING`）

| 层级 | 对齐 | 加粗 | 斜体 | 缩进 | 接排(run-in) | 说明 |
|---|---|---|---|---|---|---|
| L1 | center | ✅ | — | — | — | 居中、粗体、Title Case |
| L2 | left | ✅ | — | — | — | 左对齐、粗体、Title Case |
| L3 | left | ✅ | ✅ | — | — | 左对齐、粗体斜体、Title Case |
| L4 | left | ✅ | — | ✅ | ✅ | 缩进、粗体、句末句号、与正文接排 |
| L5 | left | ✅ | ✅ | ✅ | ✅ | 缩进、粗体斜体、句末句号、与正文接排 |

> 注：L4/L5 仅当文档使用了 Word Heading 4 / Heading 5 样式才被识别与格式化；客户用手动加粗而未设样式时，引擎维持原样并在报告提示"层级为推断/请人工核对 L2–L5"。

---

## 三、字体允许清单（两版完全相同 · 7 种）

Times New Roman · Calibri · Arial · Georgia · Lucida Sans Unicode · Lucida Console · Courier New
（不在清单内的正文字体会触发 info 提示，但**不强制改**——APA 允许多种可读字体，最终以学校要求为准）

---

## 四、fixer 实际会改哪些（修改范围）

以下为**纯格式**维度，fixer 可按规范值施加修正；学生版与专业版值相同，故两版修改范围一致：

- 页边距（上/下/左/右 1.0″）
- 正文字体（Times New Roman）、字号（12pt）、行距（双倍）
- 正文对齐（left）、首行缩进（0.5″）、段后间距（0pt）
- 标题 5 级格式化（按上表 L1–L5）
- 参考文献悬挂缩进（0.5″）

**fixer 不动（只提示）的部分**：running head、摘要/关键词、作者注、标题页字段结构——这些涉及"增删内容块"而非纯格式改写，按承诺保持"只改格式"，由用户在软件里看到提示后自行处理。

---

## 五、checker 会提示哪些（只提示、不改内容）

| 维度 | 学生版 | 专业版 | 级别 |
|---|---|---|---|
| 页码位置 top-right | 查（info） | 查（info） | 共用 |
| 参考文献另起一页 | 查（info） | 查（info） | 共用 |
| 字体允许清单 | 查（info） | 查（info） | 共用 |
| 参考文献字母序 | 提示（info） | 提示（info） | 共用 |
| 页眉短标题 running head | **不查** | 查：缺失/超 50 字（info） | 差异 |
| 摘要 + 关键词 | **不查** | 缺失→提示（info） | 差异 |
| 作者注 | 仅"多写了"温和提示 | 缺失→提示（info） | 差异 |

---

## 六、其他风格不受影响

MLA / IEEE / Chicago / Harvard / Other 在 `StyleSpec` 里上述 APA 专属字段均取**中性默认值**
（`running_head=False`、`expect_abstract=False`、`expect_author_note=False`、`running_head_max_chars=None`、
`page_number_position=None`、`reference_new_page=False`、`font_allowlist=()`、`reference_alpha_sort_hint=False`），
故第二节第 18–27 行所有检查对它们**自动跳过、零误触**。

---

## 七、不修改 / 范围不覆盖的地方（需用户手动处理）

以下属于 APA 7 的合规要求或常见内容，但软件**不自动修改、不自动生成**——只提示或完全不碰，
由用户在提示后自行处理（与"能做的做、做不了的在软件里说清楚"的边界原则一致）：

| 项目 | 软件现状 | 用户需自行处理 |
|---|---|---|
| 页眉短标题 running head（专业版） | 仅 info 提示"应有 / 超 50 字"，**不自动生成** | 手动在页眉加短标题 |
| 摘要 + 关键词 Abstract & Keywords（专业版） | 仅 info 提示缺失，**不自动生成** | 手动补摘要段 |
| 作者注 Author Note（专业版） | 仅 info 提示缺失，**不自动生成** | 手动补作者注段 |
| 标题页字段结构 | `title_page` 仅规定"标题居中+粗体+作者居中"，**不自动生成/校验**"姓名/院系/课程号/教师/日期"或"作者+单位(上标)+作者注"两套字段 | 手动按规范排标题页 |
| 页码域 PAGE 插入 | `page_number_position` 仅 info 提示，**spec 路径不自动插入页码域** | 手动插入页码（或走学校模板） |
| 交叉引用（图/表/章节引用） | **不做** | 自行维护 |
| 公式 / 方程 | **不做** | 自行处理 |
| 复杂表格（多级表头 / 合并单元格） | **不做** | 自行排版 |
| 三线表 | 仅保守版：去竖线 + 保留顶/表头底/底线；**不强行删内部横线** | 自行核对 |
| 引文 ↔ 参考文献 对应校验 | **不校验**（只查参考文献格式，不查引用是否都有对应条目） | 自行核对 |
| 参考文献内容排序 | **只提示字母序，不重排内容** | 自行排序 |
| L4/L5 标题（无 Heading 4/5 样式） | **维持原样，不重新定级** | 手动设 Heading 4/5 或由人工核对 |
| 内容语义（拼写 / 语法 / 查重） | 超出格式范畴，**不做** | 另用工具 |

---

## 八、属于「提示」的维度汇总（只提示、不改内容、info 级）

以下维度软件**只出提示、绝不自动改写文档内容**，用户据提示自行决定是否需要改：

- 页码位置应为 top-right（两版同查）
- 参考文献应另起一页（两版同查）
- 正文字体应在 7 种允许清单内（两版同查；不在清单仅提示，不强制改）
- 参考文献条目应按首作者姓氏字母序排列（两版同提示；**仅提示不重排**）
- 专业版专有：应有 running head（缺失 / 超 50 字 → 提示）
- 专业版专有：应含 Abstract + Keywords（缺失 → 提示）
- 专业版专有：应有 Author Note（缺失 → 提示）
- 学生版专有：若写了 Author Note → 温和提示"通常非必需"
- 报告末尾固定区块「Capability Boundary / Known Limitations」：集中列明上述所有"不做的项"

> 一句话区分：**改格式**（页边距/字体/字号/行距/对齐/缩进/标题5级/参考文献悬挂缩进）fixer 能自动做；
> **纯内容块**（running head/摘要/作者注/标题页字段/页码域）与**跨文档语义**（引文对应/排序/公式/表格/交叉引用）只提示或不动，由用户手动处理。
