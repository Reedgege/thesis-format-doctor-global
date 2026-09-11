审查完成。以下为对照 `docs/需求规格与功能说明.md` 第 9 节的逐项结构化报告。本次全程只读，未修改任何文件。

# 「海外版论文格式医生 v1」代码审查报告

## 1. 隔离性（对照 §9.1）

**结论：通过。** 全仓（排除 `.venv`）未出现对国内版 `D:\AgentSpace\论文格式医生\` 的 import 或文件读取；`src/` 内所有相对导入均为 `..engine.*`。对「论文格式医生」的引用仅存在于 `README.md:5` 与规格文档，属说明性文字。

- 【严重程度: 低】【README.md:5 / docs/需求规格与功能说明.md:6】【问题】隔离性目前只靠人工约定，无自动化护栏（如 import 白名单 / 路径断言），后续高版本迭代可能无声破坏。【修复建议】在 `tests/` 增加一个「禁止 import 含 `论文格式医生` / 绝对路径 `D:\AgentSpace\论文格式医生`」的静态断言测试（读取源码文本即可，不触网）。

## 2. 离线安全（对照 §9.2）

**结论：通过。** `requirements.txt` 仅 `python-docx`；`src/` 与 `tests/` 中无 `requests` / `urllib` / `socket` / `http` / `openai` / `anthropic` / `httpx` / `aiohttp` 等；`checker.py:9-17` 只导入 `re/dataclasses/typing`，无 `random`。

- 【严重程度: 低】【仓库根目录 `_run_review.py`、`_codex_exec_help.txt`、`_codex_help.txt`、`_codex_probe.txt`、`_codex_review_help.txt`、`_review_out.txt`、`_review_prompt.txt`】【问题】这些评审工具残留文件位于「全仓」内，`_run_review.py` 通过 `subprocess` 调用外部 `codex.exe`。它们不属于产品引擎，但留在发布目录里会让「全仓零触网」的承诺变得难以自证，且易被误打包。【修复建议】把评审脚本/日志移出仓库（如 `tools/` 或仓库外），或加入 `.gitignore` 并在 README 说明其为开发临时物。

## 3. 确定性（对照 §9.3）

- 【严重程度: 高】【src/engine/docx_reader.py:38-57（关键 42-46）】【问题】`_line_spacing_of` 先判断 `if pf.line_spacing is not None` 就 `return round(float(pf.line_spacing),3)`，而 python-docx 在「固定值 / 最小值」行距（`EXACTLY`/`AT_LEAST`）时 `line_spacing` 返回的是 `Length`(EMU 整数) 而非倍数。已核对本地 `docx/text/parfmt.py`：`_line_spacing()` 在非 `MULTIPLE` 分支 `return spacing_line`（Length）。因此「固定行距 24pt」会被读成 `304800.0`，与期望 `2.0` 相减 → 必然误报 `fail`；下方针对 `EXACTLY/AT_LEAST` 的返回 `None` 分支实际是死代码。这是最常见的现实文档之一。【修复建议】先看 `pf.line_spacing_rule`：仅当规则为 `MULTIPLE`（或 `None` 且 `isinstance(value,(int,float))`）时按倍数返回；`EXACTLY/AT_LEAST` 直接返回 `None`，或换算为「磅/12」再比较并单独提示。
- 【严重程度: 低】【src/engine/checker.py:19-25、40-60】【问题】比对阈值本身合理且带名字、有 fail/warn 分档，`measured is None` 时降级为 `warn` 不会误判 `fail`，Counter 众数在 CPython 下按插入序取平局、结果稳定。但 `MARGIN_TOL_WARN=0.03"`（≈0.76mm）、`FAIL=0.06"` 的 warn 带极窄，Word 模板经换算/四舍五入后容易落到 warn 而非 pass，阈值依据未写明。【修复建议】在常量处补注释说明来源（如「1/16 英寸以内视为排版误差」），并考虑把 warn 带略放宽到 ≈0.05"。

## 4. TargetProfile 收敛（对照 §9.4）

- 【严重程度: 高】【src/data/ai_questionnaire_template.json:2-3 与 src/engine/questionnaire.py:140-149】【问题】模板 `_how_to_fill` 明确要求「AI 应只产出 `filled` 对象里的扁平键值」，但 `target_from_dict` 只兼容顶层扁平键与嵌套 `page`，**完全不识别 `filled` 键**。若 AI 照说明输出 `{"filled":{...}}`，导入会静默得到空 TargetProfile，学校要求整体丢失且无任何报错，用户以为已生效。四通道在 AI-JSON 这条上并未真正收敛。【修复建议】`target_from_dict` 增加对 `filled`（以及 `data`/`example_filled` 之外的常见包装）的解包；或统一模板措辞为「输出顶层扁平 JSON」，并在解析失败/全空时显式报错而非静默通过。
- 【严重程度: 中】【src/engine/questionnaire.py:209-217】【问题】`target_from_template_docx` 把 `reference_hanging_indent_in` / `reference_numbered` 取自 **spec**（`get_spec(spec_key)`），却把这些维度的 `resolution` 标成 `"school"`（213-217）。这与「引用取规范」和函数自身 docstring「引用风格沿用所选规范」自相矛盾，会让报告裁决表把引用维度误标为「学校模板」，并让 `merge_school` 把它们当作学校显式覆盖处理。【修复建议】学校模板通道不设置任何 `reference_*` 字段（保持 `None`），仅由 spec 目标提供引用维度；若确需标注则应为 `"spec"`。
- 【严重程度: 中】【src/engine/questionnaire.py:154-156（配 170-176）】【问题】当 AI/问卷只给了 `reference_style`（如 `"IEEE"`）而未给 `reference_hanging_indent_in`/`reference_numbered` 时，这两项保持 `None`，不会从 `get_spec(reference_style)` 反填。结果同一「选 IEEE」在两个通道得到的结构不一致（spec 通道有 numbered=True/悬挂 0，问卷通道为空），违反「四通道同一结构」。【修复建议】在 `reference_style` 有效且对应字段缺失时，用 `get_spec(ref_style).reference` 反填并标注来源为 `spec`。
- 【严重程度: 中】【src/ui/app.py:139-148、src/main.py:33-44】【问题】两条目标构建都用 `elif` 链（模板 > AI-JSON > 问卷），只能生效一个通道；规格 FR2 写的是「任选其一或**组合**」。同时选了模板与 AI 问卷时，问卷会被静默忽略，且无法捕获两者冲突。【修复建议】把「组合 + 冲突归并」逻辑下沉到 engine（见 §8），对同时提供的通道做合并并记录冲突。
- 【严重程度: 中】【src/engine/report.py:71-84】【问题】冲突检测只遍历 `target.resolution` 中 `origin=="school"` 的维度，并用 `spec_page = spec.page.__dict__` 取规范值。于是：(a) `heading_levels` 被学校覆盖时不产生冲突行（`spec_page.get` 为 None）；(b) `reference_*` 冲突永远不可能出现（不在 `spec.page`）；(c) §9.4「冲突项是否都被捕获」的答案是**没有**。【修复建议】冲突比对应覆盖 page + heading_levels + reference 全字段，规范侧取自完整 `StyleSpec`（而非仅 `page`）。
- 【严重程度: 中】【src/engine/questionnaire.py:248-263】【问题】`merge_school` 对来自问卷/AI 的 `school` 参数一律写 `resolution[...] = "school"`，用户手填/ AI 的来源标签被统一成「学校模板」，与 `report.ORIGIN_LABEL` 的 `user`/`ai` 设计不符，报告来源失真。【修复建议】以 `school.source`（user/ai_import/school）为准写标签，或让调用方传入来源枚举。
- 【严重程度: 低】【src/engine/report.py:17-22、94 与 src/engine/questionnaire.py:66、185】【问题】`resolution` 实际会写入 `"ai_import"`、`"questionnaire"`，而 `ORIGIN_LABEL` 只有 `user`/`ai` 键，导致这两类来源在报告中显示为英文原值（`app.py`/`main.py` 对问卷传了 `"user"` 尚可，AI 通道 `target_from_ai_json` 固定 `source="ai_import"` 必然漏翻译）。【修复建议】统一来源枚举常量并集中映射。

## 5. 规范默认值合理性（对照 §9.5）

- 【严重程度: 高】【src/engine/specs.py:76,91,106,121,136 与 src/engine/questionnaire.py:94-98、checker.py:114-164】【问题】规格 §5 与 README 反复声明「页面排版由学校/问卷决定，规范的 page 只是兜底参考」，但 `target_from_spec` 把 spec 的 page 值全部标为 `resolution="spec"`，checker 随即把它们当硬性要求判定，报告也标为「通用规范」。结果是：只选 APA、未上传模板的学校风格论文（如学校 1.5" 页边距）会直接 `fail`，与产品自身的设计铁律冲突。【修复建议】明确二选一：(a) 规范页值降级为 `info`（仅提示不判 fail），(b) 仅在无学校/问卷输入时才作为基线并明确标注「兜底默认值，非规范强制」，或新增「仅引用规范」模式不检查页面。
- 【严重程度: 中~高】【src/engine/specs.py:121、133-136、106（参 69 注释）】【问题】页面默认值依据不均：APA/MLA 的 1"、双倍、12pt TNR 与其官方学生论文指引一致，尚可辩护；但 IEEE 写 12pt/1.5 倍（`specs.py:121`）与其真实会议模板「10pt、单倍、双栏」相反（注释 69 自承「这里取论文常用值」），Harvard 自承「无单一官方」（146），Chicago 的页边距/字体/行距并非 CMOS 规定。这些被当作规范要求执行，可能产生大量无依据的 `fail`。【修复建议】对 IEEE/Chicago/Harvard 将 page 默认值设为 `None`（不强制）或移入「常见值/参考」而非「要求」；至少在 UI/报告中标注「非官方强制，仅参考」。
- 【严重程度: 中】【src/engine/specs.py:124 与 src/engine/checker.py:63-106、193-195】【问题】`TargetProfile` 携带 `reference_hanging_indent_in` / `reference_numbered`（IEEE 悬挂=0、numbered=True），但 checker 从未读取这两项，`_check_reference` 只按 style 做文本启发式。等于规范里写了一部分引用要求却无任何检查落地，给用户「已覆盖」的错觉。【修复建议】要么实现基于批注/缩进的悬挂缩进与编号一致性检查，要么在报告中明确这些维度为「未检查」。

## 6. docx_reader 健壮性（对照 §9.6）

- 【严重程度: 高】【src/engine/docx_reader.py:42-46】【问题】固定行距被误读为天文数字倍数（详见 §3 第 1 条），属读取健壮性缺陷。【修复建议】同 §3。
- 【严重程度: 中】【src/engine/docx_reader.py:86-93】【问题】只读第一个 section 的页边距。含封面/正文不同页边距的多分节文档（很常见）会被按第一节判定，可能误报。【修复建议】遍历 `doc.sections`，取众数或对不一致情况给出「文档存在多种页边距」提示。
- 【严重程度: 中】【src/engine/docx_reader.py:114-126】【问题】字体/字号只统计 run 级显式 `run.font.name`/`run.font.size`；当字体定义在样式（style）或 `docDefaults` 上（Word 最常见做法），会得到「未知」并报 warn，属对合法文档的误判。【修复建议】按 run → paragraph style → 文档默认 的层级回退解析字体/字号。
- 【严重程度: 中】【src/engine/docx_reader.py:82 与 src/main.py:48-60】【问题】受保护/加密/损坏或非 docx 文件在 `Document(path)` 处抛异常：GUI 有兜底（`app.py:160`），但 CLI `cmd_check` 无 try/except，会向用户抛裸 traceback；无分节、只读文件亦无专门处理。【修复建议】engine 层封装读取异常为可读错误（如 `DocxReadError`），CLI 捕获并输出友好信息 + 非零退出码。
- 【严重程度: 低】【src/engine/docx_reader.py:60-74】【问题】标题样式仅识别 `Heading n` / `标题n`（需带空格或中文），对本地化样式名（`Überschrift 1`、`Titre 1`）、无空格写法 `Heading1`、学校自定义样式（`ThesisHead1`）无容错。【修复建议】用 `style.style_id`/`base_style` 识别内置 Heading 家族，正则放宽为 `^heading\s*(\d)` 与中文 `标题\s*(\d)`，自定义样式可回退到 outline level（`pPr.outlineLvl`）。
- 【严重程度: 低】【src/engine/docx_reader.py:77-78、149】【问题】参考文献标题用精确匹配 + `len(text)<=30`（魔法数字），`References Cited`、`参考书目`、`Bibliography:` 等变体会漏检。【修复建议】改为「去标点后前缀/包含匹配」，阈值提为具名常量。
- 说明：无分节场景 `if doc.sections`（86 行）在 python-docx 下几乎恒为真（始终≥1 个 section），该判断对「无 sectPr」保护有限；真正无页边距时靠 `_inch(None)` 返回 None 再降级为 warn，行为可接受。

## 7. 可维护性（对照 §9.7）

- 【严重程度: 中~高】【src/engine/questionnaire.py:170、209、50；src/engine/checker.py:72-106；src/engine/report.py:17-22】【问题】「新增一套规范是否只需改 specs.py 一处」的答案是**否**：规范 key 白名单在 `questionnaire.py` 出现两次（170、209），问卷 `reference_style` 的 options 又硬列一次（50），checker 对每种 style 写死分支（72-106），report 还有来源映射。新增规范至少改动 3~4 个文件。【修复建议】白名单统一用 `SPEC_ORDER`（排除 Other）派生；把 `_check_reference` 改为由已定义但未使用的 `ReferenceRule.checks`（`specs.py:42`，目前是死字段）驱动，实现数据驱动。
- 【严重程度: 中】【src/engine/specs.py:28-30 vs src/engine/questionnaire.py:29-42】【问题】默认值 `1.0 / "Times New Roman" / 12.0 / 2.0` 在 `PageLayout` 默认与问卷 schema 中各写一份，存在漂移风险；另有 `docx_reader.py:149` 的 `30`、正则 `\d{4}` 等散落魔法值。【修复建议】以 `specs.py`（或单一常量模块）为唯一数值源，问卷 schema 由其一并生成。
- 【严重程度: 低】【src/engine/specs.py:16】【问题】`from dataclasses import dataclass, field, asdict` 中 `field` 未使用。【修复建议】移除未用导入。

## 8. CLI / GUI 引擎复用（对照 §9.8）

- **结论：共用引擎，通过。** `src/ui/app.py:15-22` 与 `src/main.py:18-24` 导入同一套 `engine`（checker/report/questionnaire/docx_reader/specs），均消费同一 `report.markdown`，无重复实现检查逻辑。
- 【严重程度: 中】【src/ui/app.py:136-149 与 src/main.py:29-45】【问题】目标构建编排（spec → template/ai/questionnaire 的优先级合并）在 GUI 与 CLI 各写一份几乎相同的 `_build_target`，属重复逻辑，未来优先级/组合规则改动易两边不一致。【修复建议】把该编排抽到 engine（如 `questionnaire.build_target(spec, template=..., ai=..., questionnaire=...)`），两入口只做参数适配。

## 9. 测试盲区（对照 §9.9）

- 【严重程度: 中】【tests/smoke_test.py:58-89】【问题】冒烟脚本只 `print` 报告、**零断言**，无法让 CI/回归失败；且第 20 行 `import target_from_ai_json` 后从未调用，AI-JSON 通道实际未被测试。边界 case 均缺失：空文档、受保护/加密、损坏/非 docx、多分节、固定行距（能直接暴露 §3 的 bug）、样式继承字体、本地化标题、超长文档、`filled` 包装 JSON、merge 冲突列表。【修复建议】改造为 `pytest`，为上述每类补断言（至少覆盖：固定行距不误判、空文档不崩溃、AI `filled` 能导入、学校冲突出现在报告中）。（注：本次为只读，未改动测试。）

---

## 总体评价

整体架构是对的：引擎/UI 分层清晰、四通道确有统一入口、离线性与目录隔离经得起 grep 核查，代码可读性良好，阈值也做了命名与分档。但 v1 在「收敛完整性」与「设计自洽」上有几处会真实影响结论正确性的缺陷：**固定行距被读成 EMU 导致必然误判**（最常见文档形态）、**AI 填表 `filled` 包装被静默丢弃**、**规范页值被当作硬性要求执行**（与「页面归学校」的立规相矛盾）、**冲突项漏检（heading/reference）**，以及**新增规范并非只改一处**。这些不改，验收标准里「全部维度 pass / 正确 fail」在真实文档上会失真。

## 优先修改清单

- P0（先修，直接影响正确性）：`docx_reader.py:42-46` 固定行距误判；`questionnaire.py:140-149` 支持/统一 AI JSON `filled` 包装；`report.py:71-84` 冲突检测覆盖 heading 与 reference。
- P1（设计自洽 + 收敛）：`target_from_spec`/checker 明确规范页值是「兜底参考」而非「要求」；`questionnaire.py:209-217` 模板通道不再把 spec 引用值标成 school；为「只给 reference_style」反填 spec 引用字段；把通道组合逻辑下沉 engine。
- P2（健壮性与可维护性）：多分节页边距、样式级字体回退、受保护/损坏文档的友好报错；规范 key 白名单与引用检查改为数据驱动（用起 `ReferenceRule.checks`）；清理 `specs.py:16` 未用导入与根目录开发残留物。
- P3（测试）：冒烟脚本改断言式并补齐 AI-JSON、空/只读/超长/固定行距/本地化标题等边界用例。

---

## 修复状态（2026-09-11 已全部应用）

用户确认「全量 P0–P3」后，按本报告逐项修复，**pytest 10 项断言全绿**，CLI/冒烟端到端通过：

| 项 | 修复 | 验证 |
|---|---|---|
| P0-① | `docx_reader._line_spacing_of` 先判 `line_spacing_rule`，仅 MULTIPLE 按倍数返回；EXACTLY/AT_LEAST 返回 None | `test_fixed_line_spacing_not_false_fail` ✅ |
| P0-② | `target_from_dict` 解包 `filled`/`data` 等包装；结果为空显式抛 `ValueError`；`ai_questionnaire_template.json` 措辞同步 | `test_ai_filled_wrapper_imported` / `test_ai_empty_raises` ✅ |
| P0-③ | `report._collect_conflicts` 覆盖 page+heading+reference 全字段，规范侧取完整 `StyleSpec`；显式来源间冲突也入报告 | `test_school_conflict_in_report` ✅ |
| P1 | `checker` 对 spec 来源的页面/标题维度降级为 info（不强制 fail），引用始终硬判；模板通道不再把 spec 引用标 school；`reference_style` 单独给出时反填规范引用字段；四通道合并下沉 `build_target` | `test_spec_page_only_info_not_fail` / `test_reference_backfill` ✅ |
| P2 | 多分节页边距取众数+不一致标记；样式级字体回退；本地化标题（英/中/德/法/西）+ outlineLvl；损坏文档抛 `DocxReadError`；规范白名单与引用检查数据驱动（`ReferenceRule.checks`） | `test_multi_section_margins` / `test_garbage_file_raises` / `test_ieee_numbered_pass` / `test_apa_missing_year_warn` ✅ |
| P3 | 新增 `tests/test_engine.py`（10 项断言）；`smoke_test.py` 改用 `build_target` | `pytest` 全绿 ✅ |

修复过程中又发现并修正 1 个评审未点名但测试暴露的 bug：`_check_reference` 聚合「最严重项」误用 `min`（取了最轻的 pass），改为 `max`，否则缺年份的文档会被误判 pass。

国内版确认零影响：本次修改仅在 `thesis-format-doctor-global\` 内进行。