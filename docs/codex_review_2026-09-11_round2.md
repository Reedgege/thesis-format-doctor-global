# 海外版论文格式医生 v2 · 只读代码审查报告

**审查方式说明**：全程只读（`Get-Content` / Python 只读读取 / 阅读 python-docx 1.2.0 源码），未修改任何文件、未执行任何写操作；因沙箱为只读且要求"不执行写操作"，**未运行 pytest**（pytest 会写 `.pytest_cache`/临时文件）。结论基于静态阅读与依赖库源码核对。评判基准为 `docs/需求规格与功能说明.md`（FR1–FR8 / NFR1–NFR6 / §5 数据模型 / §8 已知边界 / §9 审查重点）。

---

## 一、A 组（既有基线项 1–8）

### A1 · 隔离性
**已核实无问题。** 全仓没有任何代码 import 或读取 `D:\AgentSpace\论文格式医生\`；`rg` 替代检索只命中 `README.md:5` 与文档正文中的文字说明，`src/`、`tests/` 中无该国路径、无跨目录 sys.path、无共享资源（`src/*.py` 全部使用相对/包内路径）。`.venv` 独立。FR3/NFR3 成立。

### A2 · 离线边界（FR8 / NFR1）
**已核实无问题。** 逐个检索 `urllib`/`requests`/`socket`/`http.client`/LLM SDK/`subprocess`：
- 联网仅出现在 `src/license/license.py:29-30,107-108,120-122`（`urllib.request`），符合"唯一允许联网模块"。
- `src/engine/*.py` 与 `src/engine/fixer.py` 零网络、零 LLM、无 socket、无子进程；`src/ui/app.py` 也无网络。
- 未发现 `openai`/`anthropic`/`dashscope` 等 SDK 依赖（`requirements.txt` 仅 python-docx）。

结论：格式引擎/修正引擎离线边界成立，不算违规。

### A3 · 确定性（NFR2）
**基本成立，但阈值常量有可疑笔误（nit，非误判源）。**
- `src/engine/checker.py:161-313` 无随机、无时间依赖；平局由 `Counter.most_common` 按首次出现顺序决定，同文档同参数结果稳定；比较均带容差，无浮点直接相等。
- 【严重程度：低】【`src/engine/checker.py:21-22`】`MARGIN_TOL_WARN=0.05` **小于** `MARGIN_TOL_FAIL=0.06`，使 `_cmp_margin`（55-63 行）的 warn 带只剩 0.01″（0.05<x≤0.06），几乎不可达，疑似应为 `WARN=0.05 / FAIL=0.1` 之类。【修复建议】确认原始意图并让 `WARN < FAIL` 具有实际带宽；或在注释里说明该"窄 warn 带"是有意为之。

### A4 · 五通道收敛
**收敛成立，但来源标注与优先级存在真问题（见 B 组补充项）。**
- 五通道（spec `questionnaire.py:106`、template `242`、ai `234`、questionnaire `169`、latex `325-351`）确实都产出同一个 `TargetProfile` 并由 `build_target` 合并，唯一入口成立。
- 真问题：
  1. 【中】`src/engine/questionnaire.py:342-356`：叠加顺序为 template → ai → questionnaire → latex，后写覆盖先写，**实际优先级 latex > questionnaire > ai > template，与 FR2 明文"template > ai > questionnaire > latex"相反**。【修复】按规格把 extras 反序（或显式给每个来源定优先级权重）后再叠加。
  2. 【中】`src/engine/questionnaire.py:298-309` 只对 page/heading 维度调用 `_register`；`311-321` 对 `reference_style/reference_hanging_indent_in/reference_numbered` **不登记冲突** → LaTeX 与问卷/AI 在引用维度上的互相冲突会静默丢失（FR2 要求"来源间冲突记录进 conflicts"）。
  3. 【中】LaTeX 来源的引用维度 `resolution` 被标成 `school`（见 B15-b）。
- `resolution` 其余标注准确：spec 页面维度 = `spec` 且 checker 降级为 info；template = `school`；问卷/AI = `user`/`ai_import`；仅给 reference_style 时反填悬挂/编码并标 `spec`（`questionnaire.py:205-213`，与 `tests/test_engine.py:149-153` 一致）。

### A5 · 规范默认值合理性
**已核实无问题（语义正确）。** `specs.py:20-40,163-200` 把页面值定义为"该风格典型文稿兜底参考"，`checker._downgrade`（checker.py:78-82）仅对 `is_page_or_heading=True` 且 `origin=="spec"` 的 warn/fail 降级为 info；引用维度（`reference`、`reference_hanging_indent`）不做降级、始终硬判；`report.py:167-168` 明确写出"通用规范页面维度仅作参考、引用始终以规范为准"，符合 FR4。`Other` 规范页面全 None，不产生强制判定，符合 FR1。

### A6 · docx 读取健壮性
**大体成立，但有 3 处缺口（1 中 2 低）。**
- 已核实：三级回退 run→段落样式→Normal（`docx_reader.py:163-189,373-392`）；固定值/最小值不误当倍数（`97-125`，并有 `tests/test_engine.py:77-89` 守护）；多分节取众数+不一致标记（`292-316`）；损坏/非 docx/加密抛 `DocxReadError`（`282-288`）；段落缺省语义按 Word 补齐（对齐缺省=left、缩进/段后=0，`350-357`）；参考文献识别有"后随条目"佐证且两份实现已合一（`210-247,443-449`）。
- 【中】`src/engine/docx_reader.py:405`：字体/字号采样遍历的是**全部**顶层段落（含标题段与参考文献段），与同段注释"正文段落 + 表格段落"（366 行）及"正文字体"语义不符。标题多、正文少的短文档里众数可能被标题字体/字号带偏，进而同时影响检查结论与修正目标。【修复】字体/字号统计只取 `body_paras`（可另设"标题字体"维度），或至少在报告里标注"含标题采样"。
- 【低】`src/engine/docx_reader.py:378`：只读 `run.font.name`（python-docx 映射 `w:rFonts` 的 ascii/hAnsi），不读 `w:eastAsia`；中文 Word 文档中常见的"仅设中文字体"的 run 会读到 None 并回落到 Normal → 字体误判/误报"未能读取"。【修复】补充 eastAsia/复杂文种字体读取。
- 【低】`src/engine/docx_reader.py:283-288` 只把 `Document(path)` 包进 try，其后 `doc.paragraphs`（329 行）等未保护；"能打开但内部 XML 异常"的文档会抛出非 `DocxReadError` 异常，`cmd_check` 只捕获 `DocxReadError` → 命令行直接打 traceback，违反 FR3"损坏文档须抛明确错误、不得崩溃"。【修复】把整个读取体包进 try，统一转 `DocxReadError`。
- 表头/页脚/脚注/文本框不读取：FR3 未要求、§8 已列为延期项，不算缺陷。

### A7 · 可维护性
**基本成立，若干重复/死代码（均 nit）。**
- 单一真相到位：规范值集中 `specs.py`、问卷字段集中 `QUESTIONNAIRE_SCHEMA`、来源标注集中 `resolution`、参考文献定位集中 `find_reference_section`、容差集中 `checker.py` 顶部。
- nit：「低」`fixer.py:32` 与 `checker.py:37` 各有一份完全相同的 `ALIGN_LABEL`（report 只用 checker 那份），属重复实现；`fixer.py:42-47` 的 margin 映射又抄了一遍 `checker.py:166-169`。
- nit：「低」`main.py:23-26` 导入的 `target_from_spec/target_from_ai_json/target_from_dict/target_from_template_docx/merge_school` 全部未使用；`ui/app.py:17-21` 同样存在死导入。
- nit：「低」`license.py:37` `HEARTBEAT_DAYS` 定义后从未引用（见 B12-e，不是纯风格问题，是功能未实现）。
- nit：「低」`ui/app.py:30-41` `_FIELD_RANGE` 与问卷 schema 未打通（新增字段需改两处）；`src/__init__.py` `__version__="1.0.0"` 与需求文档"v2"不一致。

### A8 · CLI / GUI 引擎复用
**引擎复用成立；GUI 功能面有缺口。**
- 已核实：CLI 与 GUI 都直接 import `src/engine/*` 与 `src/license/license.py`，门禁只有一套实现（`require_fix_entitlement` / `record_fix_used`），无"第二个门禁"。
- 【中】GUI 没有 LaTeX 通道：`src/ui/app.py:200-206` 调用 `build_target` 时不传 `latex`，界面也无 LaTeX 上传控件（`59-96`），而 FR2-⑤ 与 FR7 明确要求 GUI 支持"传模板（docx/LaTeX）"。【修复】GUI 增加 .cls/.sty/.tex 选择并传入 `build_target(..., latex=...)`。
- 【低】`src/ui/app.py:258` 对用户在保存对话框里**显式选择**的路径也强制 `unique_output_path`，与 CLI 语义（仅默认输出名防覆盖，`--out` 以用户为准）不一致；用户确认覆盖却被静默改名成 `-2`。【修复】GUI 区分"默认名"与"用户显式选择"，仅在默认名冲突时加序号。
- 【低】`src/ui/app.py:239` `except (DocxReadError, Exception)` 冗余且吞掉所有异常，建议 `except (DocxReadError, ValueError, OSError)` + 兜底。

---

## 二、B 组（第二轮新增 9–16）

### B9 · 「内容零改动」是否真守得住 —— **不成立，存在两处真 bug（最高优先级）**

写入路径穷举清单见第三节。结论先行：**绝大多数路径确实只动格式，但参考文献编号这一条会写文字，且其实现方式在极端输入下会丢失/错位内容。**

**【严重程度：高（真 bug）】【`src/engine/fixer.py:183`】** `para.runs[0].text = prefix + para.runs[0].text`。
python-docx 1.2.0 的 `Run.text` setter 会先 `clear_content()`（移除该 run 除 `w:rPr` 外的**所有**子元素）再按字符串重建。因此该 run 内的以下内容会被**物理删除**：内联图片 `w:drawing`、脚注/尾注引用 `w:footnoteReference/w:endnoteReference`、域代码 `w:fldChar/w:instrText`（Zotero/Mendeley/EndNote 域、交叉引用、自动编号域）、`w:object`、`w:sym`、`w:softHyphen`，以及**分页符 `w:br w:type="page"`**（getter 忽略它，连"逐字比对"都检测不到）。最小复现思路：造一份 docx，References 下条目首个 run 内含 ① 一张内联小图，或 ② 一个脚注引用，或 ③ 一个分页符；以 `reference_numbered=True`（IEEE）调用 `fix_docx` → 输出文档中该图/脚注标记/分页消失，脚注正文成为孤儿。【修复建议】不要用 `run.text` 整体重写；改为在首个 run 的**第一个 `w:t` 之前插入** `w:t`（或新建一个 run 并 `w:r.insert(0, new_r)` 到段落第一个子元素位置），保留 run 内其它子元素；并补一条"含图片/脚注/域的参考文献条目"回归测试。

**【严重程度：高（真 bug）】【`src/engine/fixer.py:177-185`】** 当条目的可见文字全部位于 `w:hyperlink` 内（DOI/URL 链接型参考文献，python-docx 导出的 `Paragraph.runs` 只含直接 `w:r` 子元素，而 `Paragraph.text` 含 `w:hyperlink`）时 `para.runs == []`，于是走 `para.add_run(prefix)` → `w:r` 被追加到段末，文本变成 `"Author... Title. [1] "`：编号跑到末尾、条目被改写。最小复现：参考条目为一个超链接（或段落里只有超链接 + 无直接 run），IEEE 编号修正 → 段末多出 `[1] `。【修复】改为在段落 XML 首位插入新 run（`para._p.insert(0, r)`），或先备份原 `w:hyperlink` 再有序插入；同时用"插入后 `para.text` 仍以 `[n] ` 开头"做断言。

**【严重程度：中（需求冲突，非崩溃）】【`src/engine/fixer.py:180-186`】** `[n] ` 前缀本身就是"新增文字"，与 NFR6「修正前后文档文字内容**逐字一致**」字面冲突（FR5 把"编号标记"列为可改项，语义有歧义）。更关键的是：**NFR6 要求的"自动化断言守护"并未覆盖这条唯一改文字的路径** —— `tests/test_fixer.py:47-51` 的 `_body_texts` 显式排除参考文献段，`131-157` 只断言"关键 token 仍包含"，因此"逐字一致"在参考文献维度从未被验证，测试全绿 ≠ NFR6 成立。【修复】二选一：(a) 把参考文献编号也纳入逐字一致断言（即取消 `[n]` 前缀改写，编号交给 Word 列表编号，纯格式层）；(b) 在规格中显式豁免"编号标记"并单独断言"仅新增前缀、其余字符序列完全一致 + 条目数不变"。

**【严重程度：低（真 bug，影响小）】【`src/engine/fixer.py:174-186`】** `n` 对已带 `[n]` 的条目也自增，混合编号（部分条目自带 `[3]`、部分无编号）时可能产生重复编号。【修复】按已存在编号解析最大/下一个可用序号，或只对"全部无编号"的列表补齐。

### B10 · 标题保护是否完备 —— **维度上完备，识别上不完备**

- 已核实：对**被识别为标题**的段落，六个维度全部跳过——字体/字号/行距/对齐用 `no_heading`（`fixer.py:54-63,239-243`），首行缩进/段后间距用 `body_only`（`243-248`），页边距属分节属性、不涉及标题。`fixer.py:259` 的 `headings_preserved=True` 是硬编码、无实际校验（nit）。
- 【严重程度：高（真风险）】【`src/engine/docx_reader.py:128-160`】Word 内置 **`Title` / `Subtitle` / `Title 1`** 等样式既不含关键字"heading"也不带编号数字、且通常无 `outlineLvl` → `is_heading` 返回 False。论文主标题（常 24–28pt）会被改成正文 12pt，正是 FR5 要禁止的"视觉层级崩塌"。最小复现：`doc.add_paragraph("论文标题", style="Title")` + `TargetProfile(page={"font_family":"Arial","font_size_pt":11})` → 修正后标题变 Arial 11pt。【修复】标题判定加入 `title/subtitle/标题/题目` 及基于"样式名 + 大字号 + 加粗"的伪标题启发式，或改为白名单式"仅对确认的正文样式生效"。
- 【严重程度：中（真风险）】【`src/engine/docx_reader.py:35,128-139`】本地化标题名只覆盖 `heading / título / titre / überschrift / 标题`；荷兰语 `Kop`、波兰语 `Nagłówek`、俄语 `Заголовок`、日语 `見出し`、韩语 `제목` 等一律漏判。
- 【严重程度：中（真风险）】【`src/engine/docx_reader.py:142-155`】`outlineLvl` 只读段落直接 `pPr`，不读**样式定义**中的 `outlineLvl`；自定义标题样式（名为 `MyTitle`、样式层设 `outlineLvl=0`）漏判，既漏保护也漏统计（`heading_levels_used` 偏低）。
- 【严重程度：中（真风险）】【`src/engine/fixer.py:97`】`doc.styles["Normal"].font.name = font` 会"间接"改字体：Word 中 Heading/Title/表格/页眉/脚注等大量样式 `basedOn Normal` 且不显式设字体名。作者已意识到字号继承风险（85-90 行注释）而只设字体名，但**字体名的继承同样会把标题字体改掉**，与"标题字体不得改动"冲突。最小复现：自定义一个基于 Normal、未显式设字体的标题样式（带 outlineLvl），修正后其字体随正文变化。【修复】不修改 Normal 字体名，改为按目标正文段落（明确的正文字体维度）设置；或同时把标题样式显式"钉住"为原字体/字号。
- 伪标题（无任何样式标记、仅大字号/加粗的手写标题）：**中（真风险）**，规格第 10 条已点名，当前无任何保护。
- 表格/文本框内标题：格式不变（保护成立），但也意味着"检查采样含表格、修正不覆盖表格"的不对称（见 B16/补充项）。

### B11 · 授权门禁一致性与可绕过性 —— **共享门禁成立；计数时机有一处真风险**

- 已核实无问题：
  - CLI `src/main.py:187-198` 与 GUI `src/ui/app.py:232-235` 都调用同一 `license.require_fix_entitlement()`；成功后 CLI `main.py:235`、GUI `ui/app.py:261` 调用同一 `record_fix_used()`。**不存在只在一个入口生效的门禁，也没有绕过路径。**
  - **预览不消耗试用**：CLI `main.py:206-220`、GUI `ui/app.py:236-257` 的预览路径都不会调用 `record_fix_used`；`fix_docx` 抛 `DocxReadError/ValueError` 的失败路径也不计次（计数在 `fix_docx` 返回之后）。
  - 拒绝分支退出码 3（`main.py:198`），JSON 分支含 `ok:false/reason/message`，符合 FR6。
- 【严重程度：中（真风险）】【`src/main.py:235` / `src/ui/app.py:261`】`lic.record_fix_used()` 无异常保护：`save_state`（`license.py:90-93`）可能因家目录只读/磁盘满/权限失败，此时**修正稿已落盘但试用未计次**，CLI 抛 traceback、GUI 弹出 Tk 异常；用户可重复再跑，等于白嫖。属审查重点 11 中"改了但没计数"的实例。【修复】`record_fix_used` 内部 try/except 并记录"待上报"标记（配合 B12 的服务端校验），或把计数改为落盘前用原子占位文件预留额度。
- 【严重程度：低（设计问题）】试用耗尽时 `fix --preview` 同样返回 rc=3，而预览不落盘、不消耗额度；FR5 未要求对预览设门禁。【修复】明确产品决策并写进规格；若"预览=信息展示"，应允许无门禁预览。
- 【严重程度：低】门禁结果 `gate` 只裁决一次并缓存复用；同机两个实例并发可各跑一次 fix（都读到"试用未用"），存在多消耗窗口。
- 【严重程度：低】文档已合规（预览输出"无需修正"）时仍会生成副本并消耗试用（`main.py:227-235` 不检查 `changes` 是否为空）。

### B12 · 试用防重置 —— **不成立（FR6「重装不能绕过」目前做不到）**

- 【严重程度：高（真 bug/真风险）】【`src/license/license.py:77-87,152-184`】本地状态是唯一判据：删除 `~/.thesis-format-doctor-global/license.json`（或 JSON 损坏触发 `83-84` 的 `s = {}` 回落默认值）→ `_default_state()` → `trial_fix_used=False` → 再次放行；`require_fix_entitlement` 从不查询中台，`_notify_mid_fix_used` 只写不读（"最佳努力"）。离线/在线都一样可绕过，FR6 明文"使用户卸载重装无法再次白嫖"不成立。最小复现：`fix` 一次 → 删除 license.json → `fix` 再次被允许（rc=0）。【修复】`require_fix_entitlement` 在线时先向中台查询该机器指纹的 trial/激活状态（服务器明确 `revoked/expired/trial_used` 才锁，失败按离线宽限），或改为服务端一次性令牌驱动。
- 【严重程度：高/中（真 bug）】【`src/license/license.py:85-87,152-171`】本地状态**没有做机器校验**：`load_state` 用文件里的 `machine` 覆盖当前值，`require_fix_entitlement` 从不比较 `state["machine"] == _machine_fingerprint()` → 把其他机器的 `license.json` 拷过来即可解锁/复用试用（FR6"本地状态也做机器绑定"未真正生效）；`status()`（187-196）还回显文件中的旧指纹。【修复】load 时校验指纹，不匹配则视为未激活并告警。
- 【严重程度：中】【`src/license/license.py:60-63`】指纹稳定性/碰撞：`platform.node()` 在改机名、容器、VPN/动态域名场景会变 → 试用重置；`node()` 为空时 `system||machine` 在多机间可能碰撞 → 共享试用额度；截断 16 hex（64bit）碰撞概率低但非零。
- 【严重程度：中】【`src/license/license.py:112-124,174-184`】离线时的那次试用登记永远不会补报（fire-and-forget、无重试/队列），之后联网也不会同步，形成稳定绕过窗口。
- 【严重程度：中】【`src/license/license.py:37,131-149,152-171`】docstring 承诺的"心跳 3 天、服务器返回 revoked/expired 才锁"**未实现**：`HEARTBEAT_DAYS` 是死常量，`activate` 不看 `revoked/expired`，`require_fix_entitlement` 也不看 `last_heartbeat`。被吊销的码永久有效，FR8"服务器明确返回 revoked/expired 时锁"没有落地。
- 【严重程度：中】【`src/license/license.py:103-108`】激活码放在 GET query 里，会进入中台/反向代理访问日志；建议改 POST body（另：GET 参数也可能被 CDN 缓存）。
- 【严重程度：低】【`src/license/license.py:90-93`】状态文件非原子写（无 tmp+rename），写一半崩溃即损坏 → 按上面第 1 条回落为"未用试用"。

### B13 · JSON 输出编码一致性 —— **逐分支核对：3 条 ✅，1 条 ❌**

逐分支结论（`src/main.py`）：

| 分支 | 位置 | 结论 |
|---|---|---|
| `check --json` | `main.py:105-113` → `_write_json`(108) | ✅ UTF-8 字节流 |
| `check --json --out` 写文件 | `main.py:109-112` (`encoding="utf-8"`) | ✅ |
| `fix` 拒绝分支 `--json` | `main.py:190-194` → `_write_json`(192) | ✅ |
| `fix --preview --json` | `main.py:212-215` → **`print(json.dumps(...))`(213)** | ❌ 未走字节流 |
| `fix` 成功 `--json` | `main.py:237-240` → `_write_json`(238) | ✅ |
| 人类可读输出（`check` 非 json / `activate` / `status` / `export-schema` / 提示语） | 多处 `print` | ✅（非 JSON，允许走终端编码，`_init_console` 兜底不崩） |

- 【严重程度：中（真 bug）】【`src/main.py:213`】`fix --preview --json` 用普通 `print`，中文 Windows **交互式 GBK 终端**下经 `_init_console` 的 `errors="replace"`（`main.py:317-324`）会把 `→`/`″`/emoji 等非 GBK 字符降级成 `?`，内容被篡改，与其它 JSON 分支不一致，违反 FR7"任意终端下都是合法可解析的 UTF-8"。注意：`docs/打磨轮_完成记录_2026-09-11.md` 第五节里该分支是在**管道**（subprocess 捕获）下验证的，而管道会被强制 UTF-8，恰好掩盖了这个分支的问题。最小复现：在中文交互式 cmd/PowerShell（GBK，非管道）执行 `python -m src.main fix x.docx --preview --json` → stdout 中出现 `?`。【修复】改用 `_write_json`。
- 【严重程度：低】【`src/main.py:59`】`_write_json` 的兜底 `print(text)` 在无 `sys.stdout.buffer` 的自定义流下仍走终端编码；建议改为 `sys.stdout.write` 前先 `reconfigure`，或直接抛出明确错误。
- 已核实无问题：`_write_json` 本身写 `sys.stdout.buffer` + `utf-8` + `\n` + flush，绕开编码层，正确。

### B14 · 终端健壮性 —— **基本成立**

- 已核实：`_init_console()`（`main.py:310-324`）在 `main()`（327-334）一开始调用，覆盖所有 CLI 入口（`python -m src.main ...`、`gui` 子命令、以 `main(argv)` 调用）；非 tty 强制 `utf-8`，tty 保留原编码 + `errors="replace"`；`sys.stdout/stderr` 为 `None`（pythonw）时 `isatty()` 抛 AttributeError 被 `except` 吞掉，不会崩。报告里的 ✅/❌/⚠️/ℹ️ 在 GBK 终端不再触发 `UnicodeEncodeError`。
- 【严重程度：低】【`src/main.py:92-120,181-247`】以库方式**直接调用 `cmd_check/cmd_fix`**（不经 `main()`）不会初始化控制台，非 tty 下也不会强制 UTF-8，仍可能因 emoji 抛 `UnicodeEncodeError`；`tests/test_polish.py:29` 就是直接 import `report_to_dict` 的用法，说明库式调用被鼓励。【修复】把 `_init_console()` 移到 `cmd_check/cmd_fix` 入口或模块导入时做一次幂等初始化。
- 英文/非中文控制台把报告中文显示为 `?`：`docs/打磨轮_完成记录` 已列为已知边界，本轮不算缺陷。

### B15 · LaTeX 解析正确性与边界

- 【严重程度：中/高（真 bug）】【`src/engine/latex_template.py:74-148,155-196`】**完全不处理注释**：所有抽取都在原始文本上跑正则，`% \doublespacing`、`% \bibliographystyle{apacite}`、`% \documentclass[12pt]{article}` 都会被采信，给模板打上假行距/假引用风格。最小复现：`parse_latex_source("% \\doublespacing\n\\documentclass{article}")` → `page["line_spacing"]==2.0`；`% \bibliographystyle{IEEEtran}` → `reference_style=="IEEE"`。【修复】先按行剥离未转义的 `%`（注意 `\%` 与 `\\%`）再解析，并补注释专用测试。
- 【严重程度：中（真 bug，来源标注错误）】【`latex_template.py:188-196` + `questionnaire.py:311-321`】`parse_latex_source` 设置了 `reference_style/reference_hanging_indent_in/reference_numbered`，但 `resolution` 里**没有写这三项**；`_overlay_one` 用 `.get(..., "school")` 兜底 → LaTeX 模板提供的引用维度在报告里被标成"学校模板"（FR2/FR4 的来源标注不准确）。最小复现：`build_target("APA", latex=含 \bibliographystyle{IEEEtran} 的文件)` → 报告 `resolution` 中 `reference_style = 学校模板`。【修复】在 `parse_latex_source` 里补 `resolution["reference_style"/"reference_hanging_indent_in"/"reference_numbered"] = SOURCE_LABEL`；并让 `_overlay_one` 的默认值不再硬编码 "school"。
- 【严重程度：低】【`src/engine/report.py:17-24`】`ORIGIN_LABEL` 缺 `"latex_template"`（也缺 `"latex"`）→ 页面维度来源在报告里显示原始英文串而非中文标签；`latex_template.SOURCE_LABEL="latex_template"` 与 `build_target` 的来源键 `"latex"`、docstring 用词不统一。
- 【严重程度：低（覆盖度/文档）】只取第一个 `\documentclass`/`\geometry`；选项含嵌套 `{}` 会截断；`\documentclass{IEEEtran}`（类名暗示风格）不产生引用风格；`_BIB_STYLE_MAP`（30-52）无 `acm*` 映射，而 docstring 声称覆盖 ACM/revtex —— "~80% 覆盖"略有夸大。
- 【严重程度：中/低（需求偏差）】FR2/§8 说"取不到的值**回落规范并提示用户**"，但目前 LaTeX 未识别到任何字段时只是静默回落，报告/CLI 无任何提示（`parse_latex_source` 空 `page`，`build_target` 不产生 finding）。建议在报告里加一条"LaTeX 模板未识别到字段，已回落规范"的 info。
- 未支持 `\setlength{\baselinestretch}`、多行选项、无名单位：属 §8 已知上限，可接受。

### B16 · 测试覆盖盲区 —— **列出真实用户会踩但没测的路径**

已核实测试守护到的：检查新维度/降级、修正新维度、标题（仅 Heading 样式）不被压、输出 `-2/-3`、预览只列变化、表格采样、正文中 "References" 误判、报告 dict 字段、choice 归一化、授权 6 条路径、LaTeX 9 条。

盲区（按风险排序）：
- **内容零改动**：`tests/test_fixer.py:47-51,71` 排除了参考文献段；`131-157` 只断言 token 包含。→ 完全没覆盖 B9 的 `[n]` 前缀、`run.text` 破坏 run 内容、hyperlink-only 条目错位。**（高）**
- **GUI 零测试**（`src/ui/app.py` 无任何测试）：LaTeX 缺失、输出路径去重、门禁时序、异常处理都无守护。**（中）**
- **授权防重置**：无"状态文件删除/损坏/machine 不匹配/跨机拷贝/离线登记后补报/revoked-expired"用例。**（高，对应 B12）**
- **CLI 端到端**：无 `--json` 各分支在 GBK 终端下的字节输出断言、无 rc=3/rc=2、无 `--out` 与输入同路径（大小写/相对路径）用例 —— B13 与"覆盖原件"缺陷本可被这类测试拦住。**（中/高）**
- **标题保护反例**：无 `Title` 样式、伪标题（大字号加粗）、本地化样式名、样式级 outlineLvl、Normal 继承用例。**（中）**
- **读取健壮性**：无"Document() 成功但内部 XML 坏"、`w:documentProtection` 受保护文档、超大文档、Unicode/中文文件名文档用例；`src/engine/docx_reader.py` 的三级回退也缺少"仅样式层/仅 Normal 层设值"的用例。**（中）**
- **修正对称性**：无"检查采样含表格 → 修正是否覆盖表格文字"的用例（当前不覆盖）。**（中）**
- **五通道合并**：只测了 template 通道的冲突；无 latex/ai/questionnaire 叠加优先级、引用维度冲突用例。**（中）**
- nit：`tests/test_latex_template.py:10` `import tempfile` 未使用。

---

## 三、B9 专项 · fixer.py 写入路径完整清单与判定

`src/engine/fixer.py` 中所有会写文档对象/磁盘的语句（穷举）：

| # | 行号 | 写入内容 | 是否改文字 | 判定 |
|---|---|---|---|---|
| 1 | `fixer.py:82` | `setattr(sec, attr, val)` 各分节四边页边距 | 否 | ✅ 格式 |
| 2 | `fixer.py:97` | `doc.styles["Normal"].font.name = font` | 否 | ⚠️ 格式，但影响范围外溢（含标题/表格/页眉等继承 Normal 的内容，见 B10-e） |
| 3 | `fixer.py:103` | `run.font.name = font` | 否 | ✅ 格式 |
| 4 | `fixer.py:106` | `run.font.size = Pt(...)` | 否 | ✅ 格式 |
| 5 | `fixer.py:121` | `paragraph_format.line_spacing` | 否 | ✅ 格式 |
| 6 | `fixer.py:132` | `paragraph_format.alignment` | 否 | ✅ 格式 |
| 7 | `fixer.py:144` | `paragraph_format.first_line_indent` | 否 | ✅ 格式 |
| 8 | `fixer.py:146` | `paragraph_format.left_indent = 0` | 否 | ✅ 格式（会覆盖正文段原有的左缩进，如块引用；属格式层，需产品确认） |
| 9 | `fixer.py:154` | `paragraph_format.space_after` | 否 | ✅ 格式 |
| 10 | `fixer.py:171` | 参考文献 `left_indent` | 否 | ✅ 格式 |
| 11 | `fixer.py:172` | 参考文献 `first_line_indent` | 否 | ✅ 格式 |
| 12 | **`fixer.py:183`** | **`para.runs[0].text = prefix + para.runs[0].text`** | **是** | ❌ **真 bug（B9-a）：setter 先 `clear_content()`，删除该 run 内图片/脚注引用/域/分页符等非文本内容** |
| 13 | **`fixer.py:185`** | **`para.add_run(prefix)`** | **是** | ❌ **真 bug（B9-b）：hyperlink-only 条目会把 `[n] ` 追加到段末，文字错位** |
| 14 | `fixer.py:261` | `doc.save(output_path)` | 否（落盘） | ✅ 有新文件保护，但见下方"覆盖原件"缺陷 |

**穷举核查结论**：`fixer.py` 中**不存在** `paragraph.text = ...`、`add_paragraph`、`add_table`、`insert_*`、`delete_*`、`clear()`、页眉/页脚/脚注/文本框写入、参考文献条目重排或拼接，也没有任何水印/试用标记写入 —— 除第 12/13 行外，全部为纯格式属性写入，符合 FR5"绝不新增/删除/改写正文、不得写水印"的部分；**但第 12/13 行使 NFR6 在"参考文献编号制"场景下失守**，且现有自动化断言不覆盖这两条路径。

补充（写入范围与检查范围不对称，**中**）：`read_docx` 把表格文字纳入字体/字号/行距采样（`docx_reader.py:407-417`），而 `fixer` 只处理 `doc.paragraphs` 顶层段落（`fixer.py:57,230-248`），**不改表格内文字格式** → 会出现"检查说正文字体不对 → 一键修正后报告仍报不对"的死循环（尤其以表格为主的论文）。FR5 未明文要求改表格，但按 FR3/FR5 的产品意图应一致；至少要给用户提示"表格内文字未修正"。

---

## 四、补充 · 16 条之外发现的真问题（同样按格式）

1. 【严重程度：高（真 bug，数据丢失）】【`src/main.py:222`、`src/engine/fixer.py:211`】"输出不得等于输入"只用**字符串相等**判断，未做路径规范化：Windows 大小写不敏感文件系统（`--out THESIS.DOCX` vs 输入 `thesis.docx`）、相对/绝对混写（`fix paper.docx --out .\paper.docx`）都会绕过检查，`doc.save` 覆盖原件（原件被改后的内容替换），违反 FR5"绝不覆盖原件"。最小复现：`python -m src.main fix paper.docx --out .\paper.docx`。GUI 因 `unique_output_path` 恰好不受影响，CLI 受影响。【修复】用 `os.path.realpath` + `os.path.normcase`（Windows）比较，输出已存在时再用 `os.path.samefile` 兜底。
2. 【严重程度：中（真 bug）】【`src/main.py:201-205`】`cmd_fix` 的 `_build_target` try 只捕获 `ValueError/FileNotFoundError/JSONDecodeError`，未捕获 `DocxReadError` → `fix --template 损坏.docx` 会抛 traceback（rc≠2）而不是友好报错；`cmd_check`（96-104）是正确的对照。【修复】补 `except DocxReadError`。
3. 【严重程度：中（真 bug，FR4）】【`src/engine/checker.py:131-154`】参考文献 Finding 从不设置 `source`（`Finding(...)` 无 source 参数），`--json` 里 `findings[].source` 为空串，报告无法标注"该结论以规范为准"，与 FR4"并标注该结论来源"不符。【修复】传入 `source="spec"`（或目标 `resolution` 中的对应来源）。
4. 【严重程度：中（体验/FR5）】【`src/engine/fixer.py:331-338`】`preview_fix` 只要 `target.reference_style` 存在就无条件追加一行"参考文献…"，即使文档悬挂缩进/编码已完全符合（`cur_hang == 目标值`）也照列（显示 `0.5″ → 0.5″`），违反"只列出实际会变的项"。【修复】比照悬挂缩进/编码现状决定是否输出，或输出"已符合"。
5. 【严重程度：低（真 bug）】【`src/main.py:227-234`】`fix_docx` 的 `doc.save` 异常（输出目录不存在/无写权限）与 `_write_json` 的 `--out` 写文件异常都未被捕获，会打 traceback。【修复】补 `except OSError` 并返回 rc=2。

---

## 五、总体评价

v2 的工程完成度明显高于一般"边做边改"的项目：规格与代码映射清晰，五通道收敛、spec 页面维度降级 info、引用维度硬判、参考文献段识别单一实现、三级回退、固定行距防误判、输出防覆盖、JSON 结构化输出等主线需求都真正落了地，A 组的隔离性与离线边界经逐项检索确认无问题，代码可读性与单一真相意识很好。但**产品最致命的两处承诺目前是"文字上成立、实现上不成立"**：其一，"内容零改动"在参考文献编号制下会写文字，且 `run.text` 整体重写 + 段末 `add_run` 的实现方式在"含图片/脚注/域/分页符的条目"和"超链接型条目"上会**丢失或错位文献内容**，而唯一的自动化守护恰好把参考文献段排除在外——测试全绿具有误导性；其二，"试用防重置/机器绑定"实际上是**只写不读的本地状态 + 无机器校验**，删除或损坏一个 json 即可重置试用、拷贝该文件即可跨机复用，中台登记形同日志，FR6 的表述与实现不符。此外还有若干可被真实用户直接触发的缺陷：CLI 显式 `--out` 用相对/大小写变体可覆盖原件（数据丢失）、本地化/`Title`/伪标题不受保护、`fix --template` 坏文件崩溃、`preview --json` 编码不一致。这些问题不涉及架构返工，但 P0 三项必须在任何真实用户拿它处理论文之前修掉——尤其"覆盖原件"和"内容丢失"是可能造成不可逆后果的类别。

## 六、修改清单

**P0 · 必修（会导致内容/数据不可逆损失或核心承诺失效）**
- `src/engine/fixer.py:183,185`：禁止用 `run.text` 整体重写、禁止无条件 `add_run`；改为"在段落/第一个 run 的首位插入纯文本 run"，保留图片/脚注/域/分页符，并对 hyperlink-only 条目单独处理；补"含图片/脚注/域/超链接条目"的零改动回归测试（含参考文献段的**逐字**断言）。
- `src/main.py:222` + `src/engine/fixer.py:211`：输出路径用 `realpath/normcase/samefile` 判定，杜绝任何形式的原件覆盖。
- `src/license/license.py:77-87,152-171`：`require_fix_entitlement` 增加（a）本地 `machine` 与当前指纹一致性校验、（b）在线时向中台查询该指纹的 trial/激活状态（失败再离线宽限）；状态文件损坏按"已用"fail-closed，并改原子写。
- `src/main.py:235` + `src/ui/app.py:261`：`record_fix_used()` 包 try/except，避免"已生成但未计次"。

**P1 · 建议（功能正确性/需求一致性）**
- 标题保护：加入 `Title/Subtitle/标题` 及"大字号/加粗伪标题"判定、样式级 `outlineLvl`、更多本地化样式名；去掉对 `Normal` 字体名的修改（`fixer.py:97`）。
- `src/engine/questionnaire.py:342-356`：修正五通道叠加优先级为 FR2 的 `template > ai > questionnaire > latex`；为引用维度补冲突登记（`311-321`）。
- `src/engine/latex_template.py`：剥离 `%` 注释；补引用维度 `resolution`（配合 `questionnaire.py:311-321` 去掉硬编码 `"school"` 默认）。
- `src/main.py:213`：`preview --json` 改用 `_write_json`；`main.py:201-205` 补 `except DocxReadError`。
- `src/engine/checker.py:131-154`：参考文献 Finding 补 `source`；`docx_reader.py:405` 字体/字号统计排除标题段；`docx_reader.py:378` 补 `w:eastAsia` 字体读取；`docx_reader.py:283-329` 读取体整体包错。
- `src/ui/app.py:200-206`：GUI 增加 LaTeX 通道（FR2-⑤/FR7）。
- `src/engine/fixer.py:331-338`：预览参考文献行按实际差异输出。
- 明确"预览是否受门禁约束""无变化修正是计次"两个产品决策并写入规格。

**P2 · 可选（体验/维护性/nit）**
- 消除 `ALIGN_LABEL`/margin 映射重复，清理 `main.py`/`ui/app.py` 死导入与 `HEARTBEAT_DAYS` 死常量，统一 `src/__init__.__version__` 为 v2。
- `checker.py:21-22` 澄清 `MARGIN_TOL_WARN < FAIL`；`report.py:17-24` 补 `latex_template` 中文标签；`license.py:103-108` 激活码改 POST；`main.py:59` 兜底打印改为显式失败；GUI 对用户显式选择路径不再强制加序号；修正无变化时提示而非消耗试用；表格内文字修正能力（或明确提示不修正）；补充 GUI 测试、CLI 端到端 GBK 字节断言、损坏/中文文件名/超大文档用例；`tests/test_latex_template.py:10` 清理未用 import。

---

## 附录 A · 独立复核状态（2026-09-11，由芦苇逐条核验源码后补记）

> 说明：本报告由 codex 静态只读审查产出。以下为**复核人亲自打开源码逐行核验**的结论，
> 用于区分「codex 的判断」与「已被二次证实的结论」。核验方式：直接读取被指行号及其上下文源码。
> 本次复核未修改任何代码（修复另行安排）。

| 编号 | 结论 | 核验证据（亲眼所见） |
|---|---|---|
| **P0-1** `run.text` 整体重写会丢内容 | ✅ **成立** | `fixer.py:183` 确为 `para.runs[0].text = prefix + para.runs[0].text`。python-docx `Run.text` setter 先 `clear_content()` 清除该 run 内除 `w:rPr` 外的**全部**子元素 → 内联图片 / 脚注引用 / 域代码 / 分页符被物理删除；且 `Paragraph.text` getter 忽略分页符，**逐字比对也检测不到**。 |
| **P0-1b** hyperlink-only 条目编号错位 | ✅ **成立** | `fixer.py:184-185` 确为 `para.add_run(prefix)`。`Paragraph.runs` 只含直接 `w:r`，不含 `w:hyperlink` 内的 run，故 DOI/URL 型条目 `runs == []` → 新 run 追加至**段末**，编号跑位。 |
| **P0-1c** 测试对参考文献段无逐字断言 | ✅ **成立（关键）** | `tests/test_fixer.py:49-51` 的 `_body_texts` 用 `ref_idx` 集合**显式排除参考文献段**；`131-157` 仅断言 token「包含」。→ 「修正前后逐字一致」在参考文献维度**从未被验证**，全绿存在误导性。 |
| **P0-2** 输出路径仅字符串比较可覆盖原件 | ✅ **成立** | `main.py:222` 为 `if out == args.docx`；`fixer.py:217` 同。均未做 `realpath`/`normcase`/`samefile` 规范化 → Windows 大小写不敏感与相对/绝对混写可绕过，`doc.save` 覆盖原件。 |
| **P0-3** 试用防重置不成立 | ✅ **成立** | `license.py:77-87` `load_state` 返回文件内容且**从不校验** `machine`；`require_fix_entitlement`（152-171）只读本地 `activated`/`trial_fix_used`，**从不查询中台**；`_notify_mid_fix_used`（112-124）只写不读。删除 `license.json` 即回到"试用未用"。 |
| **P0-4** `record_fix_used` 无异常保护 | ✅ **成立** | `main.py:235`、`ui/app.py:261` 直调，未包 try/except；`save_state`（license.py:90-93）为非原子写，失败时"已生成修正稿但未计次"。 |
| **B10** `Title`/伪标题漏保护 | ✅ **成立** | `docx_reader.py:128-139` 要求样式名**以 heading 类关键字开头且随后带数字**；`Title`/`Subtitle` 均不匹配 → `is_heading()` 返回 False → 主标题（24–28pt）会被压成正文 12pt。 |
| **B10b** `outlineLvl` 只读段落层 | ✅ **成立** | `docx_reader.py:147-152` 仅查 `para._p.pPr`，未读**样式定义**中的 `outlineLvl` → 自定义标题样式漏判。 |
| **B10c** 改 `Normal.font.name` 外溢 | ✅ **成立** | `fixer.py:97` 设置 Normal 样式字体名，`basedOn Normal` 的标题/表格/页眉样式会**间接**跟随变化。 |
| **B11** `record_fix_used` 时机 | ✅ **成立** | 计数在 `fix_docx` 返回之后（`main.py:227-235`），失败路径不计次（这点是对的），但保存状态失败时会漏计。 |
| **B13** `fix --preview --json` 编码不一致 | ✅ **成立** | `main.py:213` 确为 `print(json.dumps(...))`，而 `check --json`(108) / 拒绝分支(192) / fix 成功(238) 均走 `_write_json`。**本条在审查派出前已由复核人独立发现**，与 codex 结论一致 → 交叉验证成立。 |
| **B15a** LaTeX 不剥离注释 | ✅ **成立** | `latex_template.py:155-196` 的 `parse_latex_source` 直接在原始文本跑正则，全文件无 `%` 剥离逻辑 → `% \doublespacing`、`% \bibliographystyle{IEEEtran}` 会被采信为真实要求。 |
| **B15b** LaTeX 引用维度 `resolution` 缺失 | ✅ **成立** | `latex_template.py:158-178` 只写 `font_size_pt / font_family / margin_* / line_spacing` 四类，**未写** `reference_style / reference_hanging_indent_in / reference_numbered`；`questionnaire.py:313-321` 用 `.get(key, "school")` 兜底 → 报告误标"学校模板"。 |
| **A3** `MARGIN_TOL_WARN < FAIL` 疑笔误 | ✅ **成立** | `checker.py:21-22` 为 `FAIL=0.06 / WARN=0.05`，warn 带仅 0.01″；对比同类常量（SIZE 0.3<0.5、SPACING 0.1<0.15、INDENT 0.06<0.12）可判为笔误。 |
| **A4-1** 五通道叠加优先级与规格相反 | ✅ **成立** | `questionnaire.py:342-356` extras 顺序为 `template → ai → questionnaire → latex`，`_overlay_one` 为**后写覆盖先写** → 实际优先级 `latex > questionnaire > ai > template`，与 FR2 明文 `template > ai > questionnaire > latex` 相反。 |
| **A4-2** 引用维度不登记冲突 | ✅ **成立** | `questionnaire.py:298-309` 只对 page/heading 调 `_register`；`311-321` 的 reference 三分支**不登记** → 引用维度冲突静默丢失。 |
| **A6** 字体字号采样含标题段 | ✅ **成立** | `docx_reader.py:395-405` 对**全部** `paras` 调 `_sample_paragraph`（仅行距限 `body_paras`）→ 与 366 行注释「正文段落」语义不符。 |
| **补充1** `fix --template 坏文件` 崩溃 | ✅ **成立** | `main.py:201-205` 的 except 仅含 `ValueError/FileNotFoundError/JSONDecodeError`，缺 `DocxReadError`（`cmd_check:99` 有，形成对照）。 |

**复核结论**：本报告 **P0 四项 + P1 主要项全部经源码复核成立，无一条误报**；未发现需驳回项。
其中 **P0-1（可能丢图/脚注）、P0-2（可能覆盖原件）** 属"对用户不可逆"级别，建议在任何真实用户使用前修复。

**尚未复核（留待修复时验证）**：B12 的指纹稳定性/碰撞、B15 覆盖率夸大、B16 各测试盲区清单、P2 全部 nit 项 —— 这些不影响本报告主要结论成立。

---

## 七、修复状态（2026-09-11 全量修复后补记）

> 老板指令：**全量 P0 + P1 一起修**。下表为逐条落地情况与**验证证据**。
> 验证方式：`pytest` 全量（`tests/`，含本轮新增回归）+ CLI 端到端 12 步冒烟（真实 subprocess、真实 docx、临时 license 文件）。
> 最终结果：**96 passed，0 failed**；端到端 12 步全部符合预期。

### 7.1 P0 · 必修（4/4 已修）

| 编号 | 问题 | 修复方式 | 验证 |
|---|---|---|---|
| **P0-1** | `run.text` 整体重写丢图片/脚注/域/分页符 | `fixer.py` 改为 `_insert_prefix_run()`：新建纯文本 run 后**搬移到 `w:pPr` 之后的首位**，绝不触碰已有 run 的子元素 | `test_review_fixes.py`：含内联图片/分页符/域代码/超链接的参考文献条目，修正后**逐字零改动**（仅允许新增 `[n] ` 前缀） |
| **P0-1b** | hyperlink-only 条目编号跑到段末 | 同上，插入位置固定在段落头部（不依赖 `para.runs` 是否存在） | 同上用例覆盖 `runs == []` 的超链接型条目 |
| **P0-1c** | 参考文献段无逐字断言（测试误导性） | 新增**参考文献段逐字比对**回归（`_p_xml` 级） | 新增 18 项回归中的第一条即为该断言 |
| **P0-2** | 输出路径仅字符串比较可覆盖原件 | `fixer.same_file()`：`realpath + normcase + samefile` 三重判定；CLI/GUI 共用 | 大小写变体（`sample.docx` vs `SAMPLE.DOCX`）、同路径直传均 rc=2 且**原件字节不变** |
| **P0-3** | 试用防重置不成立（只写不读 / 无机器校验） | `license.py` 重写：① 本地 `machine` 与当前指纹比对（不匹配→不认激活，在线可重绑）；② `require_fix_entitlement` 在线查中台 `trial` 状态（服务器明确 `trial_used` 才锁，失败离线宽限）；③ 状态文件损坏 **fail-closed**；④ 原子写（tmp + `os.replace`） | `test_license.py` 6 项：损坏 fail-closed、异机不认、心跳重绑、服务器 revoked/trial_used 锁、record 不抛、路由共识 |
| **P0-4** | `record_fix_used` 无异常保护 | 包 try/except，**绝不抛异常**，返回是否落盘；CLI/GUI 据此提示 | 同上测试覆盖；CLI 端到端第 4 步验证计次成功 |

> P0-3 补充：中台真实路由与海外版原调用**不符**（原 `/api/auth/verify`、`/api/auth/record` 在中台不存在，激活从未通过）。已改为中台真实契约 `POST /api/activate`、`/api/heartbeat`、`/api/trial`。
> **`/api/trial` 端点需在共享中台部署后才生效**（补丁见 `server_patch/trial_endpoint.js` + `README.md`，不改动共享中台源码）；未部署时按「服务器不可达」走离线宽限，本地 fail-closed 兜底，功能不受影响。**残余风险**：离线删状态文件仍可重置，已在文档标注。

### 7.2 P1 · 建议（已修）

| 条目 | 修复方式 |
|---|---|
| 标题保护（`Title`/伪标题/本地化名/样式级 outlineLvl） | `docx_reader` 新增 `_TITLE_STYLE_KEYWORDS`、`_style_outline_level()`、`is_pseudo_heading()`（短+整段加粗+≥13pt），`is_heading()` 综合判定；样式关键字扩至多语种 |
| 不再改 `Normal` 字体名 | `_apply_font_and_size` 改为**逐 run** 设字体，并同时写 `w:rFonts` 的 ascii/hAnsi/cs/eastAsia |
| 五通道优先级与 FR2 相反 | `_CHANNEL_ORDER = (latex, questionnaire, ai, template)` 从低到高叠加 → 实际生效 `template > ai > questionnaire > latex` |
| 引用维度不登记冲突 / LaTeX 来源误标「学校模板」 | `parse_latex_source` 补三项引用维度 `resolution`；`_overlay_one` 取消硬编码 `"school"` 默认；引用维度也登记冲突 |
| LaTeX 不剥离注释 | 新增 `_strip_comments()`（正确处理 `\%` 转义），解析前先剥注释 |
| 字体/字号采样含标题段 | 统计只取 `body_paras`（用 `is_heading` 判定，而非仅 `_heading_level`） |
| `w:eastAsia` 字体未读取 | `_run_font_name()` 补 eastAsia 兜底 |
| `fix --template` 坏文件崩溃 | `cmd_fix` 的 `_build_target` 补 `except DocxReadError` |
| 读取体未整体包错 | `read_docx` 整体包 try → `DocxReadError`（含「能打开但 XML 异常」） |
| 参考文献 Finding 无 source | `_check_reference(..., source=...)`；CLI/GUI 传 `ref_source` |
| `preview --json` 编码不一致 | 改走 `_write_json`（UTF-8 字节流） |
| GUI 缺 LaTeX 通道 | GUI 新增 .cls/.sty/.tex 选择并传入 `build_target(latex=...)` |
| 预览参考文献行无差别输出 | 按悬挂缩进/编码实际差异决定是否输出 |
| 两个产品决策 | 已明确并写入规格：**预览受门禁**（属付费价值的一部分）；**无变化不消耗试用**（并在 CLI/GUI 给出「无需修正」提示） |
| `MARGIN_TOL_WARN < FAIL` 笔误 | 改为 `FAIL=0.10 / WARN=0.05` |
| `HEARTBEAT_DAYS` 死常量 / 心跳未实现 | 心跳逻辑落地（3 天宽限，仅服务器明确 revoked/expired 才锁） |
| `ALIGN_LABEL` / margin 映射重复 | `checker.ALIGN_LABEL`、`specs.MARGIN_DIMS` 单一真相，`fixer`/`report` 改为引用 |
| 死导入、`__version__` | 清理；`__version__ = "2.0.0"` |
| 表格内文字「检查含、修正不含」不对称 | `fixer` 在修正摘要中明确提示「表格内文字（N 段）不参与修正」 |

### 7.3 端到端冒烟暴露的**新增**问题（本轮修，审查报告未覆盖）

这三条是**跑端到端才暴露**的，静态审查看不出来，已一并修掉：

| 问题 | 现象 | 修复 |
|---|---|---|
| **入参校验顺序错**（新） | `--out` 指向原件时，试用已用尽的账号拿到 **rc=3「免费试用已用完」**而非 rc=2「输出路径与输入同一文件」——参数错误被报成付费问题，且该校验在耗尽态下**根本不可达** | `cmd_fix` 重排为「先校验入参（文件存在 / `--out` 安全 / 目标画像）→ 再授权门禁 → 最后落盘」；GUI `_run_fix` 同步调整为「先算改动 → 需要动手才问授权」 |
| **已合规文档也触发门禁**（新） | 试用已用尽的账号修正一份**本来就合规**的文档被拒（rc=3），等于劝人白买码 | 「已合规」分支前移到门禁之前：rc=0 + 「未生成副本、未消耗试用额度」，不触发门禁（与 `check` 同级的免费体检） |
| **报告章节号跳号 + 数值重复前缀**（新） | 无「提示」段时章节从「一、」直跳「三、」，并出现过「一之二」；页边距行显示「实测：实测 2.0"」 | 章节号**动态分配**（只给实际渲染的段占号）；`_cmp_margin` 不再自带「实测」前缀（拼装是报告层的活）；维度裁决改按来源分组（14 行同值压成 1 行） |
| **修正成功后提示错位**（新） | 刚用完免费试用却仍显示「试用状态：首次免费试用（共 1 次）」 | 新增 `license.post_fix_message()` 展示**扣减后**状态；CLI/GUI 均改用 |

### 7.4 未完成 / 明确延后

- **中台 `/api/trial` 端点部署**：待老板提供 Cloudflare Token（补丁与部署步骤已在 `server_patch/README.md`）。
- **离线重置残余风险**：断网状态下删除本地状态文件仍可重置试用；需服务端一次性令牌方案才能根治（当前按「本地 fail-closed + 在线核对」折中，已在规格 `§8 已知边界` 标注）。
- **P2 nit 剩余**：`main.py:59` 兜底打印改显式失败、GUI 对用户显式选择路径不再加序号（当前 GUI 已不强制加序号，与 CLI 语义一致）、非中文控制台显示 `?`（已知边界）。
- **B15 覆盖率表述**：docstring 已去掉「ACM ~80%」的夸大表述，如实标注覆盖范围。

### 7.5 验证证据一览

| 验证项 | 结果 |
|---|---|
| `pytest tests/ -q`（含新增 18+6 项回归） | **96 passed in 6.90s** |
| CLI 端到端 12 步（真实 subprocess + 真实 docx） | 全部符合预期（含 rc=0 / rc=2 / rc=3 三条路径、原件字节不变、损坏文档友好报错） |
| 修正稿再体检 | 通过 12 · 提示 1 · 警告 0 · **不达标 0** |
| 报告章节号 | `['## 一、维度裁决（以谁为准）', '## 二、详细检查项']`，无跳号、无「实测：实测」 |
| 全新账号 + `--out` 原件 | rc=2，且**未生成状态文件**（未消耗额度） |

