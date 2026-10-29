# PaperFormat Pro · 海外版论文格式医生

离线、安全、本机处理的海外论文格式体检桌面端。**论文不出本机**，零大模型调用；联网**仅**用于「激活校验」与「首次试用登记（只送机器码）」，其余全程离线。

> 本仓库与国内版「论文格式医生」（`D:\AgentSpace\论文格式医生\`）**完全隔离、互不影响**：
> 独立目录、独立虚拟环境、独立代码，不 import 国内版任何模块。

---

## 一、定位与设计原则

| 项 | 说明 |
|---|---|
| 定位 | 离线安全（本机处理，论文不上传） |
| 内置规范 | APA（学生版 / 专业版）、MLA、IEEE（会议 / 期刊）写死进代码（`engine/specs.py`）；**无 Chicago、无 Harvard**（以 `SPEC_ORDER` 为唯一事实源） |
| 可选学校模板 | 上传学校 Word 模板 → 读其页面排版作为「学校要求」 |
| 双轨对比 | 跑完出「分维度裁决报告」：哪些以规范为准、哪些以学校为准 |
| 格式问卷 + AI 填表 | 无模板/解析失败时的兜底；用户可手填，或把模板丢给任意 AI 填表后导入 |
| 确定性引擎 | 所有判断由 Rule Engine 完成，**不调 LLM** |

**模板驱动铁律**：页面排版（页边距/字体/行距）由学校模板或问卷决定；引用格式以所选规范为准。
真实欧美高校样本反复印证：页面排版归学校、引用格式归规范，二者分开。

---

## 二、五种目标来源，统一到一份画像

无论用户从哪条路进来，最终都收敛到 `TargetProfile`，检查逻辑只面对一种结构：

```
A) 选内置规范 (spec)
B) 手填格式问卷 (questionnaire)
C) 导入 AI 填好的 JSON (ai_import)   ← 用户把学校模板交给 GPT/Gemini/Claude 填表
D) 上传学校 Word 模板 (template)     ← 先读模板里的 Word 批注（规范原文），读不到再量实际排版
E) LaTeX 模板 (.tex)                 ← 从 geometry / setlength 等命令推断页面要求
```

`merge_school()` 规则：**页面排版以学校为准**，引用格式以规范为准（除非学校显式覆盖）。

### 学校模板内部：批注 > 实测

上传学校模板时，同一维度可能有两处来源，优先级固定为：

```
Word 批注（模板里写的规范原文）  >  实测排版（量模板自身长什么样）
```

- **批注**由 `engine/comment_spec_en.py` 解析（**纯英文**，海外版不做双语）：
  认 `Heading 1 / H2 / Level 3 / First-level heading`、`pt / inch / cm / mm`、
  `centered / justified / left-aligned`、`bold / italic`（含 `not bold`）、
  `single / 1.5 / double-spaced`、`first-line / hanging indent`、
  `Margins / Paper size / Layout（单双栏）/ Tables（三线表）/ Table spacing /
  References（APA · MLA · IEEE）/ Figure & Table captions / Page numbers`。
- 只有「类别词 + 冒号」开头的批注才被当作规范；`fix this later` 这类编辑性批注会被
  忽略并记进 `unparsed`（诊断用），**绝不误读成格式要求**。
- 批注没提到的维度一律回落实测 —— 不会因为「有批注」就丢掉模板的真实排版。
- **表后间隙**：模板没写、实测又是 0（表格紧贴文字）时，按通用兜底 ≥12pt，并在报告里
  注明「模板未明确，按通用留白处理」（来源标 `General default (template silent)`）。
- **参考文献流派**：学校模板子面板的下拉（APA / MLA / IEEE）**优先**；未选时才用模板
  批注推断；两者都没有则不查参考文献。IEEE 是顺序编码制 —— 会报未编号条目并自动补 `[n]`。

---

## 三、目录结构

```
paperformatpro/
├── requirements.txt
├── README.md
├── VERSION                     # 版本号单一来源（运行时与 exe 属性都读它）
├── main.py                     # 打包入口：无参数 → 自动进 GUI（双击即开）
├── build.py                    # Nuitka 三平台构建
├── installer.nsi               # NSIS 安装包（全英文界面）
├── assets/                     # icon.png(1024) / icon.ico / icon.icns
├── src/
│   ├── main.py                 # CLI 入口（gui / check / fix / activate / status / export-schema）
│   ├── versioninfo.py          # 版本号读取（CLI 与 GUI 共用）
│   ├── engine/
│   │   ├── specs.py            # 内置五规范（硬编码，离线）
│   │   ├── questionnaire.py    # 格式问卷 schema + TargetProfile + 五输入汇聚
│   │   ├── docx_reader.py      # 读 docx：页边距/字体/字号/行距/标题/参考文献
│   │   ├── checker.py          # Rule Engine：确定性比对
│   │   ├── fixer.py            # 纯格式层修正（绝不改内容）
│   │   └── report.py           # 分维度裁决报告
│   ├── license/
│   │   ├── license.py          # 授权：在线激活（中台）+ 离线激活
│   │   ├── crypto.py           # 内嵌 Ed25519 公钥（只验签）+ verify_offline_code
│   │   └── ed25519_verify.py   # 纯标准库 Ed25519 验签（零第三方加密依赖）
│   ├── ui/
│   │   ├── app.py              # tkinter 桌面 GUI（宣纸学术风，双栏卡片）
│   │   ├── i18n.py             # 中英文语言包（默认英文）
│   │   ├── fonts.py            # 跨平台字体兜底
│   │   └── iconpath.py         # 窗口图标路径解析
│   └── data/
│       ├── ai_questionnaire_template.yaml  # 交给 AI 填表的模板
│       └── icon.png            # 窗口图标（512）
├── tools/
│   ├── make_icon.py            # 源图 → 透明圆角 PNG/ICO/ICNS
│   ├── verify_entry.py         # 入口行为验证（无参数必须进 GUI）
│   ├── smoke_gui.py            # 真机 GUI 冒烟
│   └── check_layout.py         # 控件压缩/越界检查
└── tests/                      # pytest 回归套件
```

---

## 四、运行方式

### 1. 安装依赖（项目本地虚拟环境，已建好 `.venv`）
```bash
.venv/Scripts/python.exe -m pip install -r requirements.txt
```

### 2. 命令行体检（免费不限次）
```bash
# 纯规范
.venv/Scripts/python.exe -m src.main check 论文.docx --spec APA

# 叠加学校模板（页面以学校为准）
.venv/Scripts/python.exe -m src.main check 论文.docx --spec APA --template 学校模板.docx

# 导入 AI 填好的问卷
.venv/Scripts/python.exe -m src.main check 论文.docx --ai ai填的表单.yaml

# 导出 AI 填表模板（交给任意 AI 用）
.venv/Scripts/python.exe -m src.main export-schema ai_questionnaire_template.yaml
```

### 3. 命令行修正（首次免费，之后需激活码）
```bash
# 只预览要改的项，不落盘
.venv/Scripts/python.exe -m src.main fix 论文.docx --spec APA --preview

# 真修正（默认写 原名_fixed.docx，已存在时自动加序号、绝不覆盖原件）
.venv/Scripts/python.exe -m src.main fix 论文.docx --spec APA
```

### 4. 桌面 GUI
```bash
.venv/Scripts/python.exe -m src.main gui      # 或直接运行 main.py（无参数即进 GUI）
```
GUI 需本机带 Tk（_tkinter）。**界面默认英文**，右上角可一键切中文（偏好落盘）。打包后的 exe 双击即开。

---

## 五、授权与激活

- **免费**：格式体检（`check`）永远免费、不限次。
- **首次修正免费**；第 2 次起需要激活码（一次性买断，激活后本机永久可用）。
- **在线激活**：`activate 激活码` —— 联网向中台核验，之后每 3 天一次心跳，中台只在明确吊销时才锁。
- **离线激活**（无网环境）：客户在 GUI 激活弹窗（或 `status`）里取「本机机器码」，发给卖家；
  卖家用自己的发码器签发一条绑定该机器码的离线码；客户粘回即可，**全程不联网、不回传**。
  - 卖家端：统一发码器 `reedcode_unified.py`（仅存卖家本机 `_signing_keys/`，不入库、不入安装包）
  - 算法：对 `机器码|时间戳` 做 **Ed25519 签名**（私钥只在卖家本机；公钥内嵌客户端，
    只能验签不能签名）。密钥与国内两版**不同**，三版激活码互不可用。
- 查看状态：`status`（显示是否激活、来源、机器码）。

---

## 六、AI 填表工作流（你提的点子）

1. 运行 `export-schema` 导出 `ai_questionnaire_template.yaml`
2. 把该模板 + 你的学校模板/格式要求文档交给任意 AI
3. AI 按模板 fields 填出一份 YAML（如 `margin_left_in: 1.5` / `font_family: Times New Roman` …，也可包在 `filled:` 下）
4. 在软件里「导入 AI 填表」即可，软件只负责吃这份 YAML（旧版 JSON 也兼容）

---

## 七、当前范围与待办（P2）

- ✅ 已落地：内置规范（APA 学生/专业、MLA、IEEE 会议/期刊）、docx 读取、Rule Engine、对比报告、纯格式修正、GUI（中英双语、默认英文）、
  AI 填表导入、学校模板读取、LaTeX 模板读取、在线激活 + 离线激活码
- ⏳ PDF 范例解析（仅给 PDF 无 Word 的学校）：P2
- ⏳ 参考文献逐条语义校验（当前为著录形态启发式）：P2
- ✅ 页码位置 / running head 已落地（v2.3.90 起可查可修）；页眉页脚**内容**仍只提示
- ✅ 报告正文已英文化（海外版客户可见处不出现中文）

---

## 八、说明

本仓库由「芦苇出海侦察队」产出，代码经多轮独立 AI 评审（见 `docs/codex_review_*.md`）。
所有格式数值来自各规范官方手册的典型文稿要求，学校特殊要求以模板/问卷覆盖为准。

---

## 九、构建与发布（开发者）

源码与打包链路完全独立于此仓库之外的中台/国内版，互不影响。

- **依赖**：仅 `python-docx`（见 `requirements.txt`）；GUI 用标准库 `tkinter`；
  离线激活验签为纯标准库 Ed25519（`src/license/ed25519_verify.py`），**无第三方加密依赖**。
- **构建**：`python build.py` 调用 Nuitka 把整个 `src` 包编译成原生「文件夹版」（Windows/Linux 为 `dist/paperformatpro/`，macOS 为 `.app`）。
  - Windows 额外用 NSIS 打安装包：`makensis /DVERSION=x.y.z installer.nsi`（需先装 NSIS；界面全英文）。
  - 卖家发码器：统一发码器 `reedcode_unified.py` 仅存卖家本机（`_signing_keys/`），
    **不参与 CI、不随产品分发**（私钥绝不入库）。
- **CI 三平台构建**：`.github/workflows/build.yml` 在推送 `v*` 标签时，并行构建
  Windows（zip 免安装版 + setup 安装包）、macOS（`.app` zip）、Linux（文件夹 zip），
  然后自动发布到 GitHub Release（共 4 个产物，`release` job 会校验产物齐全才发）。
- **发版流程**：`bump VERSION` → `git tag vX.Y.Z` → `git push origin vX.Y.Z`。
- **图标**：`python tools/make_icon.py 源图.png --outdir assets --name icon` —— 自动裁白边、
  按源图圆角生成透明蒙版，输出 1024 PNG + 多尺寸 ICO + ICNS。
