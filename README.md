# Thesis Format Doctor Global · 海外版论文格式医生

离线、安全、本机处理的海外论文格式体检桌面端。**论文不出本机**，全程零网络、零大模型调用。

> 本仓库与国内版「论文格式医生」（`D:\AgentSpace\论文格式医生\`）**完全隔离、互不影响**：
> 独立目录、独立虚拟环境、独立代码，不 import 国内版任何模块。

---

## 一、定位与设计原则

| 项 | 说明 |
|---|---|
| 定位 | 离线安全（本机处理，论文不上传） |
| 内置规范 | APA / MLA / Chicago / IEEE / Harvard 写死进代码（`engine/specs.py`） |
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
D) 上传学校 Word 模板 (template)     ← 读模板自身的页面排版
E) LaTeX 模板 (.tex)                 ← 从 geometry / setlength 等命令推断页面要求
```

`merge_school()` 规则：**页面排版以学校为准**，引用格式以规范为准（除非学校显式覆盖）。

---

## 三、目录结构

```
thesis-format-doctor-global/
├── requirements.txt
├── README.md
├── VERSION                     # 版本号单一来源（运行时与 exe 属性都读它）
├── main.py                     # 打包入口：无参数 → 自动进 GUI（双击即开）
├── build.py                    # Nuitka 三平台构建
├── build_reedcode.py           # 卖家离线发码器构建（Windows）
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
│   │   └── license.py          # 授权：在线激活 + 离线激活码（HMAC 签名）
│   ├── ui/
│   │   ├── app.py              # tkinter 桌面 GUI（宣纸学术风，双栏卡片）
│   │   ├── i18n.py             # 中英文语言包（默认英文）
│   │   ├── fonts.py            # 跨平台字体兜底
│   │   └── iconpath.py         # 窗口图标路径解析
│   └── data/
│       ├── ai_questionnaire_template.json  # 交给 AI 填表的模板
│       └── icon.png            # 窗口图标（512）
├── tools/
│   ├── make_icon.py            # 源图 → 透明圆角 PNG/ICO/ICNS
│   ├── reedcode.py / reedcode_gui.py   # 卖家发码器（CLI / GUI）
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
.venv/Scripts/python.exe -m src.main check 论文.docx --ai ai填的表单.json

# 导出 AI 填表模板（交给任意 AI 用）
.venv/Scripts/python.exe -m src.main export-schema ai_questionnaire_template.json
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
  - 卖家端：`tools/reedcode_gui.py`（打包产物 `reedcode-global`，见 CI 的 `reedcode` job）
  - 算法：`机器码|时间戳|HMAC-SHA256 前 24 位`，密钥与国内版**不同**，两版激活码互不可用
- 查看状态：`status`（显示是否激活、来源、机器码）。

---

## 六、AI 填表工作流（你提的点子）

1. 运行 `export-schema` 导出 `ai_questionnaire_template.json`
2. 把该模板 + 你的学校模板/格式要求文档交给任意 AI
3. AI 按模板 fields 填出一份扁平 JSON（如 `{"margin_left_in":1.5,"font_family":"Times New Roman",...}`）
4. 在软件里「导入 AI 填表 JSON」即可，软件只负责吃这份 JSON

---

## 七、当前范围与待办（P2）

- ✅ 已落地：内置五规范、docx 读取、Rule Engine、对比报告、纯格式修正、GUI（中英双语、默认英文）、
  AI 填表导入、学校模板读取、LaTeX 模板读取、在线激活 + 离线激活码
- ⏳ PDF 范例解析（仅给 PDF 无 Word 的学校）：P2
- ⏳ 参考文献逐条语义校验（当前为著录形态启发式）：P2
- ⏳ 页码 / 页眉页脚检查：P2
- ⏳ 报告正文英文化（当前报告文案为中文，界面已双语）

---

## 八、说明

本仓库由「芦苇出海侦察队」产出，代码经多轮独立 AI 评审（见 `docs/codex_review_*.md`）。
所有格式数值来自各规范官方手册的典型文稿要求，学校特殊要求以模板/问卷覆盖为准。

---

## 九、构建与发布（开发者）

源码与打包链路完全独立于此仓库之外的中台/国内版，互不影响。

- **依赖**：仅 `python-docx`（见 `requirements.txt`）；GUI 用标准库 `tkinter`。
- **构建**：`python build.py` 调用 Nuitka 把整个 `src` 包编译成原生「文件夹版」（Windows/Linux 为 `dist/thesis-format-doctor-global/`，macOS 为 `.app`）。
  - Windows 额外用 NSIS 打安装包：`makensis /DVERSION=x.y.z installer.nsi`（需先装 NSIS；界面全英文）。
  - 卖家发码器：`python build_reedcode.py` → `dist/reedcode-global/`。
- **CI 三平台构建**：`.github/workflows/build.yml` 在推送 `v*` 标签时，并行构建
  Windows（zip 免安装版 + setup 安装包）、macOS（`.app` zip）、Linux（文件夹 zip），
  外加 Windows 发码器（`reedcode` job），然后自动发布到 GitHub Release（共 5 个产物，
  `release` job 会校验产物齐全才发）。
- **发版流程**：`bump VERSION` → `git tag vX.Y.Z` → `git push origin vX.Y.Z`。
- **图标**：`python tools/make_icon.py 源图.png --outdir assets --name icon` —— 自动裁白边、
  按源图圆角生成透明蒙版，输出 1024 PNG + 多尺寸 ICO + ICNS。
