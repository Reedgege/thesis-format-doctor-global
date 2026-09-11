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

## 二、四种目标来源，统一到一份画像

无论用户从哪条路进来，最终都收敛到 `TargetProfile`，检查逻辑只面对一种结构：

```
A) 选内置规范 (spec)
B) 手填格式问卷 (questionnaire)
C) 导入 AI 填好的 JSON (ai_import)   ← 用户把学校模板交给 GPT/Gemini/Claude 填表
D) 上传学校 Word 模板 (template)     ← 读模板自身的页面排版
```

`merge_school()` 规则：**页面排版以学校为准**，引用格式以规范为准（除非学校显式覆盖）。

---

## 三、目录结构

```
thesis-format-doctor-global/
├── requirements.txt
├── README.md
├── src/
│   ├── main.py                 # CLI 入口（gui / check / export-schema）
│   ├── engine/
│   │   ├── specs.py            # 内置五规范（硬编码，离线）
│   │   ├── questionnaire.py    # 格式问卷 schema + TargetProfile + 四输入汇聚
│   │   ├── docx_reader.py      # 读 docx：页边距/字体/字号/行距/标题/参考文献
│   │   ├── checker.py          # Rule Engine：确定性比对
│   │   └── report.py           # 分维度裁决报告
│   ├── ui/
│   │   └── app.py              # tkinter 桌面 GUI
│   └── data/
│       └── ai_questionnaire_template.json  # 交给 AI 填表的模板
└── tests/
    └── smoke_test.py           # 端到端冒烟测试
```

---

## 四、运行方式

### 1. 安装依赖（项目本地虚拟环境，已建好 `.venv`）
```bash
.venv/Scripts/python.exe -m pip install -r requirements.txt
```

### 2. 命令行体检
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

### 3. 桌面 GUI
```bash
.venv/Scripts/python.exe -m src.main gui
```
GUI 需本机带 Tk（_tkinter）。CI 打包（Nuitka/PyInstaller）后再分发，与国内版打包流程一致。

---

## 五、AI 填表工作流（你提的点子）

1. 运行 `export-schema` 导出 `ai_questionnaire_template.json`
2. 把该模板 + 你的学校模板/格式要求文档交给任意 AI
3. AI 按模板 fields 填出一份扁平 JSON（如 `{"margin_left_in":1.5,"font_family":"Times New Roman",...}`）
4. 在软件里「导入 AI 填表 JSON」即可，软件只负责吃这份 JSON

---

## 六、当前范围与待办（P2）

- ✅ v1 已落地：内置五规范、docx 读取、Rule Engine、对比报告、GUI、AI 填表导入、学校模板读取
- ⏳ LaTeX 模板（理工科，MIT 等）：独立体系，P2
- ⏳ PDF 范例解析（仅给 PDF 无 Word 的学校）：P2
- ⏳ 参考文献逐条语义校验（当前为著录形态启发式）：P2

---

## 七、说明

本 v1 由「芦苇出海侦察队」产出，待交由 GPT / Gemini 做代码评审后迭代。
所有格式数值来自各规范官方手册的典型文稿要求，学校特殊要求以模板/问卷覆盖为准。

---

## 八、构建与发布（开发者）

源码与打包链路完全独立于此仓库之外的中台/国内版，互不影响。

- **依赖**：仅 `python-docx`（见 `requirements.txt`）；GUI 用标准库 `tkinter`。
- **构建**：`python build.py` 调用 Nuitka 把整个 `src` 包编译成原生「文件夹版」（Windows/Linux 为 `dist/thesis-format-doctor-global/`，macOS 为 `.app`）。
  - Windows 额外用 NSIS 打安装包：`makensis installer.nsi`（需先装 NSIS）。
- **CI 三平台构建**：`.github/workflows/build.yml` 在推送 `v*` 标签时，并行构建
  Windows（zip 免安装版 + setup 安装包）、macOS（`.app` zip）、Linux（文件夹 zip），
  然后自动发布到 GitHub Release。
- **发版流程**：`bump VERSION` → `git tag vX.Y.Z` → `git push origin vX.Y.Z`。
- **图标**：`assets/icon.png` / `assets/icon.ico`，可用 `python tools/make_icon.py` 重新生成。
