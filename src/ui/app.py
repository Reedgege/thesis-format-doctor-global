"""Thesis Format Doctor Global —— tkinter 桌面 GUI（离线，论文不出本机）。

流程：
  选规范 →（可选）上传学校模板 / LaTeX 模板 / 手填问卷 / 导入 AI 填表
  → 选待检文档 → 运行 → 看对比报告 →（可选）一键修正。

GUI 与 CLI 共用同一套 engine，保证逻辑唯一、离线安全。
"""

from __future__ import annotations

import os
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

from ..engine import checker, report as report_mod
from ..engine.docx_reader import DocxReadError
from ..engine.fixer import compute_changes, fix_docx, same_file
from ..engine.questionnaire import (
    FIELD_RANGES, QUESTIONNAIRE_SCHEMA, TargetProfile, build_target,
)
from ..engine.specs import SPEC_ORDER
from ..license import license as lic

_HERE = os.path.dirname(os.path.abspath(__file__))
_TEMPLATE_PATH = os.path.join(_HERE, "..", "data", "ai_questionnaire_template.json")

# 数值字段的合理范围（由问卷 schema 派生，单一真相；越界只提示、不阻止保存）
_FIELD_RANGE = FIELD_RANGES


class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Thesis Format Doctor Global · 海外版论文格式医生")
        self.root.geometry("860x700")

        self.spec_key = tk.StringVar(value="APA")
        self.school_template = tk.StringVar(value="")
        self.latex_template = tk.StringVar(value="")
        self.ai_json = tk.StringVar(value="")
        self.docx_path = tk.StringVar(value="")
        self.questionnaire_data: dict = {}

        self._build()

    # ---------------------------------------------------------------- UI
    def _build(self):
        f = ttk.Frame(self.root, padding=10)
        f.pack(fill="both", expand=True)

        # 规范选择
        ttk.Label(f, text="① 选择引用规范：").grid(row=0, column=0, sticky="w", pady=2)
        cb = ttk.Combobox(f, textvariable=self.spec_key, values=SPEC_ORDER,
                          state="readonly", width=18)
        cb.grid(row=0, column=1, sticky="w")
        ttk.Label(f, text="Other=自定义，须配学校模板/问卷").grid(row=0, column=2, sticky="w")

        # 学校模板
        ttk.Label(f, text="② 学校模板（可选）：").grid(row=1, column=0, sticky="w", pady=2)
        ttk.Button(f, text="上传 Word 模板", command=self._pick_template).grid(row=1, column=1, sticky="w")
        ttk.Label(f, textvariable=self.school_template, foreground="#555").grid(
            row=1, column=2, columnspan=2, sticky="w")

        # LaTeX 模板
        ttk.Label(f, text="③ LaTeX 模板（可选）：").grid(row=2, column=0, sticky="w", pady=2)
        ttk.Button(f, text="上传 .cls/.sty/.tex", command=self._pick_latex).grid(
            row=2, column=1, sticky="w")
        ttk.Label(f, textvariable=self.latex_template, foreground="#555").grid(
            row=2, column=2, columnspan=2, sticky="w")

        # 问卷 / AI 导入
        ttk.Label(f, text="④ 格式问卷（可选）：").grid(row=3, column=0, sticky="w", pady=2)
        ttk.Button(f, text="手填问卷", command=self._open_questionnaire).grid(row=3, column=1, sticky="w")
        ttk.Button(f, text="导入 AI 填表 JSON", command=self._pick_ai_json).grid(row=3, column=2, sticky="w")
        ttk.Label(f, textvariable=self.ai_json, foreground="#555").grid(
            row=4, column=1, columnspan=2, sticky="w")

        # 待检文档
        ttk.Label(f, text="⑤ 待检文档：").grid(row=5, column=0, sticky="w", pady=2)
        ttk.Entry(f, textvariable=self.docx_path, width=40).grid(row=5, column=1, sticky="w")
        ttk.Button(f, text="浏览", command=self._pick_docx).grid(row=5, column=2, sticky="w")

        # 运行
        ttk.Button(f, text="▶ 运行格式体检", command=self._run).grid(row=6, column=1, pady=8)
        ttk.Button(f, text="✎ 一键修正（仅改格式）", command=self._run_fix).grid(row=6, column=2, sticky="w")
        ttk.Button(f, text="🔑 输入激活码", command=self._activate_dialog).grid(row=6, column=3, sticky="w")
        ttk.Button(f, text="导出 AI 问卷模板", command=self._export_template).grid(row=7, column=2, sticky="w")
        ttk.Button(f, text="保存报告", command=self._save_report).grid(row=7, column=3, sticky="w")

        # 报告区
        ttk.Label(f, text="体检报告：").grid(row=8, column=0, sticky="w", pady=(6, 2))
        self.out = scrolledtext.ScrolledText(f, wrap="word", width=104, height=24)
        self.out.grid(row=9, column=0, columnspan=4, sticky="nsew")
        f.rowconfigure(9, weight=1)

    # ---------------------------------------------------------------- handlers
    def _pick_template(self):
        p = filedialog.askopenfilename(filetypes=[("Word", "*.docx")])
        if p:
            self.school_template.set(p)

    def _pick_latex(self):
        p = filedialog.askopenfilename(
            filetypes=[("LaTeX 模板", "*.tex *.cls *.sty"), ("所有文件", "*.*")])
        if p:
            self.latex_template.set(p)

    def _pick_ai_json(self):
        p = filedialog.askopenfilename(filetypes=[("JSON", "*.json")])
        if p:
            self.ai_json.set(p)

    def _pick_docx(self):
        p = filedialog.askopenfilename(filetypes=[("Word", "*.docx")])
        if p:
            self.docx_path.set(p)

    def _open_questionnaire(self):
        win = tk.Toplevel(self.root)
        win.title("格式问卷")
        win.geometry("660x640")
        vars_ = {}
        row = 0
        for field in QUESTIONNAIRE_SCHEMA:
            key = field["key"]
            ttk.Label(win, text=field["label"]).grid(row=row, column=0, sticky="w", padx=8, pady=3)
            ftype = field.get("type")
            if ftype == "bool":
                v = tk.BooleanVar(value=bool(field.get("default", False)))
                ttk.Checkbutton(win, variable=v).grid(row=row, column=1, sticky="w")
            elif ftype == "choice":
                v = tk.StringVar(value=str(field.get("default", "")))
                ttk.Combobox(win, textvariable=v, values=field.get("options", []),
                             state="readonly", width=18).grid(row=row, column=1, sticky="w")
            else:
                v = tk.StringVar(value=str(field.get("default", "")))
                ttk.Entry(win, textvariable=v, width=20).grid(row=row, column=1, sticky="w")
            vars_[key] = v
            hint = (field.get("unit", "") or "")
            if field.get("help"):
                hint = (hint + " " + field["help"]).strip()
            ttk.Label(win, text=hint, foreground="#888",
                      wraplength=270, justify="left").grid(row=row, column=2, sticky="w")
            row += 1

        def _save():
            data = {}
            warnings = []
            for field in QUESTIONNAIRE_SCHEMA:
                key = field["key"]
                if field.get("type") == "bool":
                    data[key] = bool(vars_[key].get())
                    continue
                val = vars_[key].get().strip()
                if not val:
                    continue
                if field.get("type") in ("float", "int"):
                    try:
                        num = float(val)
                    except ValueError:
                        messagebox.showwarning(
                            "数值格式不对",
                            f"「{field['label']}」需要填数字，当前是：{val}", parent=win)
                        return
                    rng = _FIELD_RANGE.get(key)
                    if rng and not (rng[0] <= num <= rng[1]):
                        warnings.append(f"「{field['label']}」= {val}（常见范围 {rng[0]}–{rng[1]}）")
                data[key] = val
            if warnings and not messagebox.askyesno(
                    "数值超出常见范围",
                    "以下数值超出常见范围，是否仍要保存？\n\n" + "\n".join(warnings), parent=win):
                return
            self.questionnaire_data = data
            messagebox.showinfo("已保存", "问卷已填入，运行体检时生效。", parent=win)
            win.destroy()

        ttk.Button(win, text="保存问卷", command=_save).grid(row=row + 1, column=1, pady=12)

    def _export_template(self):
        dst = filedialog.asksaveasfilename(defaultextension=".json",
                                           filetypes=[("JSON", "*.json")],
                                           initialfile="ai_questionnaire_template.json")
        if not dst:
            return
        with open(_TEMPLATE_PATH, "r", encoding="utf-8") as src, \
                open(dst, "w", encoding="utf-8") as out:
            out.write(src.read())
        messagebox.showinfo("已导出", f"AI 问卷模板已导出到：\n{dst}\n把它交给任意 AI 填表即可。")

    def _save_report(self):
        content = self.out.get("1.0", tk.END).strip()
        if not content:
            messagebox.showwarning("没有报告", "请先运行「格式体检」生成报告。")
            return
        dst = filedialog.asksaveasfilename(
            defaultextension=".md", filetypes=[("Markdown", "*.md")],
            initialfile="格式体检报告.md", title="保存体检报告")
        if not dst:
            return
        with open(dst, "w", encoding="utf-8") as fh:
            fh.write(content + "\n")
        messagebox.showinfo("已保存", f"报告已保存到：\n{dst}")

    def _build_target(self) -> TargetProfile:
        return build_target(
            self.spec_key.get(),
            template=self.school_template.get() or None,
            ai=self.ai_json.get() or None,
            questionnaire=self.questionnaire_data or None,
            latex=self.latex_template.get() or None,
        )

    def _show(self, text: str):
        self.out.delete("1.0", tk.END)
        self.out.insert(tk.END, text)

    def _run(self):
        docx = self.docx_path.get()
        if not docx or not os.path.isfile(docx):
            messagebox.showwarning("缺少文档", "请先选择待检 Word 文档（.docx）。")
            return
        try:
            target = self._build_target()
            prof, findings = checker.run_check(docx, target)
            rep = report_mod.build_report(docx, target, prof, findings)
        except DocxReadError as e:
            messagebox.showerror("无法读取文档", str(e))
            return
        except (ValueError, FileNotFoundError, OSError) as e:
            messagebox.showerror("检查失败", str(e))
            return
        except Exception as e:  # noqa: BLE001  兜底，防止 GUI 因未知异常僵死
            messagebox.showerror("检查失败（未预期错误）", str(e))
            return
        self._show(rep.markdown)

    def _run_fix(self):
        """一键修正：先算改动（免费）→ 需要动手才问授权 → 预览确认 → 落盘。

        顺序与 CLI 一致：**已合规文档不触发门禁**（否则「文档本来就不用改」也会被
        报成「试用已用完」，既误导用户又等于劝人白买码）。
        """
        docx = self.docx_path.get()
        if not docx or not os.path.isfile(docx):
            messagebox.showwarning("缺少文档", "请先选择待修正 Word 文档（.docx）。")
            return

        # ① 免费：算改动（只读文档，无任何副作用）
        try:
            target = self._build_target()
            changes = compute_changes(docx, target)
        except DocxReadError as e:
            messagebox.showerror("无法读取文档", str(e))
            return
        except (ValueError, FileNotFoundError, OSError) as e:
            messagebox.showerror("预览失败", str(e))
            return
        except Exception as e:  # noqa: BLE001
            messagebox.showerror("预览失败（未预期错误）", str(e))
            return

        if not changes:
            msg = "当前文档已符合目标格式，无需修正（未生成副本、未消耗试用额度）。"
            self._show(msg)
            messagebox.showinfo("无需修正", msg)
            return

        # ② 真要动手改，才做授权裁决（模型 B：首次免费、之后需激活码）
        gate = lic.require_fix_entitlement()
        if not gate["allowed"]:
            messagebox.showwarning("需要激活", gate["message"])
            return

        preview_text = "将进行的格式修正（仅改格式、绝不改正文）：\n" + \
                       "\n".join("  · " + line for line in changes) + \
                       f"\n\n试用状态：{gate['message']}"
        self._show(preview_text)
        if not messagebox.askyesno("确认修正", preview_text +
                                   "\n\n点「是」开始修正，随后选择保存位置（原件不会被改动）。"):
            return

        default_name = os.path.splitext(os.path.basename(docx))[0] + "_fixed.docx"
        out = filedialog.asksaveasfilename(
            title="选择修正稿保存位置（默认 原名_fixed.docx）",
            defaultextension=".docx", filetypes=[("Word", "*.docx")],
            initialfile=default_name,
            initialdir=os.path.dirname(docx) or None)
        if not out:
            return
        # 用户在保存对话框里显式选择的路径以用户为准（对话框已自带覆盖确认）；
        # 只拦「与原件同一文件」这一种不可逆情形。
        if same_file(out, docx):
            messagebox.showerror("不能覆盖原件",
                                 "输出路径与待修正文档是同一个文件，请另选保存位置。")
            return

        try:
            fix_docx(docx, out, target)
            counted = lic.record_fix_used()
        except DocxReadError as e:
            messagebox.showerror("无法读取文档", str(e))
            return
        except (ValueError, OSError) as e:
            messagebox.showerror("修正失败", str(e))
            return
        # 展示**扣减后**的状态（gate['message'] 是扣减前的，会显示成「首次免费试用」）
        messagebox.showinfo("修正完成",
                            f"已生成：\n{out}\n\n{lic.post_fix_message(gate, counted)}")

    def _activate_dialog(self):
        win = tk.Toplevel(self.root)
        win.title("激活")
        win.geometry("480x380")
        win.resizable(False, False)

        # ① 在线激活（中台签发）
        ttk.Label(win, text="① 在线激活码（中台签发）：").grid(
            row=0, column=0, padx=12, pady=(12, 2), sticky="w")
        code_var = tk.StringVar()
        ttk.Entry(win, textvariable=code_var, width=40).grid(
            row=1, column=0, columnspan=2, padx=12, sticky="we")

        def _submit():
            code = code_var.get().strip()
            if not code:
                messagebox.showwarning("缺少激活码", "请先输入激活码。", parent=win)
                return
            r = lic.activate(code)
            messagebox.showinfo("激活结果", r["message"], parent=win)
            if r["ok"]:
                win.destroy()

        ttk.Button(win, text="激活", command=_submit).grid(
            row=2, column=0, padx=12, pady=6, sticky="w")

        ttk.Separator(win, orient="horizontal").grid(
            row=3, column=0, columnspan=2, padx=12, pady=8, sticky="we")

        # ② 复制本机机器码（发给卖家，由卖家生成离线码）
        ttk.Label(win, text="② 本机机器码（发给卖家生成离线码）：").grid(
            row=4, column=0, padx=12, pady=(2, 2), sticky="w")
        mc_var = tk.StringVar(value=lic._machine_fingerprint())
        ttk.Entry(win, textvariable=mc_var, width=40, state="readonly").grid(
            row=5, column=0, columnspan=2, padx=12, sticky="we")

        def _copy_mc():
            try:
                win.clipboard_clear()
                win.clipboard_append(mc_var.get())
                messagebox.showinfo(
                    "已复制", "本机机器码已复制到剪贴板，把它发给卖家即可。", parent=win)
            except Exception:
                pass

        ttk.Button(win, text="复制机器码", command=_copy_mc).grid(
            row=6, column=0, padx=12, pady=6, sticky="w")

        ttk.Separator(win, orient="horizontal").grid(
            row=7, column=0, columnspan=2, padx=12, pady=8, sticky="we")

        # ③ 离线激活（卖家签发的离线码，无需联网）
        ttk.Label(win, text="③ 离线激活码（卖家签发，无需联网）：").grid(
            row=8, column=0, padx=12, pady=(2, 2), sticky="w")
        offline_var = tk.StringVar()
        ttk.Entry(win, textvariable=offline_var, width=40).grid(
            row=9, column=0, columnspan=2, padx=12, sticky="we")

        def _submit_offline():
            code = offline_var.get().strip()
            if not code:
                messagebox.showwarning("缺少离线码", "请先粘贴卖家给的离线激活码。", parent=win)
                return
            r = lic.activate_offline(code)
            messagebox.showinfo("激活结果", r["message"], parent=win)
            if r["ok"]:
                win.destroy()

        ttk.Button(win, text="离线激活", command=_submit_offline).grid(
            row=10, column=0, padx=12, pady=6, sticky="w")


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
