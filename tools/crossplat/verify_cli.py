#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""跨平台功能等价性验证（check -> fix -> 复检）。

用途
====
在 macOS / Linux 的 CI runner 上，用**打包好的二进制**跑一遍真实论文的
「体检 -> 一键修正 -> 再体检」，把结果与 Windows 产物产出的基准逐条比对。
只要有一项对不上，就说明这个平台的产物功能受损（而不是"能启动"就算过）。

为什么不能靠单元测试
--------------------
单测跑的是源码，测不出"打包时漏收数据文件 / 平台分支写错 / 冻结后路径失效"。
这个脚本跑的是**客户真正下载到的那个可执行文件**，所以能测出来。

判定规则
--------
* 版本号与基准相同 -> 严格比对：summary 与 findings 全量必须逐条一致。
* 版本号与基准不同 -> 只打印差异并给 notice 警告，不判红。
  （基准过期是维护问题，不是产物问题；但必须显式提醒去刷新基准。）

退出码
------
0 = 通过（或版本不同但无致命错误）
1 = 功能回归（同版本下结果与基准不一致）
2 = 环境/执行失败（程序跑不起来、超时、输出不是 JSON）

用法
----
  python3 tools/crossplat/verify_cli.py \
      --exe "/path/to/paperformatpro" \
      --label "macOS arm64" \
      [--spec APA-Student] \
      [--baseline tools/crossplat/baseline_apa_student.json] \
      [--fixture tools/crossplat/fixtures/bad_apa.docx] \
      [--workdir /tmp/crossplat]
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile

for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

# GitHub Actions 注解里 % 和换行会让消息截断，统一转义
def _esc(s: str) -> str:
    return s.replace("%", "%25").replace("\r", "").replace("\n", "%0A")


def notice(msg: str) -> None:
    print("::notice::" + _esc(msg), flush=True)


def error(msg: str) -> None:
    print("::error::" + _esc(msg), flush=True)


def warn(msg: str) -> None:
    print("::warning::" + _esc(msg), flush=True)


def step(msg: str) -> None:
    print("\n=== " + msg + " ===", flush=True)


# ---------------------------------------------------------------------------
# 执行受测程序
# ---------------------------------------------------------------------------
class RunResult:
    def __init__(self, rc: int, out: str, err: str, timed_out: bool = False):
        self.rc = rc
        self.out = out
        self.err = err
        self.timed_out = timed_out


def run_exe(exe: str, args: list[str], cwd: str, timeout: int) -> RunResult:
    """跑受测程序。超时视为失败而不是挂死 CI。"""
    env = dict(os.environ)
    env["PYTHONIOENCODING"] = "utf-8"
    env["PYTHONUTF8"] = "1"
    try:
        p = subprocess.run(
            [exe] + args,
            cwd=cwd,
            capture_output=True,
            timeout=timeout,
            env=env,
        )
        return RunResult(
            p.returncode,
            p.stdout.decode("utf-8", errors="replace"),
            p.stderr.decode("utf-8", errors="replace"),
        )
    except subprocess.TimeoutExpired as e:
        out = (e.stdout or b"").decode("utf-8", errors="replace")
        err = (e.stderr or b"").decode("utf-8", errors="replace")
        return RunResult(-1, out, err, timed_out=True)
    except OSError as e:
        return RunResult(-1, "", "cannot execute: %s" % e)


def extract_json(text: str):
    """从混着日志的输出里抠出 JSON 对象。

    冻结后的程序偶尔会在 JSON 前后打印东西（Nuitka onefile 的解包提示、
    或引擎的告警），不能假设 stdout 只有 JSON。
    """
    if not text or not text.strip():
        return None
    stripped = text.strip()
    try:
        return json.loads(stripped)
    except Exception:
        pass
    start = stripped.find("{")
    end = stripped.rfind("}")
    if start < 0 or end <= start:
        return None
    try:
        return json.loads(stripped[start : end + 1])
    except Exception:
        return None


def fingerprint(payload: dict):
    """(summary, 排序后的 findings 三元组列表)。"""
    summary = payload.get("summary") or {}
    findings = payload.get("findings") or []
    triples = []
    for f in findings:
        if not isinstance(f, dict):
            continue
        triples.append(
            [
                str(f.get("dimension")),
                str(f.get("severity")),
                str(f.get("message")),
            ]
        )
    triples.sort()
    return summary, triples


def sha_of(triples) -> str:
    return hashlib.sha256(
        json.dumps(triples, ensure_ascii=False).encode("utf-8")
    ).hexdigest()


# ---------------------------------------------------------------------------
# 取证：把修正稿的页眉/页脚结构摊平打印
# ---------------------------------------------------------------------------
# 为什么需要它：跨平台差异跑出来时，注解里只有「哪个维度不一致」，
# 看不到"写进去没有 / 写成什么样"。而 CI 的完整日志要鉴权才读得到，
# 注解是唯一匿名可读的通道。所以把关键结构压成几行短文本走注解。
def diagnose_docx(path: str, max_lines: int = 8) -> list[str]:
    import zipfile

    out = []
    try:
        z = zipfile.ZipFile(path)
    except Exception as e:
        return ["cannot open %s: %s" % (os.path.basename(path), e)]
    try:
        names = z.namelist()
        hf = sorted(n for n in names
                    if re.match(r"word/(header|footer)\d*\.xml$", n))
        out.append("zip_entries=%d hf_parts=%s"
                   % (len(names), ",".join(x.split("/")[-1] for x in hf) or "(none)"))

        # 正文 sectPr：有没有 titlePg、引用了几个页眉页脚
        try:
            doc = z.read("word/document.xml").decode("utf-8", "replace")
            title_pg = len(re.findall(r"<w:titlePg\b", doc))
            href = len(re.findall(r"<w:headerReference\b", doc))
            fref = len(re.findall(r"<w:footerReference\b", doc))
            # 分节数
            nsect = len(re.findall(r"<w:sectPr\b", doc))
            out.append("sectPr=%d titlePg=%d headerRef=%d footerRef=%d"
                       % (nsect, title_pg, href, fref))
        except Exception as e:
            out.append("document.xml unreadable: %s" % e)

        for n in hf:
            try:
                x = z.read(n).decode("utf-8", "replace")
            except Exception:
                continue
            npara = len(re.findall(r"<w:p\b", x))
            # PAGE 域：instrText 里出现 PAGE
            npage = len(re.findall(r"<w:instrText[^>]*>\s*PAGE\b", x))
            # 每段的对齐与 tab
            jcs = []
            tabs = []
            texts = []
            for pm in re.finditer(r"<w:p\b.*?</w:p>", x, re.S):
                seg = pm.group(0)
                m = re.search(r'<w:jc w:val="([^"]+)"', seg)
                jcs.append(m.group(1) if m else "-")
                ntab = len(re.findall(r"<w:tab/>", seg))
                right_tab = 1 if re.search(r'<w:tab w:val="right"', seg) else 0
                tabs.append("%d%s" % (ntab, "R" if right_tab else ""))
                plain = re.sub(r"<[^>]+>", "", seg).strip()
                texts.append(plain[:24])
            out.append("%s paras=%d PAGE=%d jc=[%s] tabs=[%s] text=%s"
                       % (n.split("/")[-1], npara, npage,
                          ",".join(jcs), ",".join(tabs),
                          json.dumps(texts, ensure_ascii=False)))
    finally:
        try:
            z.close()
        except Exception:
            pass
    return out[:max_lines]


def diff_findings(expect, actual, limit: int = 12):
    """逐条列出差异，让人一眼看出"哪个检查项在这台机器上行为不同"。"""
    exp = {tuple(x) for x in expect}
    act = {tuple(x) for x in actual}
    missing = sorted(exp - act)   # 基准有、这台没有
    extra = sorted(act - exp)     # 这台多出来的
    lines = []
    for dim, sev, msg in missing[:limit]:
        lines.append("  - 缺失 [%s/%s] %s" % (dim, sev, msg))
    for dim, sev, msg in extra[:limit]:
        lines.append("  + 多出 [%s/%s] %s" % (dim, sev, msg))
    if len(missing) > limit:
        lines.append("  ... 还有 %d 条缺失未列出" % (len(missing) - limit))
    if len(extra) > limit:
        lines.append("  ... 还有 %d 条多出未列出" % (len(extra) - limit))
    return lines


def _page_field_parts(path: str) -> tuple:
    """统计「含 PAGE 域的页眉 part 数」与「页脚 part 数」。

    用来抓「两个页码」：页眉一处 + 页脚一处同时存在时，客户打印出来每页会有两个
    页码。这不是理论风险 —— 2026-11-15 实测过：Word「插入 → 页码」写的是
    ``w:fldSimple`` 简单域，而清理函数只遍历 ``w:r``（fldSimple 不是 run），
    于是页脚那条清不掉，回执却报"已清理 1 处"。
    多分节时页眉侧会有多个 part，都算一处，不误报。
    """
    import zipfile

    hdr = ftr = 0
    try:
        with zipfile.ZipFile(path) as z:
            for n in z.namelist():
                if not re.match(r"word/(header|footer)\d*\.xml$", n):
                    continue
                x = z.read(n).decode("utf-8", "replace")
                has = (re.search(r'<w:fldSimple[^>]*w:instr="[^"]*PAGE', x, re.I)
                       or re.search(r"<w:instrText[^>]*>\s*PAGE\b", x, re.I))
                if has:
                    if "/header" in n:
                        hdr += 1
                    else:
                        ftr += 1
    except Exception:
        pass
    return hdr, ftr


# ---------------------------------------------------------------------------
def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="Cross-platform functional equivalence check")
    ap.add_argument("--exe", required=True, help="受测可执行文件绝对路径")
    ap.add_argument("--label", default="unknown", help="平台标签，用于注解前缀")
    ap.add_argument("--spec", default=None, help="规范键，默认取基准文件里的")
    ap.add_argument("--baseline", default=None, help="基准 JSON，默认 tools/crossplat/baseline_apa_student.json")
    ap.add_argument("--fixture", default=None, help="输入 docx，默认 tools/crossplat/fixtures/bad_apa.docx")
    ap.add_argument("--workdir", default=None, help="工作目录，默认临时目录")
    ap.add_argument("--timeout", type=int, default=300, help="单次执行超时秒数")
    args = ap.parse_args(argv)

    here = os.path.dirname(os.path.abspath(__file__))
    baseline_path = args.baseline or os.path.join(here, "baseline_apa_student.json")
    # 素材来自 make_fixture.py：结构自洽（题名页 / 摘要 / 三级标题 / References 标题
    # + 6 条条目），格式故意全错。旧的 bad_apa.docx 有条目却没 References 标题，
    # 参考文献重排那条路径根本走不到，已被淘汰（仍留给 nuitka_repro 用）。
    fixture_path = args.fixture or os.path.join(here, "fixtures", "paper_apa_student.docx")

    tag = "[%s] " % args.label
    failures = []   # 硬失败（功能回归）
    soft = []       # 软提醒（基准过期等）

    exe = os.path.abspath(args.exe)
    if not os.path.isfile(exe):
        error(tag + "受测程序不存在：" + exe)
        return 2
    if not os.access(exe, os.X_OK):
        error(tag + "受测程序没有可执行权限：" + exe)
        return 2
    if not os.path.isfile(fixture_path):
        error(tag + "测试文档不存在：" + fixture_path)
        return 2

    fixture_sha = hashlib.sha256(open(fixture_path, "rb").read()).hexdigest()

    with open(baseline_path, encoding="utf-8") as f:
        baseline = json.load(f)

    spec = args.spec or baseline.get("spec") or "APA-Student"
    expect_version = str(baseline.get("app_version") or "")

    # 绝对路径：产物以 cwd=<workdir> 调起，相对的 workdir 会让参数里的
    # <workdir>/input.docx 叠成 <workdir>/<workdir>/input.docx。
    workdir = (os.path.abspath(args.workdir) if args.workdir
               else tempfile.mkdtemp(prefix="crossplat_"))
    os.makedirs(workdir, exist_ok=True)
    # 输入文档放在独立目录，避免 fix 的产物覆盖仓库里的基准文档
    doc_in = os.path.join(workdir, "input.docx")
    shutil.copyfile(fixture_path, doc_in)

    print("受测程序 : " + exe)
    print("平台标签 : " + args.label)
    print("规范     : " + spec)
    print("测试文档 : %s (sha256 %s)" % (fixture_path, fixture_sha[:16]))
    print("基准文件 : %s" % baseline_path)
    print("工作目录 : %s" % workdir)
    print("")

    if baseline.get("fixture_sha256") and baseline["fixture_sha256"] != fixture_sha:
        warn(tag + "测试文档与基准不一致（基准 %s…，当前 %s…），比对结果不可信"
             % (baseline["fixture_sha256"][:16], fixture_sha[:16]))

    # ---- 0. 产物环境自检 ----
    # 放在最前面：万一后面全线崩，这段也能说明「产物里的运行环境长什么样」。
    # 上一次失败就是因为只看到「页码没写」，却看不到「为什么没写」。
    step("0/6 产物环境自检（selfcheck）")
    r = run_exe(exe, ["selfcheck"], workdir, 180)
    print("rc=%s" % r.rc)
    raw = (r.out or "").strip()
    sc_payload = extract_json(raw)
    if sc_payload is None:
        print(raw[:1500] or (r.err or "")[:800])
    else:
        # 注解有长度上限，压成一行关键结论发出去
        brief = {
            "frozen": sc_payload.get("frozen"),
            "platform": sc_payload.get("platform"),
            "python": sc_payload.get("python"),
            "argv0": sc_payload.get("argv0"),
            "cwd": sc_payload.get("cwd"),
            "header_create": (str(sc_payload.get("header_create"))[:220]),
            "blank_doc": (str(sc_payload.get("blank_doc"))[:220]),
            "tpl_found": sc_payload.get("header_template_found"),
            "docx_boot": sc_payload.get("docx_boot"),
        }
        print(json.dumps(brief, ensure_ascii=False, indent=2)[:2500])
        notice(tag + "selfcheck | " + json.dumps(brief, ensure_ascii=False)[:900])
        if brief["header_create"] not in ("ok", "None"):
            failures.append(
                "产物里连一个空页眉都建不起来：%s" % str(brief["header_create"])[:400]
            )

    # ---- 1. 版本号 ----
    step("1/6 版本号")
    r = run_exe(exe, ["--version"], workdir, 120)
    ver_out = (r.out or r.err).strip()
    print(ver_out[:400] or "(无输出)")
    actual_version = ""
    m = re.search(r"(\d+\.\d+\.\d+)", ver_out)
    if m:
        actual_version = m.group(1)
    if not actual_version:
        failures.append("读不到版本号（rc=%s，超时=%s）" % (r.rc, r.timed_out))

    strict = bool(expect_version) and actual_version == expect_version
    if expect_version and actual_version and not strict:
        soft.append(
            "版本与基准不同（基准 %s，实测 %s）：本次只报告差异，不判红。"
            "发版后请刷新工具里的基准文件。" % (expect_version, actual_version)
        )
    print("基准版本 %s / 实测版本 %s -> %s"
          % (expect_version or "?", actual_version or "?", "严格比对" if strict else "宽松比对"))

    # ---- 2. 体检 ----
    step("2/6 体检（check）")
    r = run_exe(exe, ["check", doc_in, "--spec", spec, "--json"], workdir, args.timeout)
    print("rc=%s timed_out=%s" % (r.rc, r.timed_out))
    if r.err.strip():
        print("stderr: " + r.err.strip()[:800])
    before_payload = extract_json(r.out)
    if before_payload is None:
        failures.append(
            "check 没有返回可解析的 JSON（rc=%s，超时=%s）。stdout 前 500 字符：%s"
            % (r.rc, r.timed_out, (r.out or "")[:500])
        )
        before_sum, before_fp = None, None
    else:
        before_sum, before_fp = fingerprint(before_payload)
        print("summary: " + json.dumps(before_sum, ensure_ascii=False))
        notice(tag + "check 结果 %s（%d 条发现）"
               % (json.dumps(before_sum, ensure_ascii=False), len(before_fp)))

    # ---- 3. 与基准比对（修前）----
    step("3/6 与基准比对（修前）")
    if before_fp is not None:
        exp_sum = baseline.get("before", {}).get("summary") or {}
        exp_fp = baseline.get("before", {}).get("findings") or []
        if before_sum == exp_sum:
            print("summary 一致 ✓")
        else:
            msg = "summary 不一致：基准 %s，实测 %s" % (
                json.dumps(exp_sum, ensure_ascii=False),
                json.dumps(before_sum, ensure_ascii=False),
            )
            print(msg)
            (failures if strict else soft).append(msg)
        if before_fp == [list(x) for x in exp_fp] or before_fp == exp_fp:
            print("findings 逐条一致 ✓（%d 条）" % len(before_fp))
        else:
            lines = diff_findings(exp_fp, before_fp)
            print("findings 不一致：")
            print("\n".join(lines))
            msg = "修前 findings 与基准有差异（基准 %d 条 / 实测 %d 条）" % (
                len(exp_fp), len(before_fp))
            (failures if strict else soft).append(msg + "\n" + "\n".join(lines[:6]))

    # ---- 4. 预览可修项 ----
    # preview 不落盘、不消耗试用，能安全地把「这台机器认为哪些维度可修」暴露出来。
    # 为什么值得单独一步：findings 的指纹只含 dimension/severity/message，
    # 不含「有没有绑修复动作」这一层。曾出现两平台 findings 34 条逐条一致、
    # 但 macOS 上就是不肯写页码的情况 —— 差异恰好藏在这层里。
    step("4/6 预览可修项（preview）")
    r = run_exe(exe, ["fix", doc_in, "--spec", spec, "--preview", "--json"],
                workdir, args.timeout)
    print("rc=%s timed_out=%s" % (r.rc, r.timed_out))
    preview_changes = None
    preview_payload = extract_json(r.out)
    if preview_payload is None:
        print("preview 没有返回 JSON（rc=%s）" % r.rc)
        if r.err.strip():
            print("stderr: " + r.err.strip()[:600])
    else:
        preview_changes = preview_payload.get("changes") or []
        print("可修 %d 项：" % len(preview_changes))
        for c in preview_changes:
            print("  - " + str(c))
        notice(tag + "可修 %d 项：%s"
               % (len(preview_changes),
                  " | ".join(str(c) for c in preview_changes))[:900])
        has_pn = any("page number" in str(c).lower() for c in preview_changes)
        if not has_pn:
            msg = ("preview 的可修项里没有 Page number —— 说明这台机器上 "
                   "checker 没有给页码绑修复动作（fix_dims 缺 page_number）")
            print(msg)
            (failures if strict else soft).append(msg)
        else:
            print("可修项含 Page number ✓")
        # 与基准记录的可修项数量对照（数量变了说明修复范围变了）
        exp_n = baseline.get("preview_change_count")
        if isinstance(exp_n, int) and exp_n != len(preview_changes):
            msg = "可修项数量与基准不同（基准 %d 项 / 实测 %d 项）" % (
                exp_n, len(preview_changes))
            print(msg)
            (failures if strict else soft).append(msg)

    # ---- 5. 一键修正 ----
    step("5/6 一键修正（fix）")
    doc_out = os.path.join(workdir, "fixed.docx")
    r = run_exe(
        exe,
        ["fix", doc_in, "--spec", spec, "--out", doc_out, "--json"],
        workdir,
        args.timeout,
    )
    print("rc=%s timed_out=%s" % (r.rc, r.timed_out))
    if r.err.strip():
        print("stderr: " + r.err.strip()[:800])
    fix_payload = extract_json(r.out)
    fix_ok = False
    if fix_payload is None:
        failures.append(
            "fix 没有返回可解析的 JSON（rc=%s，超时=%s）。stdout 前 500 字符：%s"
            % (r.rc, r.timed_out, (r.out or "")[:500])
        )
    else:
        fix_ok = bool(fix_payload.get("ok"))
        print("ok=%s output=%s" % (fix_ok, fix_payload.get("output")))
        if not fix_ok:
            failures.append("fix 返回 ok=false：%s"
                            % json.dumps(fix_payload, ensure_ascii=False)[:400])
    produced = os.path.isfile(doc_out)
    if not produced:
        # 有的实现会改写 output 字段里的路径
        cand = (fix_payload or {}).get("output")
        if cand and os.path.isfile(os.path.join(workdir, cand)):
            doc_out = os.path.join(workdir, cand)
            produced = True
    if not produced:
        failures.append("fix 没有生成修正稿文件（期望 %s）" % doc_out)
        print("找不到修正稿，跳过复检")
        return finish(failures, soft, tag)
    print("修正稿 %.1f KiB" % (os.path.getsize(doc_out) / 1024.0))

    # ---- 5b. fix 回执里「真干了哪些活」必须对得上 ----
    # 为什么单独查这一层：2.3.122 的 mac 版 fix 返回 ok=true、复检报告也变好了，
    # 唯独 page_number 这一步根本没执行（页眉 part 没创建）。只看 ok 和 summary
    # 一辈子发现不了。applied 是修正器自己的作业回执，拿它当硬断言。
    # 期望值来自基准文件，不写死在代码里（换规范只需换基准）。
    if isinstance(fix_payload, dict):
        fx_sum = fix_payload.get("summary") or {}
        applied = fx_sum.get("applied")
        n_comments = fx_sum.get("comments_added")
        print("fix.applied=%s" % json.dumps(applied, ensure_ascii=False)[:500])
        print("fix.comments_added=%s" % n_comments)

        # 被修正器兜住的异常：兜住可以，沉默不行。
        soft_errs = fx_sum.get("soft_errors")
        if soft_errs:
            print("fix.soft_errors（被兜住的异常）:")
            for s in soft_errs[:10]:
                print("  ! " + str(s))
            failures.append(
                "修正过程中有 %d 个异常被悄悄兜住（改不动的项不会告诉客户）：%s"
                % (len(soft_errs), " | ".join(str(s) for s in soft_errs[:4]))
            )
        else:
            print("无被兜住的异常 ✓")

        for key in (baseline.get("required_applied") or []):
            ok_key = isinstance(applied, dict) and applied.get(key)
            if ok_key:
                print("applied 含 %s ✓" % key)
            else:
                failures.append(
                    "fix 回执的 applied 里没有 %s —— 修正器没执行这一步"
                    "（applied=%s）。这类『报告说改了、实际没改』的静默失效，"
                    "多半是打包后读不到自带资源。"
                    % (key, json.dumps(applied, ensure_ascii=False)[:300])
                )

        min_c = baseline.get("min_comments_added")
        if isinstance(min_c, int):
            try:
                got = int(n_comments)
            except (TypeError, ValueError):
                got = -1
            if got >= min_c:
                print("批注写入 %d 条 ✓（≥%d）" % (got, min_c))
            else:
                failures.append(
                    "fix 只写了 %s 条批注（基准要求 ≥%d）—— 批注部件没建起来，"
                    "客户拿到的修正稿里看不到任何改动说明。" % (n_comments, min_c)
                )

    # ---- 5c. 自带数据文件在产物里能不能定位到 ----
    # 这次 bug 的本质是「打包后 __file__ 变相对路径 → 找不到自带资源」。
    # 页码只是碰巧第一个暴露的；export-schema 走的是 src/data 下的资源，
    # 用它可以低成本验证这条通路在所有平台都通。
    step("5c 自带资源定位（export-schema）")
    schema_out = os.path.join(workdir, "schema.yaml")
    r = run_exe(exe, ["export-schema", schema_out], workdir, 120)
    print("rc=%s" % r.rc)
    if (r.out or r.err).strip():
        print((r.out or r.err).strip()[:400])
    size = os.path.getsize(schema_out) if os.path.isfile(schema_out) else 0
    if size > 1000:
        print("导出模板 %d 字节 ✓" % size)
    else:
        failures.append(
            "export-schema 没导出内置模板（rc=%s，文件 %d 字节）—— "
            "产物里定位不到自带数据文件，同一类冻结路径问题。" % (r.rc, size)
        )

    # ---- 5. 复检 ----
    step("6/6 复检修正稿（check again）")
    r = run_exe(exe, ["check", doc_out, "--spec", spec, "--json"], workdir, args.timeout)
    print("rc=%s timed_out=%s" % (r.rc, r.timed_out))
    after_payload = extract_json(r.out)
    if after_payload is None:
        failures.append(
            "复检没有返回可解析的 JSON（rc=%s，超时=%s）。stdout 前 500 字符：%s"
            % (r.rc, r.timed_out, (r.out or "")[:500])
        )
        return finish(failures, soft, tag)

    after_sum, after_fp = fingerprint(after_payload)
    print("summary: " + json.dumps(after_sum, ensure_ascii=False))
    notice(tag + "复检结果 %s（%d 条发现）"
           % (json.dumps(after_sum, ensure_ascii=False), len(after_fp)))

    exp_sum = baseline.get("after", {}).get("summary") or {}
    exp_fp = baseline.get("after", {}).get("findings") or []
    if after_sum == exp_sum:
        print("summary 一致 ✓")
    else:
        msg = "复检 summary 不一致：基准 %s，实测 %s" % (
            json.dumps(exp_sum, ensure_ascii=False),
            json.dumps(after_sum, ensure_ascii=False),
        )
        print(msg)
        (failures if strict else soft).append(msg)
    if after_fp == [list(x) for x in exp_fp] or after_fp == exp_fp:
        print("findings 逐条一致 ✓（%d 条）" % len(after_fp))
    else:
        lines = diff_findings(exp_fp, after_fp)
        print("findings 不一致：")
        print("\n".join(lines))
        msg = "复检 findings 与基准有差异（基准 %d 条 / 实测 %d 条）" % (
            len(exp_fp), len(after_fp))
        (failures if strict else soft).append(msg + "\n" + "\n".join(lines[:6]))

    # 额外自检：修完不应该比修前更差
    try:
        b_fail = int(before_sum.get("fail", 0)) + int(before_sum.get("warn", 0))
        a_fail = int(after_sum.get("fail", 0)) + int(after_sum.get("warn", 0))
        if a_fail > b_fail:
            failures.append(
                "修正后问题数反而增加（修前 warn+fail=%d，修后=%d）—— 修正器在自伤"
                % (b_fail, a_fail))
        else:
            print("修正效果：warn+fail %d -> %d ✓" % (b_fail, a_fail))
    except Exception:
        pass

    # ---- 6b. 客户视角：每页只能有一个页码 ----
    step("6b 页码数量（每页只许一处）")
    hdr_n, ftr_n = _page_field_parts(doc_out)
    print("含 PAGE 域的 part：页眉 %d 个 / 页脚 %d 个" % (hdr_n, ftr_n))
    if hdr_n and ftr_n:
        failures.append(
            "页眉和页脚同时有页码（页眉 %d 处 / 页脚 %d 处）—— "
            "客户打印出来每页会有两个页码。多半是清页脚那条时漏了 fldSimple 简单域。"
            % (hdr_n, ftr_n)
        )
    elif not hdr_n and not ftr_n:
        failures.append("修正稿里找不到任何页码（页眉页脚都没有 PAGE 域）")
    else:
        print("页码只在一处 ✓")

    # ---- 无论通过与否，都发一条修正稿结构证据 ----    # 为什么通过时也要发：summary 从 warn 变 pass 只说明「报告说修好了」，
    # 而 2.3.122 的 mac 版恰恰是「报告说页码已改、文档里连页眉都没创建」。
    # 唯一能证明修好了的证据，是修正稿里真的存在带 PAGE 域的 header part。
    try:
        struct = diagnose_docx(doc_out, max_lines=6)
        if struct:
            notice(tag + "修正稿结构 | " + " ; ".join(struct)[:900])
    except Exception as e:  # 取证失败不影响判定
        print("struct evidence failed: %s" % e)

    # ---- 不一致时取证：把修正稿的结构摊平，随注解一起发出去 ----
    if failures:
        step("取证")
        # ① fix 回执里到底认了什么活（applied 为空 = 修正器根本没动手）
        if isinstance(fix_payload, dict):
            applied = (fix_payload.get("summary") or {}).get("applied")
            line = "fix.applied=" + json.dumps(applied, ensure_ascii=False)
            print(line[:600])
            error(tag + "诊断 | " + line[:600])
            if isinstance(applied, dict) and "page_number" not in applied:
                error(tag + "诊断 | fix 回执里没有 page_number —— 修正器没执行页码这一步")
        # ② preview 的可修项
        if preview_changes is not None:
            line = "preview=%d项 %s" % (len(preview_changes),
                                       " | ".join(str(c) for c in preview_changes))
            error(tag + "诊断 | " + line[:700])
        # ③ 修正稿的页眉页脚结构
        for line in diagnose_docx(doc_out):
            print(line)
            error(tag + "诊断 | " + line)
        # 修前文档的结构也打一份，便于对照"本来什么样 / 变成什么样"
        try:
            for line in diagnose_docx(doc_in):
                print("[输入] " + line)
        except Exception:
            pass

    return finish(failures, soft, tag)


def finish(failures, soft, tag) -> int:
    step("结论")
    for s in soft:
        warn(tag + s)
    if failures:
        for msg in failures:
            error(tag + msg)
        print("\n结论：❌ 该平台产物存在功能性回归")
        return 1
    notice(tag + "✅ 打包产物功能与基准完全一致（check / fix / 复检三段全通）")
    print("\n结论：✅ 通过")
    return 0


if __name__ == "__main__":
    sys.exit(main())
