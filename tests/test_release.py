# -*- coding: utf-8 -*-
"""test_release.py — 可下载发布包 20 轮安装-运行测试

每轮：zip 解压到全新临时目录（模拟用户下载解压）→ check_deps → 从解压目录直接跑
CLI 全流程（extract → apply 轮换变体 → verify）→ 断言 0 FAIL + 输出可打开。
文档池与网页版互补：另 10 个合成夹具 + 10 份 AI round1 报告（配课程模板）。
"""
from __future__ import annotations

import io
import json
import os
import random
import shutil
import subprocess
import sys
import tempfile
import zipfile

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
FIX = os.path.join(HERE, "fixtures")
AI = os.path.join(HERE, "ai_reports")
PY = sys.executable

RULE_STAGES = {"page_size": 1, "page_margins": 1, "pagenum_format": 1, "pagenum_pos": 1,
               "style_body": 2, "style_h1": 2, "style_h2": 2, "style_h3": 2,
               "keep_captions": 4, "keep_figures": 4, "keep_headings": 4, "toc": 5}


def run(cmd, cwd=None, timeout=120):
    env = dict(os.environ, PYTHONIOENCODING="utf-8")
    return subprocess.run(cmd, capture_output=True, text=True,
                          encoding="utf-8", errors="replace", cwd=cwd,
                          timeout=timeout, env=env)


def pool():
    cases = []
    for name in ("reverse_field_manual", "unstyled_target", "landscape_mixed",
                 "merged_cells", "target_extra_level", "template_inconsistent",
                 "no_formulas_template", "xref_dense", "seq_continuous_fields",
                 "outline_lvl_headings"):
        d = os.path.join(FIX, name)
        cases.append((f"fixture:{name}", os.path.join(d, "template.docx"),
                      os.path.join(d, "target.docx")))
    tplA = r"G:\数据挖掘\实验一\实验一\实验1_报告模板.docx"
    if not os.path.isfile(tplA):
        tplA = os.path.join(AI, "templateA.docx")
    for name in ("docxjs", "pydocx_raw", "pandoc", "wps_ai", "word_web",
                 "chatgpt", "tongyi", "latex2docx", "md_paste", "manual_ok"):
        cases.append((f"ai:{name}", tplA, os.path.join(AI, "round1", f"{name}.docx")))
    return cases


def order_from_package(pkg_scripts):
    r = run([PY, "-c",
             "import json,sys; sys.path.insert(0,'.');"
             "from apply_format import DEFAULT_ORDER; print(json.dumps(DEFAULT_ORDER))"],
            cwd=pkg_scripts)
    return json.loads(r.stdout)


def variant(i, order, workdir):
    v = i % 5
    if v == 0:
        return None, None, "默认"
    if v == 1:
        rules = {"order": order, "disabled": ["toc", "formula"]}
        p = os.path.join(workdir, f"rules{i}.json")
        json.dump(rules, open(p, "w", encoding="utf-8"))
        return p, None, "禁用 目录+公式"
    if v == 2:
        rng = random.Random(100 + i)
        groups = {}
        for rid in order:
            groups.setdefault(RULE_STAGES.get(rid, 3), []).append(rid)
        mixed = []
        for st in sorted(groups):
            g = groups[st][:]
            rng.shuffle(g)
            mixed += g
        p = os.path.join(workdir, f"rules{i}.json")
        json.dump(mixed, open(p, "w", encoding="utf-8"))
        return p, None, "阶段内乱序"
    if v == 3:
        rules = {"order": order, "disabled": ["keep_captions", "keep_figures",
                                              "keep_headings"]}
        p = os.path.join(workdir, f"rules{i}.json")
        json.dump(rules, open(p, "w", encoding="utf-8"))
        return p, None, "禁用全部分页类"
    return None, None, "默认"


def main():
    # 构建发布包
    r = run([PY, os.path.join(ROOT, "packaging", "build_release.py")])
    assert r.returncode == 0, "构建失败:\n" + r.stderr[-400:]
    version = open(os.path.join(ROOT, "VERSION"), encoding="utf-8").read().strip()
    zip_path = os.path.join(ROOT, "release", f"docx-format-transfer-v{version}.zip")
    print(f"发布包：{zip_path}\n")
    top = f"docx-format-transfer-v{version}"

    base_tmp = tempfile.mkdtemp(prefix="dft_release_")
    failures = []
    ok_rounds = 0
    try:
        for i, (tag, tpl, tgt) in enumerate(pool(), 1):
            errs = []
            work = os.path.join(base_tmp, f"round{i:02d}")
            os.makedirs(work)
            # 全新解压
            with zipfile.ZipFile(zip_path) as z:
                z.extractall(work)
            pkg = os.path.join(work, top)
            pkg_scripts = os.path.join(pkg, "scripts")
            # 依赖体检
            r = run([PY, "check_deps.py"], cwd=pkg)
            if r.returncode != 0:
                errs.append("check_deps: " + (r.stdout + r.stderr)[-150:])
            order = order_from_package(pkg_scripts)
            rules_json, _, vlabel = variant(i, order, work)
            # extract 模板
            prof = os.path.join(work, "tpl.profile.json")
            r = run([PY, "extract_format.py", tpl, "-o", prof,
                     "--report", prof + ".md", "--map", prof + ".map.json"],
                    cwd=pkg_scripts)
            if r.returncode != 0:
                errs.append("extract: " + r.stderr[-150:])
                failures.append((tag, errs))
                print(f"[FAIL] 轮{i:>2} {tag}: extract")
                continue
            # apply（轮换变体）
            out = os.path.join(work, "out.formatted.docx")
            cmd = [PY, "apply_format.py", tgt, prof, "-o", out]
            if rules_json:
                cmd += ["--rules", rules_json]
            r = run(cmd, cwd=pkg_scripts)
            if r.returncode != 0:
                errs.append("apply: " + r.stderr[-150:])
            # verify
            vline = ""
            if not errs:
                r = run([PY, "verify_format.py", out, prof, "--original", tgt],
                        cwd=pkg_scripts)
                if r.returncode != 0:
                    errs.append("verify: " + r.stdout.strip()[-200:])
                else:
                    vline = r.stdout.strip().splitlines()[0]
                    from docx import Document
                    Document(out)
            status = "PASS" if not errs else "FAIL"
            if errs:
                failures.append((tag, errs))
            else:
                ok_rounds += 1
            print(f"[{status}] 轮{i:>2} {tag:<30} {vlabel:<12} {vline[8:].strip() if vline else ''}")
            for e in errs:
                print(f"     - {e}")

        # install-skill.bat 逻辑验证（装到临时目录，不动真实 skills 目录）
        inst_dest = os.path.join(base_tmp, "skills", "docx-format-transfer")
        r = run(["cmd", "/c", "install-skill.bat", inst_dest], cwd=os.path.join(base_tmp, "round01", top))
        inst_ok = r.returncode == 0 and os.path.isfile(
            os.path.join(inst_dest, "SKILL.md")) and \
            os.path.isfile(os.path.join(inst_dest, "scripts", "serve_ui.py"))
        print(f"[{'PASS' if inst_ok else 'FAIL'}] 附加  install-skill.bat → 临时目录注册"
              + ("" if inst_ok else f"（rc={r.returncode}）"))
        if not inst_ok:
            failures.append(("install", [(r.stdout or "")[-200:] + (r.stderr or "")[-200:]]))
        else:
            ok_rounds += 0  # 不计入 20 轮，仅附加验证
        print(f"\n下载包：{ok_rounds}/20 轮通过（另有 install-skill 验证）")
        sys.exit(1 if failures else 0)
    finally:
        shutil.rmtree(base_tmp, ignore_errors=True)


if __name__ == "__main__":
    main()
