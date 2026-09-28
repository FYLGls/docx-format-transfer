# -*- coding: utf-8 -*-
"""run_ai_eval.py — 真实模板 × AI 风格报告 全流程评测

用法：python run_ai_eval.py [round1|round2|all]
"""
from __future__ import annotations

import io
import json
import os
import re
import subprocess
import sys

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
SCRIPTS = os.path.join(ROOT, "skill", "scripts")
AI = os.path.join(HERE, "ai_reports")

TEMPLATES = {
    "round1": os.path.join(AI, "templateA.profile.json"),
    "round2": os.path.join(AI, "templateB.profile.json"),
}


def run_one(round_name, report_path, workdir):
    prof = TEMPLATES[round_name]
    out_docx = os.path.join(workdir, os.path.basename(report_path).replace(".docx", ".formatted.docx"))
    issues = []
    # 1) 目标提取（元素映射 + 歧义）
    r = subprocess.run([sys.executable, os.path.join(SCRIPTS, "extract_format.py"),
                        report_path, "-o", out_docx + ".profile.json",
                        "--report", out_docx + ".report.md",
                        "--map", out_docx + ".map.json"],
                       capture_output=True, text=True, encoding="utf-8")
    if r.returncode != 0:
        return ["EXTRACT-CRASH: " + r.stderr.strip()[-200:]], 0, 0
    emap = json.load(open(out_docx + ".map.json", encoding="utf-8"))
    amb = len(emap.get("ambiguities") or [])
    # 2) 应用
    r = subprocess.run([sys.executable, os.path.join(SCRIPTS, "apply_format.py"),
                        report_path, prof, "-o", out_docx,
                        "--map", out_docx + ".map.json"],
                       capture_output=True, text=True, encoding="utf-8")
    if r.returncode != 0:
        return [f"APPLY-CRASH: {r.stderr.strip()[-200:]}"], amb, 0
    # 3) 验证
    v = subprocess.run([sys.executable, os.path.join(SCRIPTS, "verify_format.py"),
                        out_docx, prof, "--original", report_path],
                       capture_output=True, text=True, encoding="utf-8")
    m = re.search(r"PASS (\d+) / WARN (\d+) / FAIL (\d+)", v.stdout)
    p, w_, f = (int(x) for x in m.groups()) if m else (0, 0, -1)
    if v.returncode != 0:
        for line in v.stdout.splitlines():
            if line.strip().startswith("FAIL"):
                issues.append("VERIFY " + line.split("FAIL", 1)[1].strip()[:120])
        if f <= 0:
            issues.append("VERIFY-EXIT1: " + v.stdout.strip()[-200:])
    # 4) 可打开性
    try:
        from docx import Document
        Document(out_docx)
    except Exception as e:
        issues.append(f"OPEN-FAIL: {e}")
    return issues, amb, (p, w_, f)


def main():
    rounds = sys.argv[1:] or ["round1"]
    all_issues = {}
    for rnd in (rounds if rounds != ["all"] else ["round1", "round2"]):
        folder = os.path.join(AI, rnd)
        names = sorted(n for n in os.listdir(folder) if n.endswith(".docx")
                       and "formatted" not in n)
        print(f"===== {rnd}（{len(names)} 份报告）=====")
        for n in names:
            issues, amb, st = run_one(rnd, os.path.join(folder, n), folder)
            tag = "OK " if not issues else "ISS"
            stat = f"P/W/F={st[0]}/{st[1]}/{st[2]}" if isinstance(st, tuple) else st
            print(f"[{tag}] {n:<28} 歧义{amb:>2} {stat}")
            if issues:
                all_issues[f"{rnd}/{n}"] = issues
                for it in issues:
                    print(f"       - {it}")
    out = os.path.join(AI, "ISSUES-raw.json")
    json.dump(all_issues, open(out, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print(f"\n共 {len(all_issues)} 份报告有问题，明细 -> {out}")


if __name__ == "__main__":
    main()
