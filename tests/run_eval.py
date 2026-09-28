# -*- coding: utf-8 -*-
"""run_eval.py — 批量评测：检测准确率 / verify 通过率 / 内容零损坏

用法：python run_eval.py [阶段] [场景名 ...]
  阶段 = extract（默认）| full（提取+应用+验证，M2 后可用）
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
FIXDIR = os.path.join(HERE, "fixtures")
TOL = 0.06  # cm/pt 容差


def resolve(obj, path: str):
    cur = obj
    for part in path.split("."):
        if part.startswith("[") and part.endswith("]"):
            cur = cur[int(part[1:-1])]
        elif isinstance(cur, dict):
            cur = cur.get(part)
        else:
            return None
        if cur is None:
            return None
    return cur


def match(actual, expected) -> bool:
    if isinstance(expected, dict) and "op" in expected:
        op, v = expected["op"], expected["value"]
        try:
            if op == ">=":
                return actual is not None and actual >= v
            if op == "<=":
                return actual is not None and actual <= v
            if op == ">":
                return actual is not None and actual > v
            if op == "<":
                return actual is not None and actual < v
            if op == "in":
                return actual is not None and v in actual
        except TypeError:
            return False
    if isinstance(actual, bool) or isinstance(expected, bool):
        return actual is expected or actual == expected
    if isinstance(expected, (int, float)) and isinstance(actual, (int, float)):
        return abs(actual - expected) <= TOL
    return actual == expected


def eval_extract(name: str):
    d = os.path.join(FIXDIR, name)
    tpl = os.path.join(d, "template.docx")
    prof_path = os.path.join(d, "template.profile.json")
    r = subprocess.run([sys.executable, os.path.join(SCRIPTS, "extract_format.py"),
                        tpl, "-o", prof_path, "--report", os.path.join(d, "template.report.md"),
                        "--map", os.path.join(d, "template.map.json")],
                       capture_output=True, text=True, encoding="utf-8")
    if r.returncode != 0:
        return 0, 1, [(f"__error__", f"提取器崩溃: {r.stderr.strip()[-300:]}")]
    prof = json.load(open(prof_path, encoding="utf-8"))
    expected = json.load(open(os.path.join(d, "expected.json"), encoding="utf-8"))
    fails = []
    for path, exp in expected.items():
        act = resolve(prof, path)
        if not match(act, exp):
            fails.append((path, f"期望 {exp!r}，实际 {act!r}"))
    return len(expected) - len(fails), len(expected), fails


def eval_full(name: str):
    d = os.path.join(FIXDIR, name)
    ok, n, fails = eval_extract(name)
    if fails:
        return ok, n, fails, ("extract", 0, 0)
    out_docx = os.path.join(d, "target.formatted.docx")
    apply_cmd = [sys.executable, os.path.join(SCRIPTS, "apply_format.py"),
                 os.path.join(d, "target.docx"),
                 os.path.join(d, "template.profile.json"),
                 "-o", out_docx]
    map_override = os.path.join(d, "map_override.json")
    if os.path.exists(map_override):
        apply_cmd += ["--map", map_override]
    r = subprocess.run(apply_cmd, capture_output=True, text=True, encoding="utf-8")
    if r.returncode != 0:
        return ok, n, [(f"__error__", f"应用器崩溃: {r.stderr.strip()[-300:]}")], \
            ("apply", 0, 0)
    v = subprocess.run([sys.executable, os.path.join(SCRIPTS, "verify_format.py"),
                        out_docx, os.path.join(d, "template.profile.json"),
                        "--original", os.path.join(d, "target.docx")],
                       capture_output=True, text=True, encoding="utf-8")
    m = re.search(r"PASS (\d+) / WARN (\d+) / FAIL (\d+)", v.stdout)
    stats = tuple(int(x) for x in m.groups()) if m else (0, 0, -1)
    vfails = []
    if v.returncode != 0:
        for line in v.stdout.splitlines():
            if line.strip().startswith("FAIL"):
                vfails.append(("verify:" + line.split("FAIL", 1)[1][:80], ""))
        if not vfails:
            vfails.append(("verify", v.stdout.strip()[-200:] or "验证失败"))
    return ok, n, vfails, stats


def main():
    stage = "extract"
    args = sys.argv[1:]
    if args and args[0] in ("extract", "full"):
        stage = args.pop(0)
    names = args or sorted(n for n in os.listdir(FIXDIR)
                           if os.path.isdir(os.path.join(FIXDIR, n))
                           and os.path.isfile(os.path.join(FIXDIR, n, "template.docx")))
    total_ok = total_n = 0
    any_fail = False
    for name in names:
        stats = None
        if stage == "extract":
            ok, n, fails = eval_extract(name)
        else:
            ok, n, fails, stats = eval_full(name)
        total_ok += ok
        total_n += n
        status = "PASS" if not fails else f"FAIL ({ok}/{n})"
        extra = f" · verify P/W/F={stats[0]}/{stats[1]}/{stats[2]}" if stats else ""
        print(f"[{status}] {name}{extra}")
        for path, msg in fails:
            any_fail = True
            print(f"    - {path}: {msg}")
    print(f"\n检测准确率: {total_ok}/{total_n}"
          f"（{round(100 * total_ok / max(total_n, 1))}%）")
    sys.exit(1 if any_fail else 0)


if __name__ == "__main__":
    main()
