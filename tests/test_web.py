# -*- coding: utf-8 -*-
"""test_web.py — 网页版 20 轮端到端测试

每轮：上传不同文档 → 流式检测（断言 13 步）→ 逐条执行（轮换条件变体）→
断言落盘 + verify 0 FAIL + 输出可打开。
文档池 = 10 个合成夹具（template+target 对）+ 10 份 AI 风格报告（配华科模板）。
"""
from __future__ import annotations

import io
import json
import os
import random
import subprocess
import sys
import time
import urllib.request

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
SCRIPTS = os.path.join(ROOT, "skill", "scripts")
FIX = os.path.join(HERE, "fixtures")
AI = os.path.join(HERE, "ai_reports")

PORT = 8900 + random.randint(0, 99)
BASE = f"http://127.0.0.1:{PORT}"

RULE_STAGES = {"page_size": 1, "page_margins": 1, "pagenum_format": 1, "pagenum_pos": 1,
               "style_body": 2, "style_h1": 2, "style_h2": 2, "style_h3": 2,
               "keep_captions": 4, "keep_figures": 4, "keep_headings": 4, "toc": 5}


def stage_of(rid, default=3):
    return RULE_STAGES.get(rid, default)


def pool():
    cases = []
    for name in ("thesis_standard", "mathmodel_manual", "journal_colon",
                 "chinese_numbering", "code_appendix_only", "continuation_table",
                 "floating_anchor", "formula_display", "gutter_margins",
                 "many_figures"):
        d = os.path.join(FIX, name)
        cases.append((f"fixture:{name}", os.path.join(d, "template.docx"),
                      os.path.join(d, "target.docx")))
    tplB = os.path.join(AI, "templateB.docx")
    for name in ("v2_chapter_manual", "v2_auto_num", "v2_english_cap",
                 "v2_caption_above", "v2_no_caption", "v2_fullwidth_num",
                 "v2_tight", "v2_formula_heavy", "v2_bold_body", "v2_mixed"):
        cases.append((f"ai:{name}", tplB, os.path.join(AI, "round2", f"{name}.docx")))
    return cases


def variant(i, order):
    """轮换执行条件变体 → (order, disabled, overrides, 标签)"""
    v = i % 6
    if v == 0:
        return order, [], {}, "默认顺序"
    if v == 1:
        return order, ["toc", "formula"], {}, "禁用 目录+公式"
    if v == 2:
        return order, [], {"page.margins.top_cm": 2.6}, "覆盖 上边距2.6"
    if v == 3:
        # 阶段内打乱（保持依赖组）
        rng = random.Random(i)
        groups = {}
        for rid in order:
            groups.setdefault(stage_of(rid), []).append(rid)
        mixed = []
        for st in sorted(groups):
            g = groups[st][:]
            rng.shuffle(g)
            mixed += g
        return mixed, [], {}, "阶段内乱序"
    if v == 4:
        return order, [], {}, "全选"
    return order, ["keep_captions", "keep_figures", "keep_headings"], {}, "禁用全部分页类"


def ndjson(resp):
    for line in resp.read().decode("utf-8").splitlines():
        if line.strip():
            yield json.loads(line)


def main():
    proc = subprocess.Popen([sys.executable, os.path.join(SCRIPTS, "serve_ui.py"),
                             str(PORT)],
                            stdout=subprocess.DEVNULL, stderr=subprocess.STDOUT)
    failures = []
    try:
        for _ in range(60):
            time.sleep(0.25)
            try:
                ver = json.load(urllib.request.urlopen(BASE + "/api/version", timeout=2))
                break
            except Exception:
                pass
        else:
            raise SystemExit("服务器未启动")
        print(f"网页版 v{ver['version']}（{ver['rules']} 条规则）\n")

        def post(path, data=None, raw=None):
            if raw is not None:
                req = urllib.request.Request(BASE + path, data=raw, method="POST")
            else:
                req = urllib.request.Request(
                    BASE + path, data=json.dumps(data or {}).encode(),
                    headers={"Content-Type": "application/json"}, method="POST")
            return urllib.request.urlopen(req, timeout=120)

        rules = json.load(urllib.request.urlopen(BASE + "/api/rules"))["default_order"]
        ok_rounds = 0
        for i, (tag, tpl, tgt) in enumerate(pool(), 1):
            errs = []
            try:
                # 上传
                r1 = json.load(post("/api/upload?name=w_tpl.docx",
                                    raw=open(tpl, "rb").read()))
                r2 = json.load(post("/api/upload?name=w_tgt.docx",
                                    raw=open(tgt, "rb").read()))
                # 流式检测
                steps, profile = [], None
                for ev in ndjson(post("/api/detect", {"path": r1["path"]})):
                    if ev["type"] == "step":
                        steps.append(ev["key"])
                    elif ev["type"] == "done":
                        profile = ev["profile"]
                    elif ev["type"] == "error":
                        errs.append("detect:" + ev["message"][:120])
                if len(steps) != 13:
                    errs.append(f"检测步数 {len(steps)} ≠ 13")
                # 逐条执行
                order, disabled, overrides, vlabel = variant(i, rules)
                saved, verify = None, None
                rule_events = 0
                for ev in ndjson(post("/api/apply", {
                        "target": r2["path"], "profile": profile,
                        "order": order, "disabled": disabled,
                        "overrides": overrides})):
                    if ev["type"] == "rule" and ev.get("status") in ("done", "skipped"):
                        rule_events += 1
                        if ev["status"] == "error":
                            errs.append(f"规则 {ev['id']} 失败")
                    elif ev["type"] == "saved":
                        saved = ev
                    elif ev["type"] == "verify":
                        verify = ev
                    elif ev["type"] == "error":
                        errs.append("apply:" + ev["message"][:120])
                if saved is None or not os.path.isfile(saved["output"]):
                    errs.append("输出未落盘")
                if verify is None or verify["fail"] != 0:
                    errs.append(f"verify 失败数 {verify and verify['fail']}")
                if rule_events != len(order):
                    errs.append(f"规则事件 {rule_events} ≠ {len(order)}")
                if not errs:
                    from docx import Document
                    Document(saved["output"])
                status = "PASS" if not errs else "FAIL"
                if errs:
                    failures.append((tag, errs))
                print(f"[{status}] 轮{i:>2} {tag:<28} {vlabel:<12} "
                      f"改{saved and saved['changes'] or 0:>3} 处 "
                      f"验证✅{verify and verify['pass']} ❌{verify and verify['fail']}")
                for e in errs:
                    print(f"     - {e}")
                if not errs:
                    ok_rounds += 1
            except Exception as e:
                failures.append((tag, [f"{type(e).__name__}: {e}"]))
                print(f"[FAIL] 轮{i:>2} {tag}: {type(e).__name__}: {e}")
        print(f"\n网页版：{ok_rounds}/20 轮通过")
        sys.exit(1 if failures else 0)
    finally:
        proc.terminate()


if __name__ == "__main__":
    main()
