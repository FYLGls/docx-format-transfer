# -*- coding: utf-8 -*-
"""smoke_ui.py — 界面服务器全链路烟雾测试"""
import io
import json
import os
import subprocess
import sys
import time
import urllib.request

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
SCRIPTS = os.path.join(ROOT, "skill", "scripts")
PORT = 8899
BASE = f"http://127.0.0.1:{PORT}"
FIX = os.path.join(HERE, "fixtures", "thesis_standard")

proc = subprocess.Popen([sys.executable, os.path.join(SCRIPTS, "serve_ui.py"), str(PORT)],
                        stdout=subprocess.DEVNULL, stderr=subprocess.STDOUT)
try:
    for _ in range(40):
        time.sleep(0.25)
        try:
            urllib.request.urlopen(BASE + "/api/rules", timeout=2).read()
            break
        except Exception:
            pass
    else:
        raise SystemExit("服务器未启动")

    def post(path, data=None, raw=None, headers=None):
        if raw is not None:
            req = urllib.request.Request(BASE + path, data=raw, method="POST")
        else:
            req = urllib.request.Request(
                BASE + path, data=json.dumps(data or {}).encode(),
                headers={"Content-Type": "application/json"}, method="POST")
        if headers:
            for k, v in headers.items():
                req.add_header(k, v)
        return urllib.request.urlopen(req, timeout=60)

    # 1) 规则目录
    rules = json.load(urllib.request.urlopen(BASE + "/api/rules"))
    print(f"rules: {len(rules['rules'])} 条，默认顺序 {len(rules['default_order'])} 项")

    # 2) 上传模板与目标
    up = os.path.join(SCRIPTS, "..", "..", "tests", "fixtures", "thesis_standard")
    tpl_raw = open(os.path.join(FIX, "template.docx"), "rb").read()
    tgt_raw = open(os.path.join(FIX, "target.docx"), "rb").read()
    r1 = json.load(post("/api/upload?name=smoke_tpl.docx", raw=tpl_raw))
    r2 = json.load(post("/api/upload?name=smoke_tgt.docx", raw=tgt_raw))
    print("upload:", r1["path"], r2["path"])

    # 3) 流式检测（读 NDJSON）
    resp = post("/api/detect", {"path": r1["path"]})
    steps, done, others = [], None, []
    for line in resp.read().decode("utf-8").splitlines():
        if not line.strip():
            continue
        ev = json.loads(line)
        if ev["type"] == "step":
            steps.append(ev["label"])
        elif ev["type"] == "done":
            done = ev
        else:
            others.append(ev)
    if others:
        print("non-step events:", others)
    print("detect steps:", " → ".join(steps))
    assert done, "无 done 事件"
    mtcm = done["profile"]["page"]["margins"]["top_cm"]
    print("margins:", done["profile"]["page"]["margins"])
    assert abs(mtcm - 2.54) < 0.06, f"检测值不对: {mtcm}"
    print("detect done:", done["seconds"], "s, stats:", done["profile"]["meta"]["doc_stats"])

    # 4) 自定义顺序 + 禁用 + 覆盖值的逐条执行
    order = rules["default_order"]
    body = {"target": r2["path"], "profile": done["profile"],
            "order": order, "disabled": ["toc"],
            "overrides": {"page.margins.top_cm": 2.6}}
    resp = post("/api/apply", body)
    events = [json.loads(x) for x in resp.read().decode("utf-8").splitlines()]
    kinds = {}
    for ev in events:
        kinds[ev["type"]] = kinds.get(ev["type"], 0) + 1
    print("apply events:", kinds)
    saved = next(e for e in events if e["type"] == "saved")
    verify = next(e for e in events if e["type"] == "verify")
    print("saved:", saved["changes"], "处 →", saved["output"])
    print("verify: pass", verify["pass"], "fail", verify["fail"])
    assert os.path.isfile(saved["output"]), "输出文件不存在"
    errs = [e for e in events if e["type"] == "error" or
            (e["type"] == "rule" and e.get("status") == "error")]
    assert not errs, f"执行出错：{errs}"
    # 覆盖值生效检查
    from docx import Document  # noqa: E402
    from docx.shared import Cm  # noqa: E402
    d = Document(saved["output"])
    assert abs(d.sections[0].top_margin.cm - 2.6) < 0.06, \
        f"覆盖值未生效：{d.sections[0].top_margin.cm}"
    print("override top_cm=2.6 生效 ✓  toc 已禁用 ✓  全链路 OK")
finally:
    proc.terminate()
