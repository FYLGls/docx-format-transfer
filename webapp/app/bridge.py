# -*- coding: utf-8 -*-
"""bridge.py — Pyodide 浏览器端桥接：字节进/JSON 出，复用完整核心流水线

由 webapp/index.html 在 Pyodide 就绪后加载（模块写入 /app 后 import）。
"""
from __future__ import annotations

import base64
import json
import os
import re

APP_DIR = "/app"
WORK = "/work"
os.makedirs(WORK, exist_ok=True)
sys_path = os.path.dirname(os.path.abspath(__file__))
if sys_path not in __import__("sys").path:
    __import__("sys").path.insert(0, sys_path)

from apply_format import (Applier, RULES, DEFAULT_ORDER, apply_overrides)  # noqa: E402
from extract_format import Extractor  # noqa: E402
from verify_format import content_integrity, fmt_check  # noqa: E402


def _unwrap_b64(b64: str, path: str) -> str:
    with open(path, "wb") as f:
        f.write(base64.b64decode(b64))
    return path


def _rule_of_path(path: str):
    best, best_len = None, -1
    for r in RULES:
        for p in r["paths"]:
            if path == p or path.startswith(p + "."):
                if len(p) > best_len:
                    best, best_len = r["id"], len(p)
    return best


def _val(x):
    return x.get("value") if isinstance(x, dict) and "value" in x and "rule" not in x else x


def detect(tpl_b64: str, req_text: str = "", req_b64: str = "", req_name: str = "") -> str:
    """检测模板 → JSON 字符串 {steps, profile, requirement, map_summary, error}"""
    tpl_path = _unwrap_b64(tpl_b64, os.path.join(WORK, "template.docx"))
    ex = Extractor(tpl_path)
    steps = []
    prof = {"schema_version": "1.0",
            "meta": {"source_file": "template.docx", "extracted_at": "",
                     "generator": "webapp", "doc_stats": {}},
            "flags": list(ex.flags), "dimensions_absent": []}
    for step in ex.profile_steps():
        steps.append({"key": step["key"], "label": step["label"]})
        prof[step["key"]] = step["data"]
    ex._instruction_conflict_flags(prof)
    ex._doc_stats(prof)
    ex._absent(prof)
    req_result = None
    text = req_text or ""
    if req_b64:
        ext = os.path.splitext(req_name or "")[1].lower()
        p = os.path.join(WORK, "req" + (ext or ".txt"))
        with open(p, "wb") as f:
            f.write(base64.b64decode(req_b64))
        from parse_requirements import extract_text
        try:
            text = (text + "\n" + extract_text(p)).strip()
        except Exception as e:
            req_result = {"applied": 0, "conflicts": [], "unparsed": [],
                          "error": f"要求文件解析失败：{e}"}
    if text:
        from parse_requirements import RequirementParser, merge_profile
        levels = (prof.get("headings") or {}).get("levels_in_use")
        parsed = RequirementParser(levels).parse(text)
        req_result = merge_profile(prof, parsed)
        prof = req_result["profile"]
        req_result = {"applied": req_result["applied"],
                      "conflicts": req_result["conflicts"],
                      "unparsed": req_result["unparsed"]}
    emap = ex.element_map()
    return json.dumps({
        "steps": steps, "profile": prof, "requirement": req_result,
        "map_summary": {"headings": len(emap.get("headings", [])),
                        "captions": len(emap.get("captions", [])),
                        "figures": len(emap.get("figures", [])),
                        "tables": len(emap.get("tables", [])),
                        "ambiguities": emap.get("ambiguities", [])},
        "flags": prof.get("flags", []),
    }, ensure_ascii=False)


def apply(tgt_b64: str, profile_json: str, order=None, disabled=None,
          overrides=None) -> str:
    """应用 → JSON {rules, changes, notes, verify, output_b64}"""
    tgt_path = _unwrap_b64(tgt_b64, os.path.join(WORK, "target.docx"))
    profile = json.loads(profile_json)
    if isinstance(order, str):
        order = json.loads(order)
    if isinstance(disabled, str):
        disabled = json.loads(disabled)
    if isinstance(overrides, str):
        overrides = json.loads(overrides)
    profile = apply_overrides(profile, overrides or {})
    applier = Applier(tgt_path, profile)
    rule_results = []
    for res in applier.run_rules(order, disabled):
        rule_results.append({"id": res["id"], "name": res.get("name", res["id"]),
                             "status": res["status"], "changes": res["changes"]})
    out_path = os.path.join(WORK, "output.docx")
    applier.save(out_path)
    total = sum(v for k, v in applier.changes.items() if not k.startswith("_"))
    # 验证
    out_ex = Extractor(out_path)
    checks = fmt_check(profile, out_ex.profile())
    diffs = content_integrity(tgt_path, out_path)
    for key, oa, ob in diffs:
        checks.append(("FAIL", f"内容完整性-{key}", f"原独有 {oa} / 新独有 {ob}"))
    if not diffs:
        checks.append(("PASS", "内容完整性", "正文文字逐字一致"))
    by_rule = {}
    for status, cpath, detail in checks:
        rid = _rule_of_path(cpath) or "_other"
        slot = by_rule.setdefault(rid, {"pass": 0, "warn": 0, "fail": 0, "details": []})
        key = {"PASS": "pass", "WARN": "warn", "FAIL": "fail", "SKIP": "pass"}[status]
        slot[key] += 1
        if status in ("FAIL", "WARN"):
            slot["details"].append(f"{cpath}: {detail}")
    with open(out_path, "rb") as f:
        out_b64 = base64.b64encode(f.read()).decode("ascii")
    return json.dumps({
        "rules": rule_results,
        "changes": total,
        "notes": applier.notes,
        "verify": {"by_rule": by_rule,
                   "pass": sum(1 for s, *_ in checks if s == "PASS"),
                   "warn": sum(1 for s, *_ in checks if s == "WARN"),
                   "fail": sum(1 for s, *_ in checks if s == "FAIL"),
                   "checks": len(checks)},
        "output_b64": out_b64,
    }, ensure_ascii=False)


def meta() -> str:
    return json.dumps({"rules": [r["id"] for r in RULES],
                       "default_order": DEFAULT_ORDER}, ensure_ascii=False)
