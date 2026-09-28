# -*- coding: utf-8 -*-
"""test_extra10.py — 增量 10 轮：文档场景 × 格式要求变体（全新组合）

每轮走完整网页链路：上传 → 流式检测(带要求) → 逐条执行 → verify →
要求值在输出 docx 中实测生效抽查。
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
from docx.oxml.ns import qn  # noqa: E402
HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
SCRIPTS = os.path.join(ROOT, "skill", "scripts")
FIX = os.path.join(HERE, "fixtures")
AI = os.path.join(HERE, "ai_reports")
PORT = 8960 + random.randint(0, 39)
BASE = f"http://127.0.0.1:{PORT}"

# 10 份要求变体：每份断言一个"实测生效点"
REQ_VARIANTS = [
    ("边距收紧", "页边距上下左右均为2厘米。", ("margins", 2.0)),
    ("题注翻转", "图注在图上方。表注在表下方。", ("fig_pos", "above")),
    ("页码右置", "页码居右，阿拉伯数字。", ("pagenum", "footer-right")),
    ("行距加密", "正文采用1.15倍行距。", ("linespacing", 1.15)),
    ("缩进归零", "正文首行缩进4字符。", ("indent", 4.0)),
    ("目录两级", "目录为二级目录。", (None, None)),
    ("代码装框", "代码用等宽字体装框。", (None, None)),
    ("文献编码", "参考文献按GB/T 7714顺序编码。", (None, None)),
    ("标题居中", "一级标题黑体小二居中，每章另起一页。", ("h1_size", 18.0)),
    ("边距不对称", "页边距上3厘米，下2.5厘米。", ("margins", 3.0)),
]

DOCS = [
    ("fixture:mathmodel_manual", os.path.join(FIX, "mathmodel_manual")),
    ("fixture:journal_colon", os.path.join(FIX, "journal_colon")),
    ("fixture:thesis_standard", os.path.join(FIX, "thesis_standard")),
    ("fixture:chinese_numbering", os.path.join(FIX, "chinese_numbering")),
    ("fixture:many_figures", os.path.join(FIX, "many_figures")),
    ("ai:wps_ai", os.path.join(AI, "round1")),
    ("ai:word_web", os.path.join(AI, "round1")),
    ("ai:chatgpt", os.path.join(AI, "round1")),
    ("ai:tongyi", os.path.join(AI, "round1")),
    ("ai:pandoc", os.path.join(AI, "round1")),
]


def doc_paths(doc):
    if doc[0].startswith("fixture:"):
        return (os.path.join(doc[1], "template.docx"), os.path.join(doc[1], "target.docx"))
    name = doc[0].split(":", 1)[1]
    return (os.path.join(AI, "templateA.docx") if os.path.isfile(os.path.join(AI, "templateA.docx"))
            else r"G:\数据挖掘\实验一\实验一\实验1_报告模板.docx",
            os.path.join(doc[1], f"{name}.docx"))


def ndjson(resp):
    for line in resp.read().decode("utf-8").splitlines():
        if line.strip():
            yield json.loads(line)


def main():
    proc = subprocess.Popen([sys.executable, os.path.join(SCRIPTS, "serve_ui.py"), str(PORT)],
                            stdout=subprocess.DEVNULL, stderr=subprocess.STDOUT)
    failures = []
    ok = 0
    try:
        for _ in range(60):
            time.sleep(0.25)
            try:
                urllib.request.urlopen(BASE + "/api/version", timeout=2).read()
                break
            except Exception:
                pass
        else:
            raise SystemExit("服务器未启动")

        def post(path, data=None, raw=None):
            if raw is not None:
                req = urllib.request.Request(BASE + path, data=raw, method="POST")
            else:
                req = urllib.request.Request(
                    BASE + path, data=json.dumps(data or {}).encode(),
                    headers={"Content-Type": "application/json"}, method="POST")
            return urllib.request.urlopen(req, timeout=120)

        for i, doc in enumerate(DOCS):
            label, req_text = REQ_VARIANTS[i][0], REQ_VARIANTS[i][1]
            kind, expect_val = REQ_VARIANTS[i][2]
            tpl, tgt = doc_paths(doc)
            errs = []
            r1 = json.load(post("/api/upload?name=x_tpl.docx", raw=open(tpl, "rb").read()))
            r2 = json.load(post("/api/upload?name=x_tgt.docx", raw=open(tgt, "rb").read()))
            profile = None
            applied = None
            for ev in ndjson(post("/api/detect", {
                    "path": r1["path"], "requirement_text": req_text})):
                if ev["type"] == "requirement":
                    applied = ev.get("applied")
                    if ev.get("error"):
                        errs.append("req:" + ev["error"][:100])
                elif ev["type"] == "done":
                    profile = ev["profile"]
                elif ev["type"] == "error":
                    errs.append("detect:" + ev["message"][:100])
            if applied is None or applied < 1:
                errs.append(f"要求规则 {applied}")
            saved = verify = None
            for ev in ndjson(post("/api/apply", {
                    "target": r2["path"], "profile": profile})):
                if ev["type"] == "saved":
                    saved = ev
                elif ev["type"] == "verify":
                    verify = ev
                elif ev["type"] == "error":
                    errs.append("apply:" + ev["message"][:100])
            if verify is None or verify["fail"] != 0:
                errs.append(f"verify {verify and verify['fail']}")
            # 要求值实测生效抽查
            from docx import Document
            d = Document(saved["output"])
            sec = d.sections[0]
            if kind == "margins":
                top = sec.top_margin.cm
                if abs(top - expect_val) > 0.06:
                    errs.append(f"边距未生效 {top}")
            elif kind == "linespacing":
                from docx.shared import Pt
                ls = d.paragraphs[5].paragraph_format.line_spacing
                if not (isinstance(ls, float) and abs(ls - expect_val) < 0.06):
                    errs.append(f"行距未生效 {ls}")
            elif kind == "indent":
                ind_hits = [p._p.find(qn("w:pPr") + "/" + qn("w:ind"))
                            for p in d.paragraphs]
                chars = [el.get(qn("w:firstLineChars")) for el in ind_hits
                         if el is not None and el.get(qn("w:firstLineChars"))]
                if "400" not in chars:
                    errs.append(f"缩进4字符未生效 {set(chars)}")
            elif kind == "h1_size":
                sz = None
                for p in d.paragraphs:
                    ppr = p._p.find(qn("w:pPr"))
                    if ppr is not None and ppr.find(qn("w:pStyle")) is not None and \
                            "Heading1" in ppr.find(qn("w:pStyle")).get(qn("w:val"), ""):
                        r0 = p._p.find(qn("w:r"))
                        if r0 is not None:
                            szel = r0.find(qn("w:rPr") + "/" + qn("w:sz"))
                            sz = int(szel.get(qn("w:val"))) / 2 if szel is not None else None
                        break
                if sz != expect_val:
                    errs.append(f"一级标题字号未生效 {sz}")
            elif kind == "pagenum":
                from docx.enum.text import WD_ALIGN_PARAGRAPH
                okpos = False
                for s0 in d.sections:
                    try:
                        if s0.footer.is_linked_to_previous:
                            continue
                        for para in s0.footer.paragraphs:
                            if para.alignment == WD_ALIGN_PARAGRAPH.RIGHT:
                                okpos = True
                    except Exception:
                        continue
                if not okpos:
                    errs.append("页码右置未生效")
            elif kind == "fig_pos":
                found_above = False
                paras = d.paragraphs
                for j, p in enumerate(paras[:-1]):
                    if p.text.strip().startswith("图"):
                        nxt_has_img = paras[j + 1]._p.find(".//" + qn("w:drawing")) is not None
                        if nxt_has_img:
                            found_above = True
                if not found_above:
                    errs.append("图注在上未生效")
            status = "PASS" if not errs else "FAIL"
            if errs:
                failures.append((label, errs))
            else:
                ok += 1
            print(f"[{status}] 轮{i + 1:>2} {doc[0]:<26} {label:<8} "
                  f"要求{applied:>2}条 验证❌{verify and verify['fail']}")
            for e in errs:
                print(f"     - {e}")
        print(f"\n增量测试：{ok}/10 轮通过")
        sys.exit(1 if failures else 0)
    finally:
        proc.terminate()


if __name__ == "__main__":
    main()
