# -*- coding: utf-8 -*-
"""verify_format.py — 验证输出文档是否符合确认后的格式档案 + 内容零损坏断言

用法：
  python verify_format.py <output.docx> <profile.json> [--original target.docx] [--report verify.md]
退出码：0=全部通过，1=有 FAIL。
"""
from __future__ import annotations

import argparse
import json
import os
import re
import sys

from ooxml_utils import CAPTION_RE, para_text, qn, w
from extract_format import Extractor

NUM_RE = re.compile(
    r"^(?:第[一二三四五六七八九十百\d]+\s*章|[一二三四五六七八九十]+、|（[一二三四五六七八九十]+）"
    r"|\d+(?:[-–.．]\d+)*)[\s、．.]?\s*")

TOL = 0.11  # cm/pt


def val(x):
    if isinstance(x, dict) and "value" in x:
        return x["value"]
    return x


def resolve(obj, path: str):
    cur = obj
    for part in path.split("."):
        if isinstance(cur, dict):
            cur = cur.get(part)
        else:
            return None
        if cur is None:
            return None
    return cur


# (路径, 容差 or None=精确)
CHECKS = [
    ("page.size.width_cm", TOL), ("page.size.height_cm", TOL),
    ("page.size.orientation", None),
    ("page.margins.top_cm", TOL), ("page.margins.bottom_cm", TOL),
    ("page.margins.left_cm", TOL), ("page.margins.right_cm", TOL),
    ("body.font_eastasia", None), ("body.size_pt", 0.6),
    ("body.first_line_indent_chars", 0.11), ("body.align", None),
    ("body.line_spacing.rule", None), ("body.line_spacing.value", 0.06),
    ("figures.align", None),
    ("captions.figure.position", None), ("captions.figure.separator", None),
    ("captions.figure.numbering.chapter_linked", None),
    ("captions.figure.size_pt", 0.6), ("captions.figure.bold", None),
    ("captions.table.position", None), ("captions.table.separator", None),
    ("captions.table.numbering.chapter_linked", None),
    ("captions.table.size_pt", 0.6), ("captions.table.bold", None),
    ("tables.border_rule", None),
    ("tables.three_line.top_bottom_pt", 0.11),
    ("tables.three_line.header_underline_pt", 0.11),
    ("tables.repeat_header", None), ("tables.cant_split_rows", None),
    ("tables.cell_size_pt", 0.6),
    ("code_blocks.box_style", None), ("code_blocks.font_ascii", None),
    ("bibliography.numbering_format", None),
    ("bibliography.hanging_indent_chars", 0.11),
    ("bibliography.size_pt", 0.6),
    ("formulas.omml_font", None), ("formulas.omml_size_pt", 0.6),
]
for _lvl in (1, 2, 3):
    CHECKS += [
        (f"headings.level{_lvl}.font_eastasia", None),
        (f"headings.level{_lvl}.size_pt", 0.6),
        (f"headings.level{_lvl}.bold", None),
        (f"headings.level{_lvl}.align", None),
        (f"headings.level{_lvl}.numbering.pattern", None),
    ]


def fmt_check(profile: dict, out: dict):
    results = []  # (status, path, detail)
    for path, tol in CHECKS:
        expect_raw = resolve(profile, path)
        expect = val(expect_raw)
        actual = val(resolve(out, path))
        if expect is None:
            results.append(("SKIP", path, "模板无此规则"))
            continue
        if actual is None:
            results.append(("WARN", path, f"输出未检测到（期望 {expect!r}）"))
            continue
        if tol and isinstance(expect, (int, float)) and isinstance(actual, (int, float)):
            ok = abs(actual - expect) <= tol
        else:
            ok = actual == expect
        if ok:
            results.append(("PASS", path, f"{actual!r}"))
        else:
            results.append(("FAIL", path, f"期望 {expect!r}，实际 {actual!r}"))
    # 图片宽度单独比（profile 里包了一层）
    wr = val(resolve(profile, "figures.width_rule.value")) if resolve(profile, "figures.width_rule") else None
    if wr:
        aw = val(resolve(out, "figures.width_rule.value"))
        if aw is None:
            results.append(("WARN", "figures.width_rule.value", "输出未检测到"))
        elif abs(aw - wr) > TOL:
            results.append(("FAIL", "figures.width_rule.value", f"期望 {wr}，实际 {aw}"))
        else:
            results.append(("PASS", "figures.width_rule.value", f"{aw}"))
    return results


def content_integrity(original_path: str, output_path: str):
    """除题注/标题编号外，正文文字逐字相等；题注/标题的标题部分保留。"""
    a, b = Extractor(original_path), Extractor(output_path)

    def snapshot(ex):
        """全文档段落文字（含表格内；代码装框会移动位置，故按全集比对）。
        归一化合法变化：交叉引用编号（图3-1→图1-1）、文献序号前缀、题注/标题编号。"""
        xref_norm = re.compile(r"(图|表|式)\s*\d+(?:[-–.．]\d+)?")
        bib_norm = re.compile(r"^\s*(?:\[\d+\]|\(\d+\)|\d+[\.、])\s*")
        cap_titles, texts = [], []
        for m in ex.paras:
            if not m:
                continue
            if m.get("role") == "caption":
                mm = CAPTION_RE.match(m.get("text", ""))
                cap_titles.append((str(m.get("caption_kind")),
                                   mm.group(4).strip() if mm else m.get("text", "")))
            elif m.get("role") == "special":
                continue  # 目录等占位文字允许变化
            else:
                # 标题与正文同桶（用户确认标题会改变段落角色），编号已剥离
                t = NUM_RE.sub("", m.get("text", "") or "", count=1).strip()
                texts.append(xref_norm.sub(r"\1#N", t))
        for bl in ex.flow:
            if bl["kind"] == "tbl":
                for tc in bl["el"].findall(".//" + qn("w:tc")):
                    for p0 in tc.findall(".//" + qn("w:p")):
                        t = para_text(p0).strip()
                        if t:
                            texts.append(xref_norm.sub(r"\1#N", t))
        texts = [bib_norm.sub("", t) for t in texts]
        return {"texts": sorted(texts), "caps": sorted(cap_titles)}

    sa, sb = snapshot(a), snapshot(b)
    diffs = []
    for key in ("texts", "caps"):
        if sa[key] != sb[key]:
            only_a = [x for x in sa[key] if x not in sb[key]][:3]
            only_b = [x for x in sb[key] if x not in sa[key]][:3]
            diffs.append((key, only_a, only_b))
    return diffs


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("output")
    ap.add_argument("profile")
    ap.add_argument("--original")
    ap.add_argument("--report")
    args = ap.parse_args()
    rep_path = args.report or args.output + ".verify.md"
    prof = json.load(open(args.profile, encoding="utf-8"))
    ex = Extractor(args.output)
    out = ex.profile()
    results = fmt_check(prof, out)
    if args.original:
        diffs = content_integrity(args.original, args.output)
        for key, oa, ob in diffs:
            results.append(("FAIL", f"内容完整性-{key}",
                            f"原独有 {oa} / 新独有 {ob}"))
        if not diffs:
            results.append(("PASS", "内容完整性", "正文/题注标题/标题文字/单元格 全部一致"))

    n_pass = sum(1 for s, *_ in results if s == "PASS")
    n_fail = sum(1 for s, *_ in results if s == "FAIL")
    n_warn = sum(1 for s, *_ in results if s == "WARN")
    n_skip = sum(1 for s, *_ in results if s == "SKIP")
    lines = [f"# 验证报告：{os.path.basename(args.output)}", ""]
    icon = {"PASS": "✅", "FAIL": "❌", "WARN": "⚠️", "SKIP": "⏭️"}
    for s, path, detail in results:
        if s == "PASS":
            continue
        lines.append(f"- {icon[s]} {path}: {detail}")
    lines.insert(2, f"**通过 {n_pass} · 警告 {n_warn} · 失败 {n_fail} · 跳过 {n_skip}**")
    lines.append("")
    lines.append("（通过项已省略）")
    with open(rep_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    print(f"verify: PASS {n_pass} / WARN {n_warn} / FAIL {n_fail} / SKIP {n_skip} -> {rep_path}")
    for s, path, detail in results:
        if s == "FAIL":
            print(f"  FAIL {path}: {detail}")
    sys.exit(1 if n_fail else 0)


if __name__ == "__main__":
    main()
