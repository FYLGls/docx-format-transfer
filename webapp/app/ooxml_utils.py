# -*- coding: utf-8 -*-
"""ooxml_utils.py — docx-format-transfer 公共 OOXML 工具

提供：docx 解包/打包、命名空间、单位换算、段落文本/域提取、样式继承解析、
统计归纳（众数+置信度）、边框读取。被 extract/apply/verify 三个脚本共用。
"""
from __future__ import annotations

import copy
import io
import re
import zipfile
from collections import Counter
from datetime import datetime

from lxml import etree

# ---------------------------------------------------------------- 命名空间
NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "v": "urn:schemas-microsoft-com:vml",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
}
W = NS["w"]


def qn(tag: str) -> str:
    """'w:p' -> '{...main}p'"""
    prefix, local = tag.split(":")
    return "{%s}%s" % (NS[prefix], local)


def w(tag: str) -> str:
    return qn("w:" + tag)


# ---------------------------------------------------------------- 单位换算
def twips2cm(v: float) -> float:
    return round(int(v) / 567.0, 2)


def cm2twips(v: float) -> int:
    return int(round(v * 567))


def emu2cm(v: float) -> float:
    return round(int(v) / 360000.0, 2)


def cm2emu(v: float) -> int:
    return int(round(v * 360000))


def hp2pt(v) -> float:  # 半磅 -> 磅
    return round(int(v) / 2.0, 1)


def pt2hp(v: float) -> int:
    return int(round(v * 2))


def eighth2pt(v) -> float:  # 边框 sz（1/8 磅）-> 磅
    return round(int(v) / 8.0, 2)


def pt2eighth(v: float) -> int:
    return int(round(v * 8))


def pct2pct(v) -> float:  # tblW pct 单位（1/50 %）
    return round(int(v) / 50.0, 1)


# ---------------------------------------------------------------- 解包/打包
INTERESTING = [
    "word/document.xml",
    "word/styles.xml",
    "word/numbering.xml",
    "word/settings.xml",
    "word/footnotes.xml",
    "word/_rels/document.xml.rels",
    "[Content_Types].xml",
]

_PARSER = etree.XMLParser(remove_blank_text=False, huge_tree=True)


def load_parts(path: str) -> dict:
    """返回 {'name': etree 或 None, ...}；缺失部件值为 None；'names' 为全部条目名。"""
    out = {}
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        out["names"] = names
        wanted = list(INTERESTING) + [n for n in names
                                      if re.match(r"word/(header|footer)\d*\.xml$", n)]
        for name in wanted:
            if name in out:
                continue
            if name not in names:
                out[name] = None
                continue
            data = z.read(name)
            out[name] = etree.fromstring(data, _PARSER)
    return out


def serialize(tree) -> bytes:
    return etree.tostring(tree, xml_declaration=True, encoding="UTF-8", standalone=True)


def save_docx(src_path: str, out_path: str, modified: dict):
    """把 modified {'word/document.xml': bytes,...} 写回 src 的副本 -> out_path。"""
    with zipfile.ZipFile(src_path) as zin, zipfile.ZipFile(
        out_path, "w", zipfile.ZIP_DEFLATED
    ) as zout:
        for item in zin.infolist():
            if item.filename in modified:
                zout.writestr(item, modified[item.filename])
            else:
                zout.writestr(item, zin.read(item.filename))
        for name, data in modified.items():
            if name not in zin.namelist():
                zout.writestr(name, data)


# ---------------------------------------------------------------- 文本与域
def para_text(p) -> str:
    """拼接 w:t（跳过域指令 w:instrText）。"""
    return "".join(t.text or "" for t in p.iter(w("t")))


def para_fields(p) -> str:
    """段落内全部域指令串（fldSimple@instr + instrText），空格连接。"""
    parts = []
    for el in p.iter():
        if el.tag == w("fldSimple"):
            parts.append(el.get(w("instr")) or "")
        elif el.tag == w("instrText"):
            parts.append(el.text or "")
    return " ".join(parts)


def para_has_page_field(p) -> bool:
    return bool(re.search(r"\bPAGE\b", para_fields(p)))


# ---------------------------------------------------------------- 样式解析
def _rpr_of(el):
    return el.find(w("rPr")) if el is not None else None


def _ppr_of(el):
    return el.find(w("pPr")) if el is not None else None


def parse_rpr(rpr) -> dict:
    """rPr 元素 -> 语义 dict（None 安全）。"""
    d = {}
    if rpr is None:
        return d
    fonts = rpr.find(w("rFonts"))
    if fonts is not None:
        for slot, key in (("ascii", "font_ascii"), ("hAnsi", "font_ascii"),
                          ("eastAsia", "font_eastasia")):
            v = fonts.get(w(slot))
            if v and key not in d:
                d[key] = v
    sz = rpr.find(w("sz"))
    if sz is not None and sz.get(w("val")):
        d["size_pt"] = hp2pt(sz.get(w("val")))
    if rpr.find(w("b")) is not None:
        d["bold"] = rpr.find(w("b")).get(w("val"), "1") not in ("0", "false", "none")
    if rpr.find(w("i")) is not None:
        d["italic"] = rpr.find(w("i")).get(w("val"), "1") not in ("0", "false", "none")
    return d


def parse_ppr(ppr) -> dict:
    """pPr 元素 -> 语义 dict。"""
    d = {}
    if ppr is None:
        return d
    jc = ppr.find(w("jc"))
    if jc is not None:
        d["align"] = jc.get(w("val"))
    sp = ppr.find(w("spacing"))
    if sp is not None:
        if sp.get(w("before")):
            d["space_before_pt"] = round(int(sp.get(w("before"))) / 20.0, 1)
        if sp.get(w("after")):
            d["space_after_pt"] = round(int(sp.get(w("after"))) / 20.0, 1)
        line, rule = sp.get(w("line")), sp.get(w("lineRule"))
        if line:
            if rule == "auto":
                d["line_spacing"] = {"rule": "auto", "value": round(int(line) / 240.0, 2)}
            else:
                d["line_spacing"] = {"rule": rule, "value": round(int(line) / 20.0, 1)}
    ind = ppr.find(w("ind"))
    if ind is not None:
        if ind.get(w("firstLineChars")):
            d["first_line_chars"] = round(int(ind.get(w("firstLineChars"))) / 100.0, 1)
        elif ind.get(w("firstLine")):
            d["first_line_twips"] = int(ind.get(w("firstLine")))
        if ind.get(w("hangingChars")):
            d["hanging_chars"] = round(int(ind.get(w("hangingChars"))) / 100.0, 1)
        elif ind.get(w("hanging")):
            d["hanging_twips"] = int(ind.get(w("hanging")))
    for tag, key in (("keepNext", "keep_next"), ("keepLines", "keep_lines"),
                     ("pageBreakBefore", "page_break_before"),
                     ("widowControl", "widow_control")):
        el = ppr.find(w(tag))
        if el is not None:
            d[key] = el.get(w("val"), "1") not in ("0", "false", "none")
    ol = ppr.find(w("outlineLvl"))
    if ol is not None and ol.get(w("val")) is not None:
        d["outline_lvl"] = int(ol.get(w("val")))
    numpr = ppr.find(w("numPr"))
    if numpr is not None:
        nid = numpr.find(w("numId"))
        ilvl = numpr.find(w("ilvl"))
        if nid is not None and nid.get(w("val")) not in (None, "0"):
            d["num_id"] = nid.get(w("val"))
            d["ilvl"] = int(ilvl.get(w("val"))) if ilvl is not None else 0
    pstyle = ppr.find(w("pStyle"))
    if pstyle is not None:
        d["style_id"] = pstyle.get(w("val"))
    return d


def style_element(styles_root, style_id: str):
    if styles_root is None or not style_id:
        return None
    for st in styles_root.findall(w("style")):
        if st.get(w("styleId")) == style_id:
            return st
    return None


def style_name(st) -> str:
    if st is None:
        return ""
    nm = st.find(w("name"))
    return nm.get(w("val")) if nm is not None else ""


_HEADING_NAME_RE = re.compile(r"^(?:heading|标题)\s*(\d)", re.I)


def heading_level_of_style(styles_root, style_id) -> int | None:
    """按样式名/样式 outlineLvl 判断标题级别（1-based）。"""
    st = style_element(styles_root, style_id)
    if st is None:
        return None
    m = _HEADING_NAME_RE.match(style_name(st))
    if m:
        return int(m.group(1))
    ppr = _ppr_of(st)
    if ppr is not None:
        ol = ppr.find(w("outlineLvl"))
        if ol is not None and ol.get(w("val")) is not None and int(ol.get(w("val"))) < 9:
            return int(ol.get(w("val"))) + 1
    return None


def resolve_style(styles_root, style_id: str, _seen=None) -> dict:
    """沿 basedOn 链合并样式属性：{'rpr':..., 'ppr':..., 'name':...}，父浅子深。"""
    out = {"rpr": {}, "ppr": {}, "name": "", "id": style_id}
    if styles_root is None or not style_id:
        return out
    _seen = _seen or set()
    if style_id in _seen:
        return out
    _seen.add(style_id)
    st = style_element(styles_root, style_id)
    if st is None:
        return out
    parent = resolve_style(styles_root, st.get(w("basedOn") or "", _seen)
                           ) if st.get(w("basedOn") or "") else None
    if parent:
        out["rpr"].update(parent["rpr"])
        out["ppr"].update(parent["ppr"])
    out["rpr"].update(parse_rpr(_rpr_of(st)))
    out["ppr"].update(parse_ppr(_ppr_of(st)))
    out["name"] = style_name(st)
    return out


def doc_defaults(styles_root) -> tuple[dict, dict]:
    rpr, ppr = {}, {}
    if styles_root is not None:
        dd = styles_root.find(w("docDefaults"))
        if dd is not None:
            rd = dd.find(w("rPrDefault"))
            rpr = parse_rpr(_rpr_of(rd) if rd is not None else None)
            pd = dd.find(w("pPrDefault"))
            ppr = parse_ppr(_ppr_of(pd) if pd is not None else None)
    return rpr, ppr


def effective_ppr(p, styles_root) -> dict:
    """段落有效段落属性：docDefaults → 样式链 → 直接 pPr（style_id 保留）。"""
    dd_rpr, dd_ppr = doc_defaults(styles_root)
    merged = dict(dd_ppr)
    direct = parse_ppr(_ppr_of(p))
    chain = resolve_style(styles_root, direct.get("style_id", ""))
    merged.update(chain["ppr"])
    merged.update(direct)
    return merged


def effective_rpr(p, styles_root) -> dict:
    """段落有效字符属性：docDefaults → 样式链 → 首个有文字的 run 的直接 rPr。
    没有直接子 run 时（如文字全在域内），回退到第一个后代 run。"""
    dd_rpr, _ = doc_defaults(styles_root)
    merged = dict(dd_rpr)
    direct = parse_ppr(_ppr_of(p))
    chain = resolve_style(styles_root, direct.get("style_id", ""))
    merged.update(chain["rpr"])
    runs = p.findall(w("r"))
    if not runs:
        runs = p.findall(".//" + w("r"))
    chosen = None
    for r in runs:
        t = "".join(x.text or "" for x in r.findall(w("t")))
        if t.strip():
            chosen = r
            break
    if chosen is None and runs:
        chosen = runs[0]
    if chosen is not None:
        merged.update(parse_rpr(_rpr_of(chosen)))
    return merged


# ---------------------------------------------------------------- 统计归纳
def mode_with_conf(values: list, tol: float = 0.0):
    """众数归纳。数值型给 tol 聚类；返回
    {'value':v,'count':c,'total':n,'ratio':r,'confidence':..} 或 None。"""
    vals = [v for v in values if v is not None]
    n = len(vals)
    if n == 0:
        return None
    if tol > 0:
        clusters = []
        for v in sorted(vals):
            for c in clusters:
                if abs(v - c["rep"]) <= tol:
                    c["items"].append(v)
                    c["rep"] = sum(c["items"]) / len(c["items"])
                    break
            else:
                clusters.append({"rep": v, "items": [v]})
        best = max(clusters, key=lambda c: len(c["items"]))
        value = round(best["rep"], 2)
        count = len(best["items"])
    else:
        cnt = Counter(vals)
        value, count = cnt.most_common(1)[0]
    ratio = count / n
    if n >= 3 and ratio >= 0.9:
        conf = "high"
    elif n >= 2 and ratio >= 0.5:
        conf = "medium"
    else:
        conf = "low"
    return {"value": value, "count": count, "total": n, "ratio": round(ratio, 2),
            "confidence": conf}


def conf_result(mode: dict | None, unit: str = ""):
    """mode_with_conf 结果 -> {value, confidence, evidence}；无数据返回 None。"""
    if not mode:
        return None
    return {"value": mode["value"], "confidence": mode["confidence"],
            "evidence": f"{mode['count']}/{mode['total']}{unit}"}


def flag_outliers(name: str, values: list, tol: float, mode: dict | None, flags: list,
                  unit: str = ""):
    if not mode:
        return
    for v in values:
        if v is None:
            continue
        if abs(v - mode["value"]) > tol:
            flags.append(f"{name}: 离群值 {v}{unit}（规则值 {mode['value']}{unit}），已按众数处理")


# ---------------------------------------------------------------- 边框
BORDER_SIDES = ("top", "left", "bottom", "right", "insideH", "insideV")


def border_sz(el, side: str):
    """读取 borders 容器中某边线宽（磅）；nil/none/缺失返回 None。"""
    if el is None:
        return None
    b = el.find(w(side))
    if b is None:
        return None
    val = b.get(w("val"))
    if val in ("nil", "none", "none2"):
        return None
    sz = b.get(w("sz"))
    return eighth2pt(sz) if sz else 0.25  # 缺 sz 视为极细线


def tbl_borders(tbl) -> dict:
    tblpr = tbl.find(w("tblPr"))
    sides = {s: border_sz(_borders_child(tblpr), s) for s in BORDER_SIDES}
    return sides


def _borders_child(parent):
    return parent.find(w("tblBorders")) if parent is not None else None


def cell_border_sz(tc, side: str):
    tcpr = tc.find(w("tcPr"))
    return border_sz(_borders_child(tcpr), side)


# ---------------------------------------------------------------- 正则库
CAPTION_RE = re.compile(
    r"^(图|表|式|公式|Figure\.?|Fig\.?|Table\.?|Eq\.?)\s*"
    r"([0-9]+(?:[-–.．][0-9]+)?)\s*"
    r"([:：\-—.．]?)\s*(.+)$"
)
CHAP_CN_RE = re.compile(r"^第([一二三四五六七八九十百\d]+)\s*章")
DOTTED_NUM_RE = re.compile(r"^(\d+(?:\.\d+){0,4})[\s、．.]*\S")
CN_LIST_RE = re.compile(r"^[一二三四五六七八九十]+、")
CN_PAREN_RE = re.compile(r"^（[一二三四五六七八九十]+）")
XREF_TEXT_RE = re.compile(r"[如见参]?(?:图|表|式)\s*\d+(?:[-–]\d+)?")
CONTINUATION_RE = re.compile(r"续(?:上)?表|续表\s*\d*")
BIB_NUMLIST_RE = re.compile(r"^\[(\d+)\]\s*")
BIB_NUMLIST2_RE = re.compile(r"^\((\d+)\)\s*")
BIB_NUMLIST3_RE = re.compile(r"^(\d+)[\.、]\s*\S")
EQNUM_RE = re.compile(r"[\(（]\s*(\d+)(?:[-–](\d+))?\s*[\)）]\s*$")
SEQ_FIELD_RE = re.compile(r"SEQ\s+(图|表|式|Figure|Table|Eq)", re.I)
STYLEREF_RE = re.compile(r"STYLEREF", re.I)
PAGE_FIELD_RE = re.compile(r"\bPAGE\b")
TOC_FIELD_RE = re.compile(r'TOC[^"]*\\o\s*"(\d)-(\d)"')

MONO_FONTS = {
    "consolas", "courier new", "courier", "cascadia mono", "cascadia code",
    "jetbrains mono", "source code pro", "menlo", "monaco", "mono",
    "仿consolas", "等线mono",
}
_CODE_CHARS = set(";{}()=<>#[]$")


def code_signal_score(text: str) -> float:
    if not text:
        return 0.0
    hits = sum(1 for ch in text if ch in _CODE_CHARS)
    return hits / len(text)


def is_mono_font(font_ascii: str | None) -> bool:
    return bool(font_ascii) and font_ascii.strip().lower() in MONO_FONTS


def now_iso() -> str:
    return datetime.now().isoformat(timespec="seconds")

