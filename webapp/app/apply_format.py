# -*- coding: utf-8 -*-
"""apply_format.py — 把确认后的格式档案应用到目标文档（逐条规则执行版）

用法：
  python apply_format.py <target.docx> <profile.json> [-o output.docx]
                         [--map map.json] [--dry-run] [--report changes.md]
                         [--rules rules.json] [--overrides overrides.json]

原件永不修改；输出写新文件；幂等可重跑。
规则目录 RULES 定义每条可单独执行的格式要求；默认顺序 = 容器→样式→内容→分页（从上到下
单遍）→域收尾；--rules 可指定自定义顺序（JSON 数组），--overrides 可覆盖档案中的值
（{"page.margins.top_cm": 2.5, ...}）。
"""
from __future__ import annotations

import argparse
import copy
import json
import os
import re

from lxml import etree

from ooxml_utils import (
    qn, w, save_docx, serialize, para_text,
    cm2emu, cm2twips, pt2hp, pt2eighth,
    CAPTION_RE,
)
from extract_format import Extractor  # 复用同一套检测器做目标元素映射

XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

PPR_ORDER = ["pStyle", "keepNext", "keepLines", "pageBreakBefore", "widowControl",
             "numPr", "pBdr", "shd", "tabs", "spacing", "ind", "jc", "outlineLvl", "rPr"]
RPR_ORDER = ["rStyle", "rFonts", "b", "bCs", "i", "iCs", "color", "sz", "szCs", "vertAlign"]

NUM_RE = re.compile(
    r"^(?:第[一二三四五六七八九十百\d]+\s*章|[一二三四五六七八九十]+、|（[一二三四五六七八九十]+）"
    r"|\d+(?:[-–.．]\d+)*)[\s、．.]?\s*")


def E(tag, **attrs):
    el = etree.Element(qn(tag))
    for k, v in attrs.items():
        el.set(qn(k), v)
    return el


def val(x):
    """{'value':..} 包装 -> 裸值"""
    return x.get("value") if isinstance(x, dict) else x


def set_path(obj, path: str, value):
    parts = path.split(".")
    cur = obj
    for p in parts[:-1]:
        if not isinstance(cur, dict):
            return
        cur = cur.setdefault(p, {})
    leaf = parts[-1]
    if isinstance(cur, dict) and isinstance(cur.get(leaf), dict) and "value" in cur[leaf]:
        cur[leaf]["value"] = value  # 保持 {value,confidence,evidence} 包装
    elif isinstance(cur, dict):
        cur[leaf] = value


def apply_overrides(profile: dict, overrides: dict) -> dict:
    prof = copy.deepcopy(profile)
    for path, v in overrides.items():
        set_path(prof, path, v)
    return prof


def _insert_ordered(parent, el, order):
    tag = etree.QName(el).localname
    try:
        pos = order.index(tag)
    except ValueError:
        parent.append(el)
        return
    for child in parent:
        ctag = etree.QName(child).localname
        try:
            if order.index(ctag) > pos:
                child.addprevious(el)
                return
        except ValueError:
            continue
    parent.append(el)


def get_or_add_ppr(p):
    ppr = p.find(qn("w:pPr"))
    if ppr is None:
        ppr = E("w:pPr")
        p.insert(0, ppr)
    return ppr


def set_rpr(r_parent, east=None, ascii_=None, size=None, bold=None):
    """run 元素的 rPr：只覆写给定项，按 schema 顺序插入。"""
    rpr = r_parent.find(qn("w:rPr"))
    if rpr is None:
        rpr = E("w:rPr")
        r_parent.insert(0, rpr)
    if east or ascii_:
        rf = rpr.find(qn("w:rFonts"))
        if rf is None:
            rf = E("w:rFonts")
            _insert_ordered(rpr, rf, RPR_ORDER)
        if ascii_:
            rf.set(qn("w:ascii"), ascii_)
            rf.set(qn("w:hAnsi"), ascii_)
        if east:
            rf.set(qn("w:eastAsia"), east)
    if size is not None:
        for tag in ("w:sz", "w:szCs"):
            sz = rpr.find(qn(tag))
            if sz is None:
                sz = E(tag)
                _insert_ordered(rpr, sz, RPR_ORDER)
            sz.set(qn("w:val"), str(pt2hp(size)))
    if bold is not None:
        for tag in ("w:b", "w:bCs"):
            b = rpr.find(qn(tag))
            if bold and b is None:
                _insert_ordered(rpr, E(tag), RPR_ORDER)
            if not bold and b is not None:
                rpr.remove(b)
    return rpr


def set_ppr(p, align=None, line_spacing=None, first_line_chars=None,
            space_before=None, space_after=None, keep_next=None,
            page_break_before=None, keep_lines=None, clear_ind=False, style=None):
    ppr = get_or_add_ppr(p)
    if clear_ind:
        ind = ppr.find(qn("w:ind"))
        if ind is not None:
            ppr.remove(ind)
    if style:
        pstyle = ppr.find(qn("w:pStyle"))
        if pstyle is None:
            pstyle = E("w:pStyle")
            _insert_ordered(ppr, pstyle, PPR_ORDER)
        pstyle.set(qn("w:val"), style)
    if align:
        jc = ppr.find(qn("w:jc"))
        if jc is None:
            jc = E("w:jc")
            _insert_ordered(ppr, jc, PPR_ORDER)
        jc.set(qn("w:val"), align)
    if line_spacing or space_before is not None or space_after is not None:
        sp = ppr.find(qn("w:spacing"))
        if sp is None:
            sp = E("w:spacing")
            _insert_ordered(ppr, sp, PPR_ORDER)
        if line_spacing:
            rule, v = line_spacing.get("rule", "auto"), line_spacing.get("value", 1.5)
            if rule == "auto":
                sp.set(qn("w:line"), str(int(v * 240)))
                sp.set(qn("w:lineRule"), "auto")
            else:
                sp.set(qn("w:line"), str(int(v * 20)))
                sp.set(qn("w:lineRule"), rule)
        if space_before is not None:
            sp.set(qn("w:before"), str(int(space_before * 20)))
        if space_after is not None:
            sp.set(qn("w:after"), str(int(space_after * 20)))
    if first_line_chars is not None:
        ind = ppr.find(qn("w:ind"))
        if ind is None:
            ind = E("w:ind")
            _insert_ordered(ppr, ind, PPR_ORDER)
        for att in ("w:firstLine", "w:hanging", "w:hangingChars"):
            if ind.get(qn(att)) is not None:
                del ind.attrib[qn(att)]
        ind.set(qn("w:firstLineChars"), str(int(first_line_chars * 100)))
    for flag, tag in ((keep_next, "w:keepNext"), (page_break_before, "w:pageBreakBefore"),
                      (keep_lines, "w:keepLines")):
        if flag is None:
            continue
        el = ppr.find(qn(tag))
        if flag and el is None:
            _insert_ordered(ppr, E(tag), PPR_ORDER)
        if not flag and el is not None:
            ppr.remove(el)
    return ppr


def set_para_font(p, east=None, ascii_=None, size=None, bold=None):
    """统一段落全部 run 的字体字号（不动加粗斜体等）。"""
    changed = 0
    for r in p.findall(qn("w:r")):
        set_rpr(r, east=east, ascii_=ascii_, size=size, bold=bold)
        changed += 1
    return changed


def new_run(text, east=None, ascii_=None, size=None, bold=None):
    r = E("w:r")
    t = E("w:t")
    t.text = text
    t.set(XML_SPACE, "preserve")
    r.append(t)
    set_rpr(r, east=east, ascii_=ascii_, size=size, bold=bold)
    return r


def _text_runs(p):
    return [r for r in p.findall(qn("w:r"))
            if any((t.text or "").strip() for t in r.findall(qn("w:t")))]


def _set_para_text(p, text):
    runs = _text_runs(p)
    if not runs:
        p.append(new_run(text))
        return
    ts = runs[0].findall(qn("w:t"))
    ts[0].text = text
    for t in ts[1:]:
        t.text = ""
    for r in runs[1:]:
        for t in r.findall(qn("w:t")):
            t.text = ""


class Applier:
    def __init__(self, target_path: str, profile: dict, map_json: dict | None = None):
        from parse_requirements import normalize_profile
        self.src = target_path
        self.prof = normalize_profile(json.loads(json.dumps(profile, ensure_ascii=False)))
        self.ex = Extractor(target_path)
        self.emap = self.ex.element_map()
        if map_json:  # 允许部分覆盖（如用户确认过的标题级别）
            self.emap.update(map_json)
        self.changes: dict[str, int] = {}
        self.notes: list[str] = []
        self._num_map: dict[tuple, str] = {}
        self._code_idx = None
        self.extra_parts: dict[str, bytes] = {}  # 新增 zip 部件（如新建页脚）
        self._ct_modified = False
        self._rels_modified = False

    def log(self, cat: str):
        self.changes[cat] = self.changes.get(cat, 0) + 1

    def code_tbl_idxs(self):
        if self._code_idx is None:
            self._code_idx = {bl[0]["idx"] for bl in self.ex._code_blocks()
                              if bl[0].get("tbl_code")}
        return self._code_idx

    # ============================================================ 容器（阶段1）
    def r_page_size(self):
        size = self.prof.get("page", {}).get("size", {})
        if not (size.get("width_cm") and size.get("height_cm")):
            return
        for sect in self.ex.sections():
            pgsz = sect.find(qn("w:pgSz"))
            if pgsz is None:
                pgsz = E("w:pgSz")
                sect.insert(0, pgsz)
            pgsz.set(qn("w:w"), str(cm2twips(size["width_cm"])))
            pgsz.set(qn("w:h"), str(cm2twips(size["height_cm"])))
            if size.get("orientation") == "landscape":
                pgsz.set(qn("w:orient"), "landscape")
            self.log("页面-纸张")

    def r_page_margins(self):
        mar = self.prof.get("page", {}).get("margins", {})
        if not mar:
            return
        for sect in self.ex.sections():
            pgmar = sect.find(qn("w:pgMar"))
            if pgmar is None:
                pgmar = E("w:pgMar")
                sect.insert(1, pgmar)
            for att, key in (("top", "top_cm"), ("bottom", "bottom_cm"),
                             ("left", "left_cm"), ("right", "right_cm"),
                             ("gutter", "gutter_cm"), ("header", "header_cm"),
                             ("footer", "footer_cm")):
                v = mar.get(key)
                if v is not None:
                    pgmar.set(qn(f"w:{att}"), str(cm2twips(v)))
            self.log("页面-边距")

    def _footer_parts(self):
        return [(name, self.ex.parts[name]) for name in list(self.ex.parts)
                if name and re.match(r"word/footer\d*\.xml$", name)
                and self.ex.parts[name] is not None]

    def r_pagenum_format(self):
        body_pn = self.prof.get("page_numbers", {}).get("body") or {}
        fmt_map = {"decimal": "arabic", "lowerRoman": "roman",
                   "upperRoman": "ROMAN", "chineseCounting": "CHINESENUM3"}
        if not body_pn.get("format"):
            return
        switch = fmt_map.get(body_pn["format"], "arabic")
        for _, part in self._footer_parts():
            for fld in part.iter(qn("w:fldSimple")):
                instr = fld.get(qn("w:instr")) or ""
                if re.search(r"\bPAGE\b", instr):
                    fld.set(qn("w:instr"),
                            re.sub(r"\\\*\s*\w+", lambda _: "\\* " + switch, instr))
                    self.log("页码-域格式")
            for el in part.iter(qn("w:instrText")):
                instr = el.text or ""
                if re.search(r"\bPAGE\b", instr):
                    el.text = re.sub(r"\\\*\s*\w+", lambda _: "\\* " + switch, instr)
                    self.log("页码-域格式")
        for sect in self.ex.sections():
            num = sect.find(qn("w:pgNumType"))
            if num is not None:
                num.set(qn("w:fmt"), body_pn["format"])
        if body_pn.get("start"):
            last = self.ex.sections()[-1]
            num = last.find(qn("w:pgNumType"))
            if num is None:
                num = E("w:pgNumType")
                last.append(num)
            num.set(qn("w:start"), str(body_pn["start"]))
        if self.prof.get("page_numbers", {}).get("front_matter"):
            self.notes.append("前置部分罗马页码：模板存在分节页码结构；"
                              "目标文档节结构不同时未自动重组分节，请检查页码起始")

    def r_pagenum_pos(self):
        body_pn = self.prof.get("page_numbers", {}).get("body") or {}
        pos = body_pn.get("position")
        if not pos:
            return
        if pos == "footer-outside":
            self.notes.append("模板为奇偶页外侧页码；需在 Word 中开启『奇偶页不同』"
                              "并分别设置左右页脚，已跳过单侧对齐改写")
            return
        align = {"footer-center": "center", "footer-right": "right",
                 "footer-left": "left"}.get(pos, "center")
        touched = 0
        for _, part in self._footer_parts():
            for p in part.iter(qn("w:p")):
                blob = "".join(t.text or "" for t in p.iter(qn("w:instrText")))
                for fld in p.iter(qn("w:fldSimple")):
                    blob += " " + (fld.get(qn("w:instr")) or "")
                if re.search(r"\bPAGE\b", blob):
                    set_ppr(p, align=align)
                    set_para_font(p, east=body_pn.get("font"),
                                  size=body_pn.get("size_pt"))
                    touched += 1
                    self.log("页码-位置字体")
        if touched == 0:
            # 目标没有页码 → 按要求新建页脚页码（居左/居中/居右）
            self._create_footer_pagenum(align, body_pn)

    def _create_footer_pagenum(self, align, body_pn):
        """目标无页码时新建页脚：footer 部件 + ContentType + 关系 + sectPr 引用。"""
        names = self.ex.parts.get("names") or []
        idx = 1
        while f"word/footer{idx}.xml" in set(names):
            idx += 1
        part_name = f"word/footer{idx}.xml"
        fmt_map = {"decimal": "arabic", "lowerRoman": "roman",
                   "upperRoman": "ROMAN", "chineseCounting": "CHINESENUM3"}
        switch = fmt_map.get(body_pn.get("format") or "decimal", "arabic")
        rpr = ""
        if body_pn.get("font") or body_pn.get("size_pt"):
            inner = ""
            if body_pn.get("font"):
                inner += (f'<w:rFonts w:ascii="{body_pn["font"]}" '
                          f'w:hAnsi="{body_pn["font"]}" w:eastAsia="{body_pn["font"]}"/>')
            if body_pn.get("size_pt"):
                half = int(body_pn["size_pt"] * 2)
                inner += f'<w:sz w:val="{half}"/><w:szCs w:val="{half}"/>'
            rpr = f"<w:rPr>{inner}</w:rPr>"
        xml = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            "<w:p><w:pPr><w:jc w:val=\""
            f"{align}"
            '"/></w:pPr>'
            f"<w:r>{rpr}<w:fldChar w:fldCharType=\"begin\"/></w:r>"
            f"<w:r>{rpr}<w:instrText xml:space=\"preserve\"> PAGE \\* {switch} \\* MERGEFORMAT </w:instrText></w:r>"
            f"<w:r>{rpr}<w:fldChar w:fldCharType=\"separate\"/></w:r>"
            f"<w:r>{rpr}<w:t>1</w:t></w:r>"
            f"<w:r>{rpr}<w:fldChar w:fldCharType=\"end\"/></w:r>"
            "</w:p></w:ftr>")
        self.extra_parts[part_name] = xml.encode("utf-8")
        # [Content_Types].xml 增 Override
        ct = self.ex.parts.get("[Content_Types].xml")
        if ct is not None:
            have = any(o.get("PartName") == f"/{part_name}" for o in ct)
            if not have:
                ov = etree.SubElement(ct, "{http://schemas.openxmlformats.org/package/2006/content-types}Override")
                ov.set("PartName", "/" + part_name)
                ov.set("ContentType",
                       "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml")
                self._ct_modified = True
        # document.xml.rels 增关系
        rels = self.ex.parts.get("word/_rels/document.xml.rels")
        if rels is not None:
            max_id = 0
            for rel in rels:
                m = re.match(r"rId(\d+)$", rel.get("Id") or "")
                if m:
                    max_id = max(max_id, int(m.group(1)))
            rid = f"rId{max_id + 1}"
            rel_el = etree.SubElement(
                rels, "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship")
            rel_el.set("Id", rid)
            rel_el.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer")
            rel_el.set("Target", os.path.basename(part_name))
            self._rels_modified = True
            # 最后一个节挂 footerReference（default）
            sects = self.ex.sections()
            if sects:
                sect = sects[-1]
                ref = E("w:footerReference")
                ref.set(qn("w:type"), "default")
                ref.set(qn("r:id"), rid)
                sect.insert(0, ref)
        self.log("页码-新建页脚")
        self.notes.append(f"目标文档原本没有页码，已按要求新建页脚页码（{align}）")

    # ============================================================ 样式（阶段2）
    def _styles_root(self):
        return self.ex.parts.get("word/styles.xml")

    def r_style_body(self):
        styles_root = self._styles_root()
        prof_body = self.prof.get("body", {})
        if styles_root is None:
            self.notes.append("目标缺 styles.xml，跳过样式层")
            return
        normal = None
        for st in styles_root.findall(qn("w:style")):
            if st.get(qn("w:type")) == "paragraph":
                nm = st.find(qn("w:name"))
                if nm is not None and nm.get(qn("w:val")) in ("Normal", "正常"):
                    normal = st
        if normal is None:
            return
        rpr = normal.find(qn("w:rPr"))
        if rpr is None:
            rpr = E("w:rPr")
            normal.append(rpr)
        set_rpr(rpr, east=val(prof_body.get("font_eastasia")),
                ascii_=val(prof_body.get("font_ascii")),
                size=val(prof_body.get("size_pt")))
        ppr = normal.find(qn("w:pPr"))
        if ppr is None:
            ppr = E("w:pPr")
            normal.append(ppr)
        self._style_ppr(ppr, prof_body.get("line_spacing"),
                        val(prof_body.get("first_line_indent_chars")),
                        val(prof_body.get("align")))
        self.log("样式-正文")

    def r_style_heading(self, lvl: int):
        styles_root = self._styles_root()
        h = (self.prof.get("headings") or {}).get(f"level{lvl}")
        if not h:
            return
        if styles_root is None:
            self.notes.append("目标缺 styles.xml，跳过样式层")
            return
        st = None
        for s in styles_root.findall(qn("w:style")):
            nm = s.find(qn("w:name"))
            if nm is not None and nm.get(qn("w:val"), "").lower() == f"heading {lvl}":
                st = s
                break
        if st is None:
            st = E("w:style")
            st.set(qn("w:type"), "paragraph")
            st.set(qn("w:styleId"), f"Heading{lvl}")
            nm = E("w:name")
            nm.set(qn("w:val"), f"heading {lvl}")
            st.append(nm)
            styles_root.append(st)
            self.log("样式-新建标题样式")
        rpr = st.find(qn("w:rPr"))
        if rpr is None:
            rpr = E("w:rPr")
            st.append(rpr)
        set_rpr(rpr, east=val(h.get("font_eastasia")), ascii_=val(h.get("font_ascii")),
                size=val(h.get("size_pt")), bold=val(h.get("bold")))
        ppr = st.find(qn("w:pPr"))
        if ppr is None:
            ppr = E("w:pPr")
            st.append(ppr)
        self._style_ppr(ppr, h.get("line_spacing"), None, val(h.get("align")),
                        before=val(h.get("space_before_pt")),
                        after=val(h.get("space_after_pt")))
        ol = ppr.find(qn("w:outlineLvl"))
        if ol is None:
            ol = E("w:outlineLvl")
            ppr.append(ol)
        ol.set(qn("w:val"), str(lvl - 1))
        self.log("样式-标题")

    def r_style_h1(self):
        self.r_style_heading(1)

    def r_style_h2(self):
        self.r_style_heading(2)

    def r_style_h3(self):
        self.r_style_heading(3)

    @staticmethod
    def _style_ppr(ppr, line_spacing, first_line_chars, align,
                   before=None, after=None):
        if align:
            jc = ppr.find(qn("w:jc"))
            if jc is None:
                jc = E("w:jc")
                ppr.append(jc)
            jc.set(qn("w:val"), align)
        sp = ppr.find(qn("w:spacing"))
        if sp is None:
            sp = E("w:spacing")
            ppr.append(sp)
        if line_spacing:
            rule, v = line_spacing.get("rule", "auto"), line_spacing.get("value", 1.5)
            if rule == "auto":
                sp.set(qn("w:line"), str(int(v * 240)))
                sp.set(qn("w:lineRule"), "auto")
            else:
                sp.set(qn("w:line"), str(int(v * 20)))
                sp.set(qn("w:lineRule"), rule)
        if before is not None:
            sp.set(qn("w:before"), str(int(before * 20)))
        if after is not None:
            sp.set(qn("w:after"), str(int(after * 20)))
        if first_line_chars:
            ind = ppr.find(qn("w:ind"))
            if ind is None:
                ind = E("w:ind")
                ppr.append(ind)
            ind.set(qn("w:firstLineChars"), str(int(first_line_chars * 100)))

    # ============================================================ 内容（阶段3）
    def _level_mapping(self):
        """目标标题级别 → (模板级别, 是否参与重编号)。
        按级别排名对应；目标比模板多的层级压到最深层但不重编号（避免编号冲突）。"""
        tlevels = sorted((self.prof.get("headings") or {}).get("levels_in_use")
                         or [1, 2, 3])
        levels = sorted({h.get("level") or h.get("text_level") or 1
                         for h in self.emap.get("headings", [])})
        return {lv: (tlevels[min(i, len(tlevels) - 1)], i < len(tlevels))
                for i, lv in enumerate(levels)}

    def _top_level(self):
        lv = (self.prof.get("headings") or {}).get("levels_in_use") or [1]
        return min(lv)

    @staticmethod
    def _render_number(pattern, c1, c2, c3):
        CN = "零一二三四五六七八九十"

        def cn(n):
            return CN[n] if 0 < n < 10 else str(n)
        s = (pattern.replace("{n}", str(c1)).replace("{n2}", str(c2))
             .replace("{n3}", str(c3)).replace("{n4}", str(c3))
             .replace("{c}", cn(c1)).replace("{c2}", cn(c2)))
        return s

    def r_headings(self):
        hprof = self.prof.get("headings", {})
        patterns = {}
        for lvl in hprof.get("levels_in_use", []):
            num = (hprof.get(f"level{lvl}") or {}).get("numbering") or {}
            patterns[lvl] = num.get("pattern")
        mapping = self._level_mapping()
        tlevels = sorted({v[0] for v in mapping.values()}) or [1]
        counters = [0, 0, 0]
        for h in self.emap.get("headings", []):
            raw_lvl = h.get("level") or h.get("text_level") or 1
            lvl, allow_renumber = mapping.get(raw_lvl, (min(raw_lvl, 3), True))
            lvl = min(lvl, 3)
            meta = self.ex.paras[h["idx"]]
            meta["level"] = lvl  # 同步给题注计数用
            meta["role"] = "heading"  # 用户确认的标题：防止后续正文格式化覆盖
            p = self.ex.flow[h["idx"]]["el"]
            title = NUM_RE.sub("", meta.get("text", "")).strip()
            pattern = patterns.get(lvl) if allow_renumber else None
            try:
                rank = tlevels.index(lvl)
            except ValueError:
                rank = len(tlevels) - 1
            counters[rank] += 1
            for k in range(rank + 1, 3):
                counters[k] = 0
            if pattern:
                prefix = self._render_number(pattern, *counters)
                new_text = f"{prefix} {title}".strip() if title else prefix
                _set_para_text(p, new_text)
            set_ppr(p, style=f"Heading{lvl}")
            hp = hprof.get(f"level{lvl}") or {}
            set_para_font(p, east=val(hp.get("font_eastasia")),
                          ascii_=val(hp.get("font_ascii")),
                          size=val(hp.get("size_pt")), bold=val(hp.get("bold")))
            set_ppr(p, align=val(hp.get("align")))
            self.log("标题-样式与编号")

    def r_body(self):
        bp = self.prof.get("body", {})
        east, ascii_ = val(bp.get("font_eastasia")), val(bp.get("font_ascii"))
        size = val(bp.get("size_pt"))
        ls = bp.get("line_spacing")
        align = val(bp.get("align")) or "both"
        for i, m in enumerate(self.ex.paras):
            if not m or m.get("role") != "body" or m.get("zone") is not None:
                continue
            p = self.ex.flow[i]["el"]
            set_ppr(p, align=align, line_spacing=ls,
                    first_line_chars=val(bp.get("first_line_indent_chars")))
            set_para_font(p, east=east, ascii_=ascii_, size=size)
            self.log("正文-格式")

    def _is_obj_block(self, j, kind):
        b = self.ex.flow[j]
        if kind == "table" and b["kind"] == "tbl":
            return b["idx"] not in self.code_tbl_idxs()
        if kind == "figure" and b["kind"] == "p":
            pm = self.ex.paras[j]
            return bool(pm and pm.get("images"))
        return False

    def _fix_caption_position(self, i, kind, want):
        p = self.ex.flow[i]["el"]
        found = None  # (side, j)
        for d in (1, 2):
            j = i - d
            if 0 <= j < len(self.ex.flow) and self._is_obj_block(j, kind):
                found = ("before", j)
                break
            k = i + d
            if 0 <= k < len(self.ex.flow) and self._is_obj_block(k, kind):
                found = ("after", k)
                break
        if not found:
            return
        side, j = found
        obj = self.ex.flow[j]["el"]
        if want == "below":
            if side == "after":      # 对象在题注后 → 题注应移到对象之后
                obj.addnext(p)
                self.log("题注-位置调整")
        else:  # above
            if side == "before":     # 对象在题注前 → 题注应移到对象之前
                obj.addprevious(p)
                self.log("题注-位置调整")

    def _rebuild_captions(self, kinds):
        caps = self.prof.get("captions", {})
        counters = {k: {} for k in kinds}
        gcount = {k: 0 for k in kinds}
        chapter = 0
        top_lvl = self._top_level()
        order = []
        for i, m in enumerate(self.ex.paras):
            if not m:
                continue
            if m.get("role") == "heading" and m.get("level") == top_lvl:
                chapter += 1
            elif m.get("caption_kind") in kinds:
                order.append((i, m, chapter))
        for i, m, ch in order:
            kind = m["caption_kind"]
            cp = caps.get(kind)
            if not cp:
                continue
            p = self.ex.flow[i]["el"]
            text = m.get("text", "")
            mm = CAPTION_RE.match(text)
            old_num = mm.group(2) if mm else None
            title = (mm.group(4).strip() if mm else text.strip())
            if cp.get("numbering", {}).get("chapter_linked"):
                counters[kind][ch] = counters[kind].get(ch, 0) + 1
                num = f"{ch}-{counters[kind][ch]}"
            else:
                gcount[kind] += 1
                num = str(gcount[kind])
            label = cp.get("label") or ("图" if kind == "figure" else "表")
            sep = cp.get("separator") or ""
            new_text = f"{label}{num}{sep}{title}"
            for child in list(p):
                if child.tag != qn("w:pPr"):
                    p.remove(child)
            p.append(new_run(new_text, east=val(cp.get("font_eastasia")),
                             ascii_=val(cp.get("font_ascii")),
                             size=val(cp.get("size_pt")),
                             bold=val(cp.get("bold"))))
            set_ppr(p, align=val(cp.get("align")) or "center")
            self._fix_caption_position(i, kind,
                                       cp.get("position") or
                                       ("below" if kind == "figure" else "above"))
            if old_num and old_num != num:
                self._num_map.setdefault((kind, old_num), num)
            self.log(f"题注-{'图' if kind == 'figure' else '表'}")
        for kind in kinds:
            cp = caps.get(kind)
            if cp and str(cp.get("numbering", {}).get("mode", "")).startswith("field") \
                    and cp.get("numbering", {}).get("chapter_linked"):
                self.notes.append(
                    f"{'图' if kind == 'figure' else '表'}题注：模板为按章域编号，"
                    "已按章重排为字面编号（显示正确且稳定；如需 Word 域自动编号可再告知）")

    def r_captions_fig(self):
        self._rebuild_captions(("figure",))

    def r_captions_tbl(self):
        self._rebuild_captions(("table",))

    def r_xrefs(self):
        if not self._num_map:
            return
        pat = re.compile(r"(图|表|式)\s*([0-9]+(?:[-–.．][0-9]+)?)")

        def repl(mm):
            label, num = mm.group(1), mm.group(2)
            lab2kind = {"图": "figure", "表": "table", "式": "equation"}
            kind = lab2kind.get(label)
            for (k2, old), new in self._num_map.items():
                if k2 == kind and old == num:
                    return f"{label}{new}"
            return mm.group(0)

        for i, m in enumerate(self.ex.paras):
            if not m or m.get("role") != "body":
                continue
            for t in self.ex.flow[i]["el"].iter(qn("w:t")):
                if t.text and pat.search(t.text):
                    new = pat.sub(repl, t.text)
                    if new != t.text:
                        t.text = new
                        self.log("交叉引用-更新")

    def r_bib(self):
        bib = self.prof.get("bibliography", {})
        if not bib.get("detected"):
            return
        fmt = bib.get("numbering_format") or "[{n}]"
        hang = val(bib.get("hanging_indent_chars"))
        size = val(bib.get("size_pt"))
        n = 0
        for i, m in enumerate(self.ex.paras):
            if not m or m.get("zone") != "bibliography" or m.get("role") != "body":
                continue
            if not m.get("text"):
                continue
            n += 1
            p = self.ex.flow[i]["el"]
            text = re.sub(r"^\s*(?:\[\d+\]|\(\d+\)|\d+[\.、])\s*", "", m["text"])
            prefix = fmt.replace("{n}", str(n))
            _set_para_text(p, f"{prefix} {text}")
            set_ppr(p, clear_ind=True)
            if hang:
                ppr = get_or_add_ppr(p)
                ind = ppr.find(qn("w:ind"))
                if ind is None:
                    ind = E("w:ind")
                    _insert_ordered(ppr, ind, PPR_ORDER)
                ind.set(qn("w:hangingChars"), str(int(hang * 100)))
                ind.set(qn("w:leftChars"), str(int(hang * 100)))
            set_para_font(p, size=size)
            self.log("参考文献-版式")

    def _tables(self):
        return [b for b in self.ex.flow
                if b["kind"] == "tbl" and b["idx"] not in self.code_tbl_idxs()]

    def r_table_borders(self):
        tp = self.prof.get("tables", {})
        if not tp:
            return
        tl = tp.get("three_line") or {}
        hdr = tp.get("header_row") or {}
        rule = tp.get("border_rule") or \
            ("three-line" if val(tl.get("enabled")) else None)
        thick = pt2eighth(tl.get("top_bottom_pt") or 1.5)
        thin = pt2eighth(tl.get("header_underline_pt") or 0.75)
        for b in self._tables():
            tbl = b["el"]
            tblpr = tbl.find(qn("w:tblPr"))
            if tblpr is None:
                tblpr = E("w:tblPr")
                tbl.insert(0, tblpr)
            if rule not in ("three-line", "grid", "outline", "none"):
                continue
            borders = tblpr.find(qn("w:tblBorders"))
            if borders is not None:
                tblpr.remove(borders)
            borders = E("w:tblBorders")
            if rule == "three-line":
                plan = (("top", "single", thick), ("bottom", "single", thick),
                        ("left", "nil", None), ("right", "nil", None),
                        ("insideH", "nil", None), ("insideV", "nil", None))
            elif rule == "grid":
                plan = tuple((s, "single", 4) for s in
                             ("top", "left", "bottom", "right", "insideH", "insideV"))
            elif rule == "outline":
                plan = (("top", "single", 4), ("bottom", "single", 4),
                        ("left", "nil", None), ("right", "nil", None),
                        ("insideH", "nil", None), ("insideV", "nil", None))
            else:  # none
                plan = tuple((s, "nil", None) for s in
                             ("top", "left", "bottom", "right", "insideH", "insideV"))
            for side, v, sz in plan:
                el = E(f"w:{side}")
                el.set(qn("w:val"), v)
                if sz:
                    el.set(qn("w:sz"), str(sz))
                el.set(qn("w:color"), "000000")
                el.set(qn("w:space"), "0")
                borders.append(el)
            tblpr.append(borders)
            if rule == "three-line":
                first = tbl.find(qn("w:tr"))
                if first is not None:
                    for tc in first.findall(qn("w:tc")):
                        tcpr = tc.find(qn("w:tcPr"))
                        if tcpr is None:
                            tcpr = E("w:tcPr")
                            tc.insert(0, tcpr)
                        cb = tcpr.find(qn("w:tcBorders"))
                        if cb is not None:
                            tcpr.remove(cb)
                        cb = E("w:tcBorders")
                        bo = E("w:bottom")
                        bo.set(qn("w:val"), "single")
                        bo.set(qn("w:sz"), str(thin))
                        bo.set(qn("w:color"), "000000")
                        bo.set(qn("w:space"), "0")
                        cb.append(bo)
                        tcpr.append(cb)
                        if not val(hdr.get("shading_fill")):
                            shd = tcpr.find(qn("w:shd"))
                            if shd is not None:
                                tcpr.remove(shd)
            self.log(f"表格-边框规则化({rule})")

    def r_table_rows(self):
        tp = self.prof.get("tables", {})
        if not tp:
            return
        rep = val(tp.get("repeat_header"))
        cs = val(tp.get("cant_split_rows"))
        for b in self._tables():
            tbl = b["el"]
            rows = tbl.findall(qn("w:tr"))
            if not rows:
                continue
            first = rows[0]
            trpr = first.find(qn("w:trPr"))
            if rep is True:
                if trpr is None:
                    trpr = E("w:trPr")
                    first.insert(0, trpr)
                if trpr.find(qn("w:tblHeader")) is None:
                    trpr.append(E("w:tblHeader"))
            elif rep is False and trpr is not None:
                th = trpr.find(qn("w:tblHeader"))
                if th is not None:
                    trpr.remove(th)
            if cs is not None:
                for tr in rows:
                    trpr = tr.find(qn("w:trPr"))
                    if cs:
                        if trpr is None:
                            trpr = E("w:trPr")
                            tr.insert(0, trpr)
                        if trpr.find(qn("w:cantSplit")) is None:
                            trpr.append(E("w:cantSplit"))
                    elif trpr is not None:
                        csp = trpr.find(qn("w:cantSplit"))
                        if csp is not None:
                            trpr.remove(csp)
            self.log("表格-跨页属性")
        if (tp.get("continuation_convention") or {}).get("detected"):
            self.notes.append("续表惯例：已用『行不拆分 + 跨页重复表头』等价实现；"
                              "渲染级整表合页检测需 LibreOffice，未安装时跳过")

    def r_table_cells(self):
        tp = self.prof.get("tables", {})
        if not tp:
            return
        hdr = tp.get("header_row") or {}
        cs = val(tp.get("cell_size_pt"))
        ce = val(tp.get("cell_font_eastasia"))
        ca = val(tp.get("cell_align"))
        wm = tp.get("width_mode") or {}
        for b in self._tables():
            tbl = b["el"]
            tblpr = tbl.find(qn("w:tblPr"))
            if tblpr is None:
                tblpr = E("w:tblPr")
                tbl.insert(0, tblpr)
            if wm.get("type") == "pct":
                tblw = tblpr.find(qn("w:tblW"))
                if tblw is None:
                    tblw = E("w:tblW")
                    tblpr.append(tblw)
                tblw.set(qn("w:type"), "pct")
                tblw.set(qn("w:w"), str(wm.get("value") or 5000))
            for ri, tr in enumerate(tbl.findall(qn("w:tr"))):
                for tc in tr.findall(qn("w:tc")):
                    for p0 in tc.findall(".//" + qn("w:p")):
                        if ri == 0:
                            set_para_font(p0, east=ce, size=cs,
                                          bold=val(hdr.get("bold")))
                            set_ppr(p0, align=val(hdr.get("align")) or ca)
                        else:
                            set_para_font(p0, east=ce, size=cs)
                            set_ppr(p0, align=ca)
            self.log("表格-单元格")

    def r_fig_size(self):
        fp = self.prof.get("figures", {})
        wr = fp.get("width_rule") or {}
        target_cm = wr.get("value") if wr.get("type") == "fixed" else None
        if not target_cm:
            return
        for i, m in enumerate(self.ex.paras):
            if not m or not (m.get("images") or m.get("anchors")):
                continue
            p = self.ex.flow[i]["el"]
            for drawing in p.findall(".//" + qn("w:drawing")):
                resized = False
                for ext in drawing.iter(qn("wp:extent")):
                    if ext.get("cx"):
                        old_cx, old_cy = int(ext.get("cx")), int(ext.get("cy"))
                        new_cx = cm2emu(target_cm)
                        ext.set("cx", str(new_cx))
                        ext.set("cy", str(int(old_cy * new_cx / max(old_cx, 1))))
                        resized = True
                for xfrm_ext in drawing.iter(qn("a:ext")):
                    if xfrm_ext.get("cx"):
                        old_cx, old_cy = int(xfrm_ext.get("cx")), int(xfrm_ext.get("cy"))
                        new_cx = cm2emu(target_cm)
                        xfrm_ext.set("cx", str(new_cx))
                        xfrm_ext.set("cy", str(int(old_cy * new_cx / max(old_cx, 1))))
                        resized = True
                if resized:
                    self.log("图片-尺寸")

    def r_fig_align(self):
        align = val(self.prof.get("figures", {}).get("align"))
        if not align:
            return
        for i, m in enumerate(self.ex.paras):
            if not m or not (m.get("images") or m.get("anchors")):
                continue
            set_ppr(self.ex.flow[i]["el"], align=align)
            self.log("图片-对齐")

    def r_formula(self):
        fm = self.prof.get("formulas")
        if not fm or "formulas" in self.prof.get("dimensions_absent", []):
            return
        font = val(fm.get("omml_font"))
        size = val(fm.get("omml_size_pt"))
        if not font and not size:
            return
        for r in self.ex.doc.findall(".//" + qn("m:r")):
            set_rpr(r, ascii_=font, east=font, size=size)
            self.log("公式-字体")

    def r_code(self):
        cb = self.prof.get("code_blocks", {})
        if not cb.get("detected") or "code_blocks" in self.prof.get("dimensions_absent", []):
            return
        box = cb.get("box_style") or "single-cell-table"
        size = val(cb.get("size_pt"))
        font = val(cb.get("font_ascii")) or "Consolas"
        fill = cb.get("shading_fill") or "F2F2F2"
        for block in self.ex._code_blocks():
            if block[0].get("tbl_code"):
                continue
            paras = [self.ex.flow[it["idx"]]["el"] for it in block]
            for p in paras:
                set_para_font(p, ascii_=font, east=font, size=size)
                set_ppr(p, line_spacing={"rule": "auto", "value": 1.0})
            if box == "single-cell-table":
                first = paras[0]
                anchor_parent = first.getparent()
                anchor_idx = anchor_parent.index(first)
                tbl = E("w:tbl")
                tblpr = E("w:tblPr")
                tblw = E("w:tblW")
                tblw.set(qn("w:type"), "pct")
                tblw.set(qn("w:w"), "5000")
                tblpr.append(tblw)
                borders = E("w:tblBorders")
                for side in ("top", "left", "bottom", "right"):
                    el = E(f"w:{side}")
                    el.set(qn("w:val"), "single")
                    el.set(qn("w:sz"), "4")
                    el.set(qn("w:color"), "000000")
                    el.set(qn("w:space"), "0")
                    borders.append(el)
                tblpr.append(borders)
                tbl.append(tblpr)
                grid = E("w:tblGrid")
                col = E("w:gridCol")
                col.set(qn("w:w"), "9638")
                grid.append(col)
                tbl.append(grid)
                tr = E("w:tr")
                trpr = E("w:trPr")
                trpr.append(E("w:cantSplit"))
                tr.append(trpr)
                tc = E("w:tc")
                tcpr = E("w:tcPr")
                tcw = E("w:tcW")
                tcw.set(qn("w:type"), "pct")
                tcw.set(qn("w:w"), "5000")
                tcpr.append(tcw)
                shd = E("w:shd")
                shd.set(qn("w:val"), "clear")
                shd.set(qn("w:fill"), fill)
                tcpr.append(shd)
                tc.append(tcpr)
                anchor_parent.insert(anchor_idx, tbl)
                for p in paras:
                    tc.append(p)  # 移动段落进单元格（表已在原位）
                tr.append(tc)
                tbl.append(tr)
            else:
                for p in paras:
                    ppr = get_or_add_ppr(p)
                    pbdr = E("w:pBdr")
                    for side in ("top", "left", "bottom", "right"):
                        el = E(f"w:{side}")
                        el.set(qn("w:val"), "single")
                        el.set(qn("w:sz"), "4")
                        el.set(qn("w:color"), "000000")
                        el.set(qn("w:space"), "4")
                        pbdr.append(el)
                    _insert_ordered(ppr, pbdr, PPR_ORDER)
                    shd = E("w:shd")
                    shd.set(qn("w:val"), "clear")
                    shd.set(qn("w:fill"), fill)
                    _insert_ordered(ppr, shd, PPR_ORDER)
            self.log("代码-装框")

    # ============================================================ 分页（阶段4）
    def r_keep_captions(self):
        caps = self.prof.get("captions", {})
        for kind in ("figure", "table"):
            cp = caps.get(kind)
            if not cp:
                continue
            want = cp.get("position") or ("below" if kind == "figure" else "above")
            for i, m in enumerate(self.ex.paras):
                if not m or m.get("caption_kind") != kind:
                    continue
                p = self.ex.flow[i]["el"]
                if (kind == "figure" and want == "below") or \
                        (kind == "table" and want == "above"):
                    set_ppr(p, keep_next=True)
                else:
                    set_ppr(p, keep_lines=True)
                self.log("分页-题注绑定")

    def r_keep_figures(self):
        caps = self.prof.get("captions", {})
        cp = caps.get("figure")
        if not cp:
            return
        if cp.get("position", "below") == "below":
            for i, m in enumerate(self.ex.paras):
                if m and m.get("images"):
                    set_ppr(self.ex.flow[i]["el"], keep_next=True)
                    self.log("分页-图与题注绑定")

    def r_keep_headings(self):
        hprof = self.prof.get("headings", {})
        top_lvl = self._top_level()
        for h in self.emap.get("headings", []):
            meta = self.ex.paras[h["idx"]]
            lvl = meta.get("level") or 1
            hp = hprof.get(f"level{lvl}") or {}
            p = self.ex.flow[h["idx"]]["el"]
            if val(hp.get("keep_next")) is not False:
                set_ppr(p, keep_next=True)
            if lvl == top_lvl and val(hp.get("page_break_before")):
                set_ppr(p, page_break_before=True)
            self.log("分页-标题绑定")

    # ============================================================ 收尾（阶段5）
    def r_toc(self):
        toc = self.prof.get("toc", {})
        if toc.get("mode") != "field":
            return
        has_toc = False
        for p_el in self.ex.doc.iter(qn("w:p")):
            blob = "".join(t.text or "" for t in p_el.iter(qn("w:instrText")))
            for fld in p_el.iter(qn("w:fldSimple")):
                blob += " " + (fld.get(qn("w:instr")) or "")
            if re.search(r"\bTOC\b", blob):
                has_toc = True
                break
        if has_toc:
            return
        for sp in self.emap.get("specials", []):
            if sp.get("role") == "toc":
                p = self.ex.flow[sp["idx"]]["el"]
                tocp = E("w:p")
                fld = E("w:fldSimple")
                fld.set(qn("w:instr"),
                        f' TOC \\o "1-{toc.get("depth") or 3}" \\h \\z \\u ')
                r = E("w:r")
                t = E("w:t")
                t.text = "（目录域：在 Word 中右键→更新域）"
                r.append(t)
                fld.append(r)
                tocp.append(fld)
                p.addnext(tocp)
                self.log("目录-插入域")
                break

    # ============================================================ 执行器
    def run_rules(self, order=None, disabled=None):
        """逐条执行规则（生成器：yield 每条结果）。
        order: 规则 id 列表（自定义顺序）；缺省用默认顺序。
        disabled: 跳过的规则 id 集合。"""
        disabled = set(disabled or [])
        ids = order or [r["id"] for r in RULES]
        for rid in ids:
            rule = RULE_BY_ID.get(rid)
            if rule is None:
                yield {"id": rid, "status": "skipped", "changes": 0,
                       "reason": "未知规则"}
                continue
            if rid in disabled:
                yield {"id": rid, "name": rule["name"], "status": "skipped",
                       "changes": 0, "reason": "用户禁用"}
                continue
            before = sum(self.changes.values())
            notes_before = len(self.notes)
            try:
                getattr(self, rule["fn"])()
                status = "done"
            except Exception as e:  # 单条失败不中断整体
                status = "error"
                self.notes.append(f"规则 {rule['name']} 执行失败：{e}")
            yield {"id": rid, "name": rule["name"], "status": status,
                   "changes": sum(self.changes.values()) - before,
                   "notes": self.notes[notes_before:]}

    def run(self):
        for _ in self.run_rules():
            pass
        return self.changes

    def save(self, out_path):
        modified = {"word/document.xml": serialize(self.ex.doc)}
        styles_root = self._styles_root()
        if styles_root is not None:
            modified["word/styles.xml"] = serialize(styles_root)
        for name in list(self.ex.parts):
            if name and re.match(r"word/(header|footer)\d*\.xml$", name) and \
                    self.ex.parts[name] is not None:
                modified[name] = serialize(self.ex.parts[name])
        if self._ct_modified:
            modified["[Content_Types].xml"] = serialize(
                self.ex.parts["[Content_Types].xml"])
        if self._rels_modified:
            modified["word/_rels/document.xml.rels"] = serialize(
                self.ex.parts["word/_rels/document.xml.rels"])
        modified.update(self.extra_parts)
        save_docx(self.src, out_path, modified)


# ---------------------------------------------------------------- 规则目录
def _R(rid, name, cat, stage, fn, paths, paginated=False, desc=""):
    return {"id": rid, "name": name, "cat": cat, "stage": stage, "fn": fn,
            "paths": paths, "paginated": paginated, "desc": desc}


RULES = [
    _R("page_size", "纸张大小与方向", "容器", 1, "r_page_size",
       ["page.size.width_cm", "page.size.height_cm", "page.size.orientation"]),
    _R("page_margins", "页边距（含页眉页脚距离）", "容器", 1, "r_page_margins",
       ["page.margins.top_cm", "page.margins.bottom_cm", "page.margins.left_cm",
        "page.margins.right_cm", "page.margins.gutter_cm"]),
    _R("pagenum_format", "页码格式与起始页", "容器", 1, "r_pagenum_format",
       ["page_numbers.body.format", "page_numbers.body.start"]),
    _R("pagenum_pos", "页码位置与字体", "容器", 1, "r_pagenum_pos",
       ["page_numbers.body.position", "page_numbers.body.font", "page_numbers.body.size_pt"]),
    _R("style_body", "正文样式定义", "样式", 2, "r_style_body",
       ["body.font_eastasia", "body.font_ascii", "body.size_pt",
        "body.line_spacing", "body.first_line_indent_chars", "body.align"]),
    _R("style_h1", "一级标题样式", "样式", 2, "r_style_h1",
       ["headings.level1"]),
    _R("style_h2", "二级标题样式", "样式", 2, "r_style_h2",
       ["headings.level2"]),
    _R("style_h3", "三级标题样式", "样式", 2, "r_style_h3",
       ["headings.level3"]),
    _R("headings", "标题重排与套样式", "内容", 3, "r_headings",
       ["headings"], desc="按模板编号体系重排全部标题并套用标题样式"),
    _R("body", "正文段落格式", "内容", 3, "r_body", ["body"]),
    _R("captions_fig", "图题注重建（位置/编号/分隔符）", "内容", 3, "r_captions_fig",
       ["captions.figure.position", "captions.figure.separator",
        "captions.figure.numbering", "captions.figure.size_pt", "captions.figure.bold"]),
    _R("captions_tbl", "表题注重建（位置/编号/分隔符）", "内容", 3, "r_captions_tbl",
       ["captions.table.position", "captions.table.separator",
        "captions.table.numbering", "captions.table.size_pt", "captions.table.bold"]),
    _R("xrefs", "交叉引用编号同步", "内容", 3, "r_xrefs", ["cross_references"],
       desc="题注编号变化后同步正文中的“如图X所示”"),
    _R("bib", "参考文献版式", "内容", 3, "r_bib", ["bibliography"]),
    _R("table_borders", "表格边框规则（三线/网格/上下框/无框）", "内容", 3,
       "r_table_borders", ["tables.border_rule", "tables.three_line"]),
    _R("table_rows", "表头行与跨页属性（重复表头/行不拆分）", "内容", 3, "r_table_rows",
       ["tables.repeat_header", "tables.cant_split_rows"]),
    _R("table_cells", "表内字体与表宽", "内容", 3, "r_table_cells",
       ["tables.cell_size_pt", "tables.cell_font_eastasia", "tables.cell_align",
        "tables.width_mode"]),
    _R("fig_size", "图片尺寸（等比缩放）", "内容", 3, "r_fig_size",
       ["figures.width_rule"]),
    _R("fig_align", "图片对齐", "内容", 3, "r_fig_align", ["figures.align"]),
    _R("formula", "公式字体（OMML）", "内容", 3, "r_formula",
       ["formulas.omml_font", "formulas.omml_size_pt"]),
    _R("code", "代码块装框与字体", "内容", 3, "r_code",
       ["code_blocks.box_style", "code_blocks.font_ascii", "code_blocks.size_pt"]),
    _R("keep_captions", "题注与图表绑定（不跨页分离）", "分页", 4, "r_keep_captions",
       ["captions.figure.position"], paginated=True,
       desc="分页类：需内容高度定型后从上到下单遍处理，默认排在最后"),
    _R("keep_figures", "图与题注绑定（keepNext）", "分页", 4, "r_keep_figures",
       ["figures.keep_with_caption"], paginated=True),
    _R("keep_headings", "标题防孤行/每章另起页", "分页", 4, "r_keep_headings",
       ["headings.level1.keep_next", "headings.level1.page_break_before"],
       paginated=True),
    _R("toc", "目录域", "收尾", 5, "r_toc", ["toc"]),
]
RULE_BY_ID = {r["id"]: r for r in RULES}
DEFAULT_ORDER = [r["id"] for r in RULES]


def render_changes(changes: dict, notes: list, out: str, dry: bool):
    lines = ["# 修改报告" + ("（dry-run，未写入文件）" if dry else ""), ""]
    total = sum(v for k, v in changes.items() if not k.startswith("_"))
    lines.append(f"共 {total} 处修改：")
    for k, v in sorted(changes.items()):
        lines.append(f"- {k}: {v}")
    if notes:
        lines.append("")
        lines.append("## 说明与跳过项")
        for n in notes:
            lines.append(f"- {n}")
    with open(out, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    return total


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("target")
    ap.add_argument("profile")
    ap.add_argument("-o", "--out")
    ap.add_argument("--map", dest="map_json")
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--report")
    ap.add_argument("--rules", dest="rules_json",
                    help="自定义规则顺序 JSON（数组或 {order,disabled}）")
    ap.add_argument("--overrides", help="值覆盖 JSON {profile.path: value}")
    args = ap.parse_args()
    stem = os.path.splitext(args.target)[0]
    out = args.out or stem + ".formatted.docx"
    rep = args.report or stem + ".changes.md"
    prof = json.load(open(args.profile, encoding="utf-8"))
    map_json = json.load(open(args.map_json, encoding="utf-8")) if args.map_json else None
    order, disabled = None, None
    if args.rules_json:
        rj = json.load(open(args.rules_json, encoding="utf-8"))
        if isinstance(rj, list):
            order = rj
        else:
            order, disabled = rj.get("order"), rj.get("disabled")
    if args.overrides:
        prof = apply_overrides(
            prof, json.load(open(args.overrides, encoding="utf-8")))
    applier = Applier(args.target, prof, map_json)
    results = list(applier.run_rules(order, disabled))
    total = render_changes(applier.changes, applier.notes, rep, args.dry_run)
    if args.dry_run:
        print(f"[dry-run] 将修改 {total} 处；报告 -> {rep}")
    else:
        applier.save(out)
        print(f"已写入 {out}；共 {total} 处修改；报告 -> {rep}")
    for r in results:
        print(f"  [{r['status']}] {r.get('name', r['id'])}: {r['changes']} 处")


if __name__ == "__main__":
    main()
