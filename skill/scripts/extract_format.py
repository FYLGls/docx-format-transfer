# -*- coding: utf-8 -*-
"""extract_format.py — 从 docx 提取格式档案（profile.json）+ 可读清单（report.md）+ 元素映射（map.json）

用法：
  python extract_format.py <input.docx> [-o profile.json] [--report report.md] [--map map.json]
不写盘时默认输出到输入文件同目录：<stem>.profile.json / <stem>.report.md / <stem>.map.json
"""
from __future__ import annotations

import argparse
import json
import os
import re
import sys

from lxml import etree

from ooxml_utils import (
    NS, W, qn, w, load_parts, para_text, para_fields,
    twips2cm, emu2cm, hp2pt, pct2pct, border_sz,
    parse_ppr, parse_rpr, resolve_style, doc_defaults, effective_ppr, effective_rpr,
    heading_level_of_style, style_element, style_name,
    mode_with_conf, conf_result, flag_outliers,
    CAPTION_RE, CHAP_CN_RE, DOTTED_NUM_RE, CN_LIST_RE, CN_PAREN_RE,
    XREF_TEXT_RE, CONTINUATION_RE, BIB_NUMLIST_RE, BIB_NUMLIST2_RE, BIB_NUMLIST3_RE,
    EQNUM_RE, SEQ_FIELD_RE, STYLEREF_RE, PAGE_FIELD_RE, TOC_FIELD_RE,
    code_signal_score, is_mono_font, now_iso,
)

SPECIAL_TITLES = [
    ("摘要", "abstract"), ("摘 要", "abstract"), ("ABSTRACT", "abstract"),
    ("关键词", "keywords"), ("目 录", "toc"), ("目录", "toc"), ("目次", "toc"),
    ("参考文献", "bibliography"), ("致 谢", "acknowledgement"), ("致谢", "acknowledgement"),
    ("附录", "appendix"), ("结 论", "conclusion"), ("结论", "conclusion"),
    ("结束语", "conclusion"),
]
_SPECIAL_EXACT = {"abstract", "keywords", "key words", "toc", "references",
                  "acknowledgement", "conclusion"}


def _is_special_title(text: str):
    t = text.strip()
    tl = t.lower()
    if tl in _SPECIAL_EXACT:
        return "abstract"
    for key, role in SPECIAL_TITLES:
        if t == key or (t.startswith(key) and len(t) <= len(key) + 6 and re.match(
                re.escape(key) + r"[：:\s0-9A-Za-z一二三四五六七八九十]*$", t)):
            return role
    return None


def _sep_from_match(text: str, m) -> str:
    between = text[m.end(2):m.start(4)]
    if m.group(3):
        return m.group(3)
    if "\u3000" in between:
        return "\u3000"
    if " " in between:
        return " "
    return ""


class Extractor:
    def __init__(self, path: str):
        self.path = path
        self.parts = load_parts(path)
        self.doc = self.parts.get("word/document.xml")
        if self.doc is None:
            raise SystemExit("不是有效的 docx（缺少 word/document.xml）")
        self.body = self.doc.find(w("body"))
        self.styles = self.parts.get("word/styles.xml")
        self.numbering = self.parts.get("word/numbering.xml")
        self.settings = self.parts.get("word/settings.xml")
        self.rels = self._load_rels()
        self.flags: list[str] = []
        self.flow = []          # 文档顺序块：{'kind':'p'|'tbl','el':..,'idx':..}
        self.paras = []         # 段落元数据（与 flow 中 p 对齐）
        self._build_flow()
        self._classify()

    # ------------------------------------------------------------- 基础
    def _load_rels(self):
        rels = {}
        root = self.parts.get("word/_rels/document.xml.rels")
        if root is not None:
            for rel in root:
                rels[rel.get("Id")] = rel.get("Target")
        return rels

    def _build_flow(self):
        for child in self.body:
            if child.tag == w("p"):
                self.flow.append({"kind": "p", "el": child, "idx": len(self.flow)})
            elif child.tag == w("tbl"):
                self.flow.append({"kind": "tbl", "el": child, "idx": len(self.flow)})

    def _classify(self):
        zone = None  # bibliography / appendix / None
        for block in self.flow:
            meta = {"zone": zone, "role": "body"}
            if block["kind"] == "tbl":
                self.paras.append(None)
                continue
            p = block["el"]
            text = para_text(p).strip()
            fields = para_fields(p)
            ppr = effective_ppr(p, self.styles)
            rpr = effective_rpr(p, self.styles)
            meta.update({"text": text, "fields": fields, "ppr": ppr, "rpr": rpr})

            drawings = p.findall(".//" + qn("w:drawing"))
            inlines = p.findall(".//" + qn("wp:inline"))
            anchors = p.findall(".//" + qn("wp:anchor"))
            meta["images"] = len(inlines)
            meta["anchors"] = len(anchors)
            meta["drawing"] = len(drawings)
            if inlines or anchors:
                meta["role"] = "image"
                meta["widths"] = [emu2cm(e.get("cx")) for e in
                                  p.findall(".//" + qn("wp:extent"))
                                  if e is not None and e.get("cx")]

            omath = p.findall(".//" + qn("m:oMath"))
            meta["omath"] = len(omath)
            meta["omathpara"] = len(p.findall(".//" + qn("m:oMathPara")))
            meta["ole"] = len(p.findall(".//" + w("object")))
            meta["pict"] = len(p.findall(".//" + w("pict")))

            # 标题
            lvl = None
            if ppr.get("outline_lvl") is not None and ppr["outline_lvl"] < 9:
                lvl = ppr["outline_lvl"] + 1
            else:
                sid = ppr.get("style_id")
                lvl = heading_level_of_style(self.styles, sid) if sid else None
            meta["style_heading_level"] = lvl
            text_lvl, pattern = self._heading_numbering_from_text(text)
            meta["text_heading_level"], meta["text_pattern"] = text_lvl, pattern
            textish = bool(text_lvl and len(text) < 40 and
                           not re.search(r"[。；，,;：:]", text))
            # 自动编号标题：Word 多级编号（一、二、…）在 XML 里，文本看不到编号
            numpr_headish = False
            if lvl is None and text and "num_id" in ppr and ppr.get("ilvl", 9) < 3 \
                    and len(text) < 30 and not re.search(r"[。；，,;：]", text) \
                    and rpr.get("bold"):
                numpr_headish = True
                lvl = ppr["ilvl"] + 1
            if lvl or textish or numpr_headish:
                meta["role"] = "heading"
                meta["level"] = lvl or text_lvl
                meta["auto_numbered"] = "num_id" in ppr
            special = _is_special_title(text) if len(text) < 30 else None
            if special:
                meta["role"] = "special"
                meta["special"] = special
                if special == "bibliography":
                    zone = "bibliography"
                elif special == "appendix":
                    zone = "appendix"
                elif meta.get("level") is None and special in ("abstract", "toc", "conclusion", "acknowledgement"):
                    zone = None

            # 题注
            m = CAPTION_RE.match(text) if text and len(text) < 100 else None
            seq = SEQ_FIELD_RE.search(fields)
            if seq and not m:
                label = seq.group(1)
                m2 = re.match(r"^(.{0,30})", text)
                meta["role"] = "caption"
                meta["caption_kind"] = "figure" if label in ("图", "Figure", "Fig") else (
                    "table" if label in ("表", "Table") else "equation")
                meta["caption_num"] = None
                meta["caption_sep"] = None
            elif m and (m.group(4) or seq):
                kind = {"图": "figure", "Figure": "figure", "Fig": "figure",
                        "表": "table", "Table": "table",
                        "式": "equation", "公式": "equation", "Eq": "equation"}.get(
                    m.group(1).rstrip("."), None)
                # 排除"图1结论：……"这类分析句：标题含句号/分号，或含逗号且偏长 → 不是题注
                title_g = m.group(4) or ""
                sentence_like = bool(re.search(r"[。；]", title_g)) or \
                    ("，" in title_g and len(text) > 30)
                if kind and not sentence_like:
                    meta["role"] = "caption"
                    meta["caption_kind"] = kind
                    meta["caption_num"] = m.group(2)
                    meta["caption_sep"] = _sep_from_match(text, m)
                    meta["caption_title"] = m.group(4)

            if meta["role"] == "body":
                meta["zone"] = zone
            self.paras.append(meta)
        # 连续编号段（1. ×× / 2. ×× / 3. ××，无标题样式）→ 实为列表项，降级为正文
        run_start = None
        for i in range(len(self.paras) + 1):
            m = self.paras[i] if i < len(self.paras) else None
            listish = bool(
                m and m.get("role") == "heading"
                and m.get("style_heading_level") is None
                and m.get("text_heading_level") == 1
                and len(m.get("text", "")) >= 12)
            if listish and run_start is None:
                run_start = i
            elif not listish and run_start is not None:
                if i - run_start >= 3:
                    for j in range(run_start, i):
                        self.paras[j]["role"] = "body"
                        self.paras[j].pop("level", None)
                run_start = None
        # 编号序列延续：无样式"N. ××"标题的下一段是"N+1."开头的正文 → 列表项，降级。
        # 倒序遍历：后面的项先降级为正文，前面的项才能观察到"下一项是正文"而连锁降级。
        for i in range(len(self.paras) - 1, -1, -1):
            m = self.paras[i]
            if not m or m.get("role") != "heading" \
                    or m.get("style_heading_level") is not None \
                    or m.get("text_heading_level") != 1:
                continue
            num_m = re.match(r"^(\d+)[.、．]", m.get("text", ""))
            if not num_m:
                continue
            nxt = num_m.group(1)
            for j in range(i + 1, min(i + 3, len(self.paras))):
                m2 = self.paras[j]
                if m2 and m2.get("text"):
                    if m2.get("role") == "body" and re.match(
                            rf"^{int(nxt) + 1}[.、．]", m2.get("text", "")):
                        m["role"] = "body"
                        m.pop("level", None)
                    break

    @staticmethod
    def _heading_numbering_from_text(text: str):
        if not text or len(text) > 60:
            return None, None
        if CHAP_CN_RE.match(text):
            grp = CHAP_CN_RE.match(text).group(1)
            return 1, "第{n}章" if grp.isdigit() else "第{c}章"
        m = DOTTED_NUM_RE.match(text)
        if m and re.match(r"^\d", text):
            depth = m.group(1).count(".") + 1
            if 1 <= depth <= 4:
                pat = {1: "{n}", 2: "{n}.{n2}",
                       3: "{n}.{n2}.{n3}", 4: "{n}.{n2}.{n3}.{n4}"}[depth]
                # 编号与标题之间的分隔符形态要进 pattern（1、与1. 与"1 "是三种体系）
                sep = re.match(r"^[\s、．.]*", text[len(m.group(1)):]).group(0)
                if "、" in sep:
                    pat += "、"
                elif depth == 1 and sep[:1] in (".", "．"):
                    pat += "."
                return depth, pat
        if CN_LIST_RE.match(text):
            return 2, "{c}、"
        if CN_PAREN_RE.match(text):
            return 3, "（{c2}）"
        return None, None

    # ------------------------------------------------------------- 节
    def sections(self):
        out = []
        for block in self.flow:
            if block["kind"] != "p":
                continue
            sect = block["el"].find(w("pPr") + "/" + w("sectPr"))
            if sect is not None:
                out.append(sect)
        final = self.body.find(w("sectPr"))
        if final is not None:
            out.append(final)
        return out

    def _sect_props(self, sect):
        pg = sect.find(w("pgSz"))
        mar = sect.find(w("pgMar"))
        cols = sect.find(w("cols"))
        num = sect.find(w("pgNumType"))
        props = {
            "orient": (pg.get(w("orient")) or "portrait") if pg is not None else "portrait",
            "w_cm": twips2cm(pg.get(w("w"))) if pg is not None and pg.get(w("w")) else None,
            "h_cm": twips2cm(pg.get(w("h"))) if pg is not None and pg.get(w("h")) else None,
            "margins": {k: twips2cm(mar.get(w(k))) for k in
                        ("top", "bottom", "left", "right", "gutter", "header", "footer")}
            if mar is not None else {},
            "cols": int(cols.get(w("num")) or 1) if cols is not None else 1,
            "fmt": num.get(w("fmt")) if num is not None else None,
            "start": num.get(w("start")) if num is not None else None,
            "titlePg": sect.find(w("titlePg")) is not None,
            "footer_ids": [r.get(qn("r:id")) for r in
                           sect.findall(w("footerReference"))],
            "header_ids": [r.get(qn("r:id")) for r in
                           sect.findall(w("headerReference"))],
        }
        return props

    def _footer_part(self, sect_props):
        for rid in reversed(sect_props.get("footer_ids") or []):
            target = self.rels.get(rid)
            if target:
                part = self.parts.get("word/" + os.path.basename(target))
                if part is not None:
                    return part
        return None

    # ------------------------------------------------------------- 档案聚合
    def profile_steps(self):
        """流式检测：逐维度 yield {"key","label","data"}，供界面边检测边列出。"""
        yield {"key": "page", "label": "页面与边距", "data": self._page_profile()}
        pn, hd = self._pagenum_header_profile()
        yield {"key": "page_numbers", "label": "页码设置", "data": pn}
        yield {"key": "headers", "label": "页眉页脚", "data": hd}
        yield {"key": "headings", "label": "标题体系", "data": self._headings_profile()}
        yield {"key": "body", "label": "正文格式", "data": self._body_profile()}
        yield {"key": "figures", "label": "图片", "data": self._figures_profile()}
        yield {"key": "captions", "label": "题注", "data": self._captions_profile()}
        yield {"key": "tables", "label": "表格", "data": self._tables_profile()}
        yield {"key": "formulas", "label": "公式", "data": self._formulas_profile()}
        yield {"key": "code_blocks", "label": "代码块", "data": self._code_profile()}
        yield {"key": "bibliography", "label": "参考文献",
               "data": self._bibliography_profile()}
        toc, xref = self._toc_xref_profile()
        yield {"key": "toc", "label": "目录", "data": toc}
        yield {"key": "cross_references", "label": "交叉引用", "data": xref}

    def profile(self) -> dict:
        prof = {
            "schema_version": "1.0",
            "meta": {"source_file": os.path.basename(self.path),
                     "extracted_at": now_iso(), "generator": "extract_format.py 1.0",
                     "doc_stats": {}},
            "flags": self.flags,
            "dimensions_absent": [],
        }
        for step in self.profile_steps():
            prof[step["key"]] = step["data"]
        self._instruction_conflict_flags(prof)
        self._doc_stats(prof)
        self._absent(prof)
        return prof

    SIZE_TO_HAO = {42: "初号", 36: "小初", 26: "一号", 24: "小一", 22: "二号",
                   18: "小二", 16: "三号", 15: "小三", 14: "四号", 12: "小四",
                   10.5: "五号", 9: "小五", 7.5: "六号"}

    def _instruction_conflict_flags(self, prof):
        """模板说明文字（如"正文五号宋体"）与实测格式的矛盾 → flag 提示用户。"""
        blob = "；".join(m.get("text", "") for m in self.paras if m)[:6000]
        for m in re.finditer(r"(标题|正文)[^。；，]{0,8}?([一二三四五六]号|小[一二三四五六]号)", blob):
            part, hao = m.group(1), m.group(2)
            node = prof.get("body", {}) if part == "正文" else \
                (prof.get("headings") or {}).get("level1", {})
            size = (node or {}).get("size_pt")
            size = size.get("value") if isinstance(size, dict) else size
            actual_hao = self.SIZE_TO_HAO.get(round(size)) if size else None
            if actual_hao and actual_hao != hao:
                self.flags.append(
                    f"template-instruction-conflict: 说明文字称{part}应为{hao}，"
                    f"但模板实例实测为 {actual_hao}（{size}pt）——请确认以哪个为准")

    def _doc_stats(self, prof):
        ps = prof["meta"]["doc_stats"]
        ps["sections"] = len(self.sections())
        ps["paragraphs"] = sum(1 for b in self.flow if b["kind"] == "p")
        ps["images"] = sum(m.get("images", 0) + m.get("anchors", 0)
                           for m in self.paras if m)
        ps["tables"] = sum(1 for b in self.flow if b["kind"] == "tbl")
        ps["captions_figure"] = sum(1 for m in self.paras
                                    if m and m.get("caption_kind") == "figure")
        ps["captions_table"] = sum(1 for m in self.paras
                                   if m and m.get("caption_kind") == "table")
        ps["formulas_omml"] = sum(m.get("omath", 0) for m in self.paras if m)
        ps["formulas_ole"] = sum(m.get("ole", 0) for m in self.paras if m)
        ps["code_blocks"] = len(self._code_blocks())

    def _absent(self, prof):
        absent = []
        st = prof["meta"]["doc_stats"]
        if st["formulas_omml"] + st["formulas_ole"] == 0:
            absent.append("formulas")
        if st["code_blocks"] == 0:
            absent.append("code_blocks")
        if not prof["toc"].get("mode"):
            absent.append("toc")
        if not prof["bibliography"].get("detected"):
            absent.append("bibliography")
        if st["images"] == 0:
            absent.append("figures")
        if st["tables"] == 0:
            absent.append("tables")
        prof["dimensions_absent"] = absent

    def _page_profile(self):
        sects = [self._sect_props(s) for s in self.sections()]
        body_sect = sects[-1] if sects else {}
        margins = body_sect.get("margins", {})
        orients = {s["orient"] for s in sects}
        cols = {s["cols"] for s in sects}
        return {
            "size": {"width_cm": body_sect.get("w_cm"), "height_cm": body_sect.get("h_cm"),
                     "orientation": "landscape" if orients and orients == {"landscape"} else "portrait",
                     "mixed_orientation": len(orients) > 1,
                     "confidence": "high" if sects else "low",
                     "evidence": f"{len(sects)}/{len(sects)} 节"},
            "margins": {"top_cm": margins.get("top"), "bottom_cm": margins.get("bottom"),
                        "left_cm": margins.get("left"), "right_cm": margins.get("right"),
                        "gutter_cm": margins.get("gutter"),
                        "confidence": "high" if sects else "low",
                        "evidence": f"{len(sects)}/{len(sects)} 节一致"},
            "header_footer_distance": {"header_cm": margins.get("header"),
                                       "footer_cm": margins.get("footer"),
                                       "confidence": "high" if sects else "low"},
            "columns": {"count": max(cols) if cols else 1,
                        "confidence": "high" if cols and cols <= {1, 2} else "medium"},
        }

    def _pagenum_position(self, footer_part):
        if footer_part is None:
            return None, None, None, None
        for p in footer_part.iter(w("p")):
            if PAGE_FIELD_RE.search(para_fields(p)):
                ppr = effective_ppr(p, self.styles)
                rpr = effective_rpr(p, self.styles)
                jc = ppr.get("align", "left")
                pos = {"center": "footer-center", "right": "footer-right",
                       "left": "footer-left"}.get(jc, "footer-" + jc)
                instr = para_fields(p)
                return pos, rpr.get("font_eastasia"), rpr.get("size_pt"), instr
        return None, None, None, None

    def _pagenum_header_profile(self):
        sects = self.sections()
        props = [self._sect_props(s) for s in sects]
        out_num = {"cover": {"has_page_number": None, "confidence": "low"},
                   "front_matter": None, "body": None}
        FMT_NORMALIZE = {"arabic": "decimal", "roman": "lowerRoman",
                         "ROMAN": "upperRoman", "CHinese": "chineseCounting"}
        for i, sp in enumerate(props):
            footer = self._footer_part(sp)
            has_field = False
            fmt_from_field = None
            if footer is not None:
                for p in footer.iter(w("p")):
                    instr = para_fields(p)
                    if PAGE_FIELD_RE.search(instr):
                        has_field = True
                        m = re.search(r"\\\*\s*(ROMAN|roman|Arabic|arabic|CHinese)", instr)
                        if m:
                            fmt_from_field = FMT_NORMALIZE.get(m.group(1), m.group(1))
                        break
            entry = {"format": fmt_from_field or sp["fmt"] or "decimal",
                     "start": int(sp["start"]) if sp["start"] else None}
            pos, fnt, sz, _ = self._pagenum_position(footer)
            entry.update({"position": pos, "font": fnt, "size_pt": sz})
            fmt_l = (entry["format"] or "").lower()
            roman = "oman" in fmt_l or fmt_l.startswith("chinese")
            if i == 0 and not has_field:
                out_num["cover"] = {"has_page_number": False, "confidence": "medium",
                                    "evidence": "第 1 节页脚无页码域"}
            elif roman and out_num["front_matter"] is None:
                out_num["front_matter"] = entry
            elif has_field:
                out_num["body"] = entry
        if out_num["body"] is None and out_num["front_matter"]:
            out_num["body"] = dict(out_num["front_matter"], format="decimal")
        if out_num["cover"]["has_page_number"] is None:
            out_num["cover"] = {"has_page_number": False, "confidence": "low"}
        # 奇偶页脚 + 左右分布 → 外侧页码
        evenodd = self.settings is not None and \
            self.settings.find(w("evenAndOddHeaders")) is not None
        if evenodd and out_num["body"] and \
                out_num["body"].get("position") in ("footer-left", "footer-right"):
            out_num["body"]["position"] = "footer-outside"
        headers = {
            "even_odd_different": self.settings is not None and
            self.settings.find(w("evenAndOddHeaders")) is not None,
            "first_page_different": any(sp["titlePg"] for sp in props),
            "content_font": None, "content_size_pt": None, "has_rule_line": False,
        }
        for i, sp in enumerate(props):
            for rid in reversed(sp.get("header_ids") or []):
                target = self.rels.get(rid)
                part = self.parts.get("word/" + os.path.basename(target)) if target else None
                if part is None:
                    continue
                for p in part.iter(w("p")):
                    if para_text(p).strip():
                        rpr = effective_rpr(p, self.styles)
                        headers["content_font"] = rpr.get("font_eastasia") or rpr.get("font_ascii")
                        headers["content_size_pt"] = rpr.get("size_pt")
                        ppr = p.find(w("pPr"))
                        if ppr is not None and ppr.find(w("pBdr")) is not None:
                            headers["has_rule_line"] = True
                        break
                break
            if headers["content_font"]:
                break
        headers["confidence"] = "medium"
        return out_num, headers

    def _lvl_abstract(self, num_id, ilvl):
        if self.numbering is None or not num_id:
            return None
        for num in self.numbering.findall(w("num")):
            if num.get(w("numId")) == num_id:
                aid = num.find(w("abstractNumId"))
                if aid is None:
                    return None
                for ab in self.numbering.findall(w("abstractNum")):
                    if ab.get(w("abstractNumId")) == aid.get(w("val")):
                        for lvl in ab.findall(w("lvl")):
                            if int(lvl.get(w("ilvl"), "0")) == (ilvl or 0):
                                lt = lvl.find(w("lvlText"))
                                nf = lvl.find(w("numFmt"))
                                txt = lt.get(w("val")) if lt is not None else ""
                                # 中文数字编号用 {c} 占位符，与文本检测统一
                                fmt_val = (nf.get(w("val")) or "").lower() if nf is not None else ""
                                ph = "{c}" if fmt_val.startswith("chinese") else "{n}"
                                txt = (txt.replace("%1", ph).replace("%2", "{n2}")
                                       .replace("%3", "{n3}").replace("%4", "{n4}"))
                                return {"pattern": txt or None,
                                        "numFmt": nf.get(w("val")) if nf is not None else None}
        return None

    def _headings_profile(self):
        out = {"levels_in_use": []}
        by_level = {}
        for m in self.paras:
            if not m or m.get("role") != "heading":
                continue
            by_level.setdefault(m["level"], []).append(m)
        for lvl in sorted(by_level):
            items = by_level[lvl]
            rprs = [m["rpr"] for m in items]
            pprs = [m["ppr"] for m in items]
            numberings = []
            for m in items:
                if m.get("auto_numbered"):
                    ab = self._lvl_abstract(m["ppr"].get("num_id"), m["ppr"].get("ilvl"))
                    if ab:
                        numberings.append(("auto-numbering", ab["pattern"]))
                elif m.get("text_pattern"):
                    numberings.append(("manual", m["text_pattern"]))
            num_mode, num_pat = None, None
            if numberings:
                cnt = {}
                for mode, pat in numberings:
                    cnt[(mode, pat)] = cnt.get((mode, pat), 0) + 1
                (num_mode, num_pat), c = max(cnt.items(), key=lambda kv: kv[1])
                conf = "high" if c == len(numberings) and c >= 3 else (
                    "medium" if c >= 2 else "low")
            else:
                conf = "low"
            def agg(key, tol=0.0, default=None):
                md = mode_with_conf([r.get(key) for r in rprs], tol)
                return conf_result(md) if md else default

            def aggp(key, tol=0.0, default=None):
                md = mode_with_conf([r.get(key) for r in pprs], tol)
                return conf_result(md) if md else default
            out["levels_in_use"].append(lvl)
            out[f"level{lvl}"] = {
                "style_name": items[0]["ppr"].get("style_id"),
                "font_eastasia": agg("font_eastasia"),
                "font_ascii": agg("font_ascii"),
                "size_pt": agg("size_pt", 0.5),
                "bold": agg("bold"),
                "italic": agg("italic"),
                "align": aggp("align"),
                "space_before_pt": aggp("space_before_pt", 0.5),
                "space_after_pt": aggp("space_after_pt", 0.5),
                "line_spacing": None,
                "page_break_before": aggp("page_break_before"),
                "keep_next": aggp("keep_next"),
                "keep_lines": aggp("keep_lines"),
                "numbering": {"mode": num_mode, "pattern": num_pat,
                              "chapter_linked": bool(num_pat and
                                                     ("{n}" in num_pat and "{n2}" in num_pat or
                                                      num_pat.startswith("第"))),
                              "confidence": conf,
                              "evidence": f"{len(numberings)}/{len(items)} 个标题带编号"},
            }
            ls = mode_with_conf([r.get("line_spacing") and
                                 (r["line_spacing"]["rule"], r["line_spacing"]["value"])
                                 for r in pprs if r.get("line_spacing")])
            if ls:
                rule, val = ls["value"]
                out[f"level{lvl}"]["line_spacing"] = {"rule": rule, "value": val}
        out["levels_in_use"].sort()
        return out

    def _body_profile(self):
        samples = [m for m in self.paras
                   if m and m.get("role") == "body" and m.get("zone") is None
                   and len(m.get("text", "")) >= 15]
        if not samples:
            samples = [m for m in self.paras if m and m.get("role") == "body"]
        n = len(samples)

        def agg(key, tol=0.0):
            return conf_result(mode_with_conf([m["rpr"].get(key) for m in samples], tol))
        ind_chars = []
        for m in samples:
            ppr = m["ppr"]
            if ppr.get("first_line_chars") is not None:
                ind_chars.append(ppr["first_line_chars"])
            elif ppr.get("first_line_twips"):
                sz = m["rpr"].get("size_pt") or 12
                ind_chars.append(round(ppr["first_line_twips"] / (sz * 20), 1))
        ind_md = mode_with_conf(ind_chars, 0.1)
        ls_md = mode_with_conf([m["ppr"].get("line_spacing") and
                                (m["ppr"]["line_spacing"]["rule"],
                                 m["ppr"]["line_spacing"]["value"])
                                for m in samples if m["ppr"].get("line_spacing")])
        return {
            "font_eastasia": agg("font_eastasia"), "font_ascii": agg("font_ascii"),
            "size_pt": agg("size_pt", 0.5), "bold": agg("bold"),
            "line_spacing": {"rule": ls_md["value"][0], "value": ls_md["value"][1]}
            if ls_md else None,
            "first_line_indent_chars": conf_result(ind_md, " 段"),
            "align": conf_result(mode_with_conf([m["ppr"].get("align") for m in samples])),
            "widow_control": agg("widow_control"),
            "confidence": "high" if n >= 10 else ("medium" if n >= 3 else "low"),
            "evidence": f"{n} 段正文样本",
        }

    def _figures_profile(self):
        imgs = [m for m in self.paras if m and m.get("images")]
        widths = [wd for m in imgs for wd in m.get("widths", [])]
        wmd = mode_with_conf(widths, 0.1)
        flag_outliers("figure-width", widths, 0.1, wmd, self.flags, "cm")
        align = conf_result(mode_with_conf([m["ppr"].get("align") for m in imgs]))
        keep = conf_result(mode_with_conf([m["ppr"].get("keep_next") for m in imgs]))
        # 段距只有在与正文默认不同时才作为规则（避免继承噪音）
        body_samples = [m for m in self.paras
                        if m and m.get("role") == "body" and m.get("zone") is None
                        and len(m.get("text", "")) >= 15]
        b_after = mode_with_conf([m["ppr"].get("space_after_pt") for m in body_samples], 0.5)
        b_before = mode_with_conf([m["ppr"].get("space_before_pt") for m in body_samples], 0.5)
        f_after = mode_with_conf([m["ppr"].get("space_after_pt") for m in imgs], 0.5)
        f_before = mode_with_conf([m["ppr"].get("space_before_pt") for m in imgs], 0.5)
        sa = f_after if (f_after and b_after and
                         abs(f_after["value"] - b_after["value"]) > 0.5) else None
        sb = f_before if (f_before and b_before and
                          abs(f_before["value"] - b_before["value"]) > 0.5) else None
        if any(m.get("anchors") for m in self.paras if m):
            self.flags.append("document-contains-floating-anchors: 存在浮动图片，应用时保持锚定不改嵌入")
        return {
            "width_rule": {"type": "fixed", **(conf_result(wmd, " 张图") or {}),
                           } if wmd else None,
            "align": align, "keep_with_caption": keep,
            "space_before_pt": conf_result(sb), "space_after_pt": conf_result(sa),
        }

    def _nearest_object(self, idx, kind):
        """向前/向后 2 个块内找最近的图段或表格，返回 ('below'|'above', obj_idx)。"""
        target = "image" if kind == "figure" else "tbl"
        for d in (1, 2):
            j = idx - d
            if 0 <= j < len(self.flow):
                b = self.flow[j]
                hit = (b["kind"] == "p" and self.paras[j] and self.paras[j].get("images")
                       and target == "image") or (b["kind"] == target if target == "tbl" else False)
                if hit:
                    return "below", j
            k = idx + d
            if 0 <= k < len(self.flow):
                b = self.flow[k]
                hit = (b["kind"] == "p" and self.paras[k] and self.paras[k].get("images")
                       and target == "image") or (b["kind"] == target if target == "tbl" else False)
                if hit:
                    return "above", k
        return None, None

    def _captions_profile(self):
        out = {}
        for kind in ("figure", "table"):
            caps = [m for m in self.paras if m and m.get("caption_kind") == kind]
            pos, nums_chapter, seps, rprs, modes = [], [], [], [], []
            for i, m in enumerate(caps):
                flow_idx = self._flow_idx_of_caption(m)
                if flow_idx is None:
                    continue
                where, _ = self._nearest_object(flow_idx, kind)
                if where:
                    pos.append(where)
                num = m.get("caption_num")
                if num:
                    if re.search(r"[-–.．]", num):
                        nums_chapter.append(True)
                    else:
                        nums_chapter.append(False)
                if m.get("caption_sep") is not None:
                    seps.append(m["caption_sep"])
                f = m.get("fields", "")
                if SEQ_FIELD_RE.search(f):
                    modes.append("field-seq-styleref" if STYLEREF_RE.search(f) else "field-seq")
                elif num:
                    modes.append("manual")
            rprs = [m["rpr"] for m in caps]
            pprs = [m["ppr"] for m in caps]

            def agg(key, tol=0.0):
                return conf_result(mode_with_conf([r.get(key) for r in rprs], tol))
            align_md = mode_with_conf([r.get("align") for r in pprs])
            ch_md = mode_with_conf(nums_chapter)
            mode_md = mode_with_conf(modes)
            pos_md = mode_with_conf(pos)
            sep_md = mode_with_conf(seps)
            label = {"figure": "图", "table": "表"}[kind]
            # 域里含 STYLEREF 即为按章编号（域字面结果可能是未更新的旧值，不能作数）
            any_styleref = any(STYLEREF_RE.search(m.get("fields", "")) for m in caps)
            chapter_linked = any_styleref or bool(ch_md and ch_md["value"])
            if sep_md:
                sep = sep_md["value"]
            else:
                sep = " "
            if chapter_linked:
                template = f"{label}{{chapter}}-{{index}}{{sep}}{{text}}"
            else:
                template = f"{label}{{index}}{{sep}}{{text}}"
            num_mode = mode_md["value"] if mode_md else "manual"
            n_total = len(caps)
            conf = "high" if n_total >= 3 and pos_md and pos_md["ratio"] >= 0.9 else (
                "medium" if n_total >= 1 else None)
            out[kind] = {
                "label": label,
                "position": pos_md["value"] if pos_md else ("below" if kind == "figure" else "above"),
                "numbering": {"mode": num_mode, "chapter_linked": chapter_linked},
                "format_template": template, "separator": sep,
                "font_eastasia": agg("font_eastasia"), "font_ascii": agg("font_ascii"),
                "size_pt": agg("size_pt", 0.5), "bold": agg("bold"),
                "align": conf_result(align_md),
                "confidence": conf or "low",
                "evidence": f"{n_total} 个{label}题注" + (
                    f"，位置 {pos_md['count']}/{pos_md['total']}" if pos_md else "") + (
                    f"，分隔符 {sep_md['count']}/{sep_md['total']}" if sep_md else ""),
            }
            if n_total == 0:
                out[kind] = None
        # 公式编号：含公式（oMath/oMathPara）的段落，取尾部编号
        eqnums = []
        for m in self.paras:
            if m and (m.get("omath") or m.get("omathpara")):
                mm = EQNUM_RE.search(m.get("text", ""))
                if mm:
                    eqnums.append(f"({mm.group(1)}-{mm.group(2)})" if mm.group(2)
                                  else f"({mm.group(1)})")
        eq_md = mode_with_conf(eqnums)
        out["equation"] = {
            "template": eq_md["value"] if eq_md else None,
            "position": "right",
            "confidence": "medium" if eq_md else "low",
            "evidence": f"{len(eqnums)} 个编号公式",
        }
        out = {k: v for k, v in out.items() if v is not None or k == "figure" or k == "table"}
        if out.get("figure") is None:
            out.pop("figure", None)
        if out.get("table") is None:
            out.pop("table", None)
        return out

    def _flow_idx_of_caption(self, meta):
        for i, m in enumerate(self.paras):
            if m is meta:
                return i
        return None

    def _tables_profile(self):
        code_tbl_idxs = {bl[0]["idx"] for bl in self._code_blocks()
                         if bl[0].get("tbl_code")}
        tbl_blocks = [b for b in self.flow
                      if b["kind"] == "tbl" and b["idx"] not in code_tbl_idxs]
        three_line, header_bold, header_align, shading = [], [], [], []
        repeat, cantsplit, cell_sizes, cell_fonts, cell_aligns = [], [], [], [], []
        widths = []
        tl_tops, tl_bottoms, tl_headers = [], [], []
        border_kinds = []
        for b in tbl_blocks:
            tbl = b["el"]
            tb = tbl.find(w("tblPr"))
            borders = {}
            bEl = tb.find(w("tblBorders")) if tb is not None else None
            for s in ("top", "bottom", "left", "right", "insideH", "insideV"):
                borders[s] = border_sz(bEl, s)
            first_row = tbl.find(w("tr"))
            header_bottom = None
            if first_row is not None:
                bts = []
                for tc in first_row.findall(w("tc")):
                    v = border_sz(tc.find(w("tcPr") + "/" + w("tcBorders")), "bottom")
                    if v:
                        bts.append(v)
                hbm = mode_with_conf(bts)
                header_bottom = hbm["value"] if hbm else None
                hdr_cells = first_row.findall(w("tc"))
                hb, ha, hsh = [], [], []
                for tc in hdr_cells[:6]:
                    p0 = tc.find(".//" + w("p"))
                    if p0 is None:
                        continue
                    rpr = effective_rpr(p0, self.styles)
                    ppr = effective_ppr(p0, self.styles)
                    hb.append(rpr.get("bold"))
                    ha.append(ppr.get("align"))
                    shd = tc.find(w("tcPr") + "/" + w("shd"))
                    hsh.append(shd.get(w("fill")) if shd is not None and
                               shd.get(w("fill")) not in (None, "auto", "FFFFFF") else None)
                for lst, val in ((header_bold, hb), (header_align, ha), (shading, hsh)):
                    md = mode_with_conf(val)
                    if md:
                        lst.append(md["value"])
                trpr = first_row.find(w("trPr"))
                repeat.append(trpr is not None and
                              trpr.find(w("tblHeader")) is not None)
            rows = tbl.findall(w("tr"))
            cs = [r.find(w("trPr")) is not None and
                  r.find(w("trPr")).find(w("cantSplit")) is not None for r in rows]
            cantsplit.append(all(cs) if cs else False)
            is3 = (borders["top"] and borders["top"] >= 0.5 and
                   borders["bottom"] and borders["bottom"] >= 0.5 and
                   not borders["insideV"] and not borders["left"] and not borders["right"] and
                   (not borders["insideH"] or header_bottom))
            three_line.append(bool(is3))
            if is3:
                tl_tops.append(borders["top"])
                tl_bottoms.append(borders["bottom"])
                if header_bottom:
                    tl_headers.append(header_bottom)
            # 边框形态分类（不止三线：模板不是三线时同样要能规整目标）
            has_v = bool(borders["insideV"] or borders["left"] or borders["right"])
            if is3:
                border_kinds.append("three-line")
            elif has_v and borders["insideH"]:
                border_kinds.append("grid")
            elif borders["top"] and borders["bottom"] and not has_v:
                border_kinds.append("outline")
            elif not (borders["top"] or borders["bottom"] or has_v or borders["insideH"]):
                border_kinds.append("none")
            else:
                border_kinds.append("other")
            # 表内字体抽样
            for tc in tbl.findall(".//" + w("tc"))[:12]:
                p0 = tc.find(".//" + w("p"))
                if p0 is None:
                    continue
                rpr = effective_rpr(p0, self.styles)
                ppr = effective_ppr(p0, self.styles)
                cell_sizes.append(rpr.get("size_pt"))
                cell_fonts.append(rpr.get("font_eastasia") or rpr.get("font_ascii"))
                cell_aligns.append(ppr.get("align"))
            tblw = tb.find(w("tblW")) if tb is not None else None
            if tblw is not None and tblw.get(w("type")) == "pct":
                widths.append({"type": "pct", "value": int(tblw.get(w("w")) or 0)})
            elif tblw is not None and tblw.get(w("type")) == "dxa":
                widths.append({"type": "dxa", "value": int(tblw.get(w("w")) or 0)})
            else:
                widths.append({"type": "auto", "value": None})
        n = len(tbl_blocks)
        tl_md = mode_with_conf(three_line)
        cont = CONTINUATION_RE.search("".join(
            m.get("text", "") or "" for m in self.paras if m))

        def conf_(md, unit=""):
            return conf_result(md, unit) if md else None
        width_md = mode_with_conf([str(x) for x in widths]) if widths else None
        top_md = mode_with_conf(tl_tops, 0.1)
        bot_md = mode_with_conf(tl_bottoms, 0.1)
        hdr_md = mode_with_conf(tl_headers, 0.1)
        rule_md = mode_with_conf(border_kinds)
        border_rule = rule_md["value"] if rule_md else None
        return {
            "border_rule": border_rule,
            "three_line": {"enabled": border_rule == "three-line",
                           "top_bottom_pt": top_md["value"] if top_md else None,
                           "header_underline_pt": hdr_md["value"] if hdr_md else None,
                           **({"confidence": tl_md["confidence"],
                               "evidence": f"{tl_md['count']}/{n} 张表为三线表"}
                              if tl_md and n else {}),
                           } if n else None,
            "header_row": {"bold": conf_(mode_with_conf(header_bold)),
                           "align": conf_(mode_with_conf(header_align)),
                           "shading_fill": conf_(mode_with_conf(shading))},
            "repeat_header": conf_(mode_with_conf(repeat)),
            "cant_split_rows": conf_(mode_with_conf(cantsplit)),
            "keep_caption_with_table": None,
            "continuation_convention": {"detected": bool(cont), "label": "续表" if cont else None},
            "cell_font_eastasia": conf_(mode_with_conf(cell_fonts)),
            "cell_size_pt": conf_(mode_with_conf(cell_sizes, 0.5)),
            "cell_align": conf_(mode_with_conf(cell_aligns)),
            "width_mode": eval(width_md["value"]) if width_md else None,  # noqa: S307
        }

    def _formulas_profile(self):
        rprs = []
        omml_runs = self.doc.findall(".//" + qn("m:r"))
        for r in omml_runs[:80]:
            rpr = parse_rpr(r.find(w("rPr")))
            if rpr:
                rprs.append(rpr)
        n_omath = len(self.doc.findall(".//" + qn("m:oMath")))
        n_omathpara = len(self.doc.findall(".//" + qn("m:oMathPara")))
        n_ole = len(self.doc.findall(".//" + w("object")))
        n_pict = len(self.doc.findall(".//" + w("pict")))
        if n_ole:
            self.flags.append(f"ole-formulas: 检测到 {n_ole} 个 OLE 公式对象，只能识别不能自动改")
        if n_pict:
            self.flags.append(f"picture-formulas: 检测到 {n_pict} 个图片公式，只能识别")
        fnt = conf_result(mode_with_conf([r.get("font_ascii") for r in rprs]))
        sz = conf_result(mode_with_conf([r.get("size_pt") for r in rprs], 0.5))
        return {
            "omml_font": fnt, "omml_size_pt": sz,
            "align": "center", "number_position": "right",
            "ole_detected": n_ole > 0, "pict_detected": n_pict > 0,
            "counts": {"omml": n_omath, "omathpara": n_omathpara, "ole": n_ole, "pict": n_pict},
        }

    def _code_blocks(self):
        blocks, cur = [], []
        for i, m in enumerate(self.paras):
            if not m:
                if cur:
                    blocks.append(cur)
                    cur = []
                continue
            score = 0
            fa = m["rpr"].get("font_ascii")
            if is_mono_font(fa):
                score += 2
            p_el = self.flow[i]["el"]
            ppr = p_el.find(w("pPr"))
            has_shd = ppr is not None and ppr.find(w("shd")) is not None and \
                ppr.find(w("shd")).get(w("fill")) not in (None, "auto", "FFFFFF")
            has_pbdr = ppr is not None and ppr.find(w("pBdr")) is not None
            if has_shd:
                score += 1
            if has_pbdr:
                score += 1
            txt = m.get("text", "")
            if code_signal_score(txt) > 0.08:
                score += 1
            in_tbl = False
            if score >= 2:
                cur.append({"idx": i, "mono": is_mono_font(fa), "shd": has_shd,
                            "pbdr": has_pbdr, "in_tbl": in_tbl, "text": txt[:40]})
            else:
                if cur:
                    blocks.append(cur)
                    cur = []
        if cur:
            blocks.append(cur)
        # 单格表代码框
        for bi, b in enumerate(self.flow):
            if b["kind"] != "tbl":
                continue
            tbl = b["el"]
            rows = tbl.findall(w("tr"))
            if len(rows) == 1:
                cells = rows[0].findall(w("tc"))
                if len(cells) == 1:
                    ps = cells[0].findall(".//" + w("p"))
                    mono = sum(1 for p in ps if is_mono_font(effective_rpr(p, self.styles).get("font_ascii")))
                    if ps and mono >= max(1, len(ps) // 2):
                        blocks.append([{"idx": bi, "tbl_code": True, "mono": True}])
        return [bl for bl in blocks if len(bl) >= 1 or (bl and bl[0].get("tbl_code"))]

    def _code_profile(self):
        blocks = self._code_blocks()
        if not blocks:
            return {"detected": False, "confidence": "low"}
        tbl_box = any(b[0].get("tbl_code") for b in blocks)
        para_box = any(not b[0].get("tbl_code") and (b[0].get("shd") or b[0].get("pbdr"))
                       for b in blocks)
        fonts, sizes = [], []
        for b in blocks:
            for item in b:
                if item.get("tbl_code"):
                    tbl = self.flow[item["idx"]]["el"]
                    for p0 in tbl.findall(".//" + w("p"))[:10]:
                        rpr = effective_rpr(p0, self.styles)
                        fonts.append(rpr.get("font_ascii"))
                        sizes.append(rpr.get("size_pt"))
                    continue
                m = self.paras[item["idx"]]
                fonts.append(m["rpr"].get("font_ascii"))
                sizes.append(m["rpr"].get("size_pt"))
        fnt = conf_result(mode_with_conf(fonts))
        sz = conf_result(mode_with_conf(sizes, 0.5))
        shd_fill = None
        for b in blocks:
            if b[0].get("tbl_code"):
                continue
            ppr = self.flow[b[0]["idx"]]["el"].find(w("pPr"))
            if ppr is not None:
                shd = ppr.find(w("shd"))
                if shd is not None and shd.get(w("fill")) not in (None, "auto"):
                    shd_fill = shd.get(w("fill"))
                    break
        return {
            "detected": True,
            "box_style": "single-cell-table" if tbl_box else (
                "paragraph-border" if para_box else "none"),
            "font_ascii": fnt, "size_pt": sz,
            "line_spacing_single": True, "shading_fill": shd_fill,
            "confidence": "medium" if len(blocks) >= 2 else "low",
            "evidence": f"{len(blocks)} 个代码块",
        }

    def _bibliography_profile(self):
        items = [m for m in self.paras if m and m.get("zone") == "bibliography"
                 and m.get("role") == "body" and len(m.get("text", "")) > 5]
        if not items:
            return {"detected": False}
        fmts, hangs, sizes = [], [], []
        for m in items:
            t = m["text"]
            if BIB_NUMLIST_RE.match(t):
                fmts.append("[{n}]")
            elif BIB_NUMLIST2_RE.match(t):
                fmts.append("({n})")
            elif BIB_NUMLIST3_RE.match(t):
                fmts.append("{n}.")
            ppr = m["ppr"]
            if ppr.get("hanging_chars"):
                hangs.append(ppr["hanging_chars"])
            elif ppr.get("hanging_twips"):
                hangs.append(round(ppr["hanging_twips"] / 240.0, 1))
            sizes.append(m["rpr"].get("size_pt"))
        fmt_md = mode_with_conf(fmts)
        return {
            "detected": True,
            "numbering_format": fmt_md["value"] if fmt_md else None,
            "hanging_indent_chars": conf_result(mode_with_conf(hangs, 0.1), " 条"),
            "size_pt": conf_result(mode_with_conf(sizes, 0.5)),
            "align": conf_result(mode_with_conf([m["ppr"].get("align") for m in items])),
            "count": len(items),
        }

    def _toc_xref_profile(self):
        instr_all = []
        for el in self.doc.iter():
            if el.tag == w("instrText"):
                instr_all.append(el.text or "")
            elif el.tag == w("fldSimple"):
                instr_all.append(el.get(w("instr")) or "")
        blob = " ".join(instr_all)
        m = TOC_FIELD_RE.search(blob)
        toc = {"mode": "field" if m else None, "depth": int(m.group(2)) if m else None,
               "leader": "dot"}
        if not m:
            for meta in self.paras:
                if meta and meta.get("special") == "toc":
                    toc = {"mode": "manual", "depth": None, "leader": "dot"}
                    break
        ref_fields = sum(1 for s in instr_all if re.match(r"\s*REF\s", s))
        text_xrefs = 0
        for meta in self.paras:
            if meta and meta.get("role") == "body":
                text_xrefs += len(XREF_TEXT_RE.findall(meta.get("text", "")))
        xref = {"mode": "field" if ref_fields >= text_xrefs or ref_fields > 0 else "text",
                "count": max(ref_fields, text_xrefs)}
        return toc, xref

    # ------------------------------------------------------------- 元素映射
    def element_map(self):
        headings, captions, figures, tables, specials, ambiguities = [], [], [], [], [], []
        body_size = (self._body_profile().get("size_pt") or {}).get("value") or 12
        for i, m in enumerate(self.paras):
            if not m:
                continue
            if m.get("role") == "heading":
                headings.append({"idx": i, "level": m["level"],
                                 "style_level": m.get("style_heading_level"),
                                 "text_level": m.get("text_heading_level"),
                                 "auto_numbered": m.get("auto_numbered"),
                                 "text": m.get("text", "")[:60]})
                if m.get("style_heading_level") is None:
                    ambiguities.append({"idx": i, "type": "heading-without-style",
                                        "text": m.get("text", "")[:40],
                                        "reason": "无标题样式，按文本编号推断"})
            elif m.get("role") == "caption":
                captions.append({"idx": i, "kind": m.get("caption_kind"),
                                 "text": m.get("text", "")[:60],
                                 "num": m.get("caption_num"), "sep": m.get("caption_sep")})
            elif m.get("role") == "image":
                figures.append({"idx": i, "inline": m.get("images"), "anchor": m.get("anchors"),
                                "widths": m.get("widths")})
            elif m.get("role") == "special":
                specials.append({"idx": i, "role": m.get("special"), "text": m.get("text", "")[:30]})
            elif m.get("role") == "body":
                t = m.get("text", "")
                rpr = m["rpr"]
                if t and len(t) < 30 and (rpr.get("bold") or
                                          (rpr.get("size_pt") or 0) > body_size + 1.5):
                    if not t.endswith(("。", "，", "；", "：")):
                        ambiguities.append({"idx": i, "type": "possible-heading",
                                            "text": t[:40],
                                            "reason": "加粗/大字号短段落，可能是未标样式标题"})
        for bi, b in enumerate(self.flow):
            if b["kind"] == "tbl":
                tbl = b["el"]
                tables.append({"flow_idx": bi, "rows": len(tbl.findall(w("tr"))),
                               "cols": len(tbl.find(w("tr")).findall(w("tc")))
                               if tbl.find(w("tr")) is not None else 0})
        return {
            "source_file": os.path.basename(self.path),
            "headings": headings, "captions": captions, "figures": figures,
            "tables": tables, "specials": specials,
            "code_blocks": [[it["idx"] for it in bl] for bl in self._code_blocks()],
            "ambiguities": ambiguities,
        }


# ---------------------------------------------------------------- 报告
ICON = {"high": "✅ 已确认", "medium": "⚠️ 基本确定", "low": "❓ 需要你拍板"}


def _cell(v):
    if isinstance(v, dict):
        val = v.get("value")
        conf = v.get("confidence")
        ev = v.get("evidence", "")
        s = f"{val}（{ICON.get(conf, '')} {ev}）" if conf else str(val)
        return s
    return str(v) if v is not None else "—"


def render_report(prof: dict) -> str:
    L = ["# 格式档案清单（模板检测结果）", ""]
    L.append("> ✅ 已确认=高置信可直接执行 · ⚠️ 基本确定=建议核对 · ❓ 需要你拍板=必问项")
    L.append("")
    pg = prof.get("page", {})
    L.append("## 一、页面与页码")
    L.append("| 项 | 检测值 |")
    L.append("|---|---|")
    L.append(f"| 纸张 | {_cell(pg.get('size', {}).get('width_cm'))} × "
             f"{_cell(pg.get('size', {}).get('height_cm'))} cm，"
             f"{pg.get('size', {}).get('orientation')} |" if pg.get("size") else "")
    mar = pg.get("margins", {})
    L.append(f"| 页边距 cm | 上 {mar.get('top_cm')} / 下 {mar.get('bottom_cm')} / "
             f"左 {mar.get('left_cm')} / 右 {mar.get('right_cm')} |" if mar else "")
    pn = prof.get("page_numbers", {})
    if pn.get("body"):
        L.append(f"| 正文页码 | {_cell(pn['body'].get('format'))} 从 "
                 f"{pn['body'].get('start')} 起，{_cell(pn['body'].get('position'))} |")
    if pn.get("front_matter"):
        L.append(f"| 前置页码 | {_cell(pn['front_matter'].get('format'))} |")
    L.append(f"| 封面页码 | {'无' if pn.get('cover', {}).get('has_page_number') is False else '有'} |")
    hd = prof.get("headings", {})
    L.append("")
    L.append("## 二、标题与正文")
    L.append("| 项 | 检测值 |")
    L.append("|---|---|")
    for lvl in hd.get("levels_in_use", []):
        h = hd.get(f"level{lvl}", {})
        num = h.get("numbering") or {}
        L.append(f"| {lvl} 级标题 | {_cell(h.get('font_eastasia'))} {_cell(h.get('size_pt'))}pt "
                 f"{'加粗' if (h.get('bold') or {}).get('value') else ''} "
                 f"{_cell(h.get('align'))}；编号：{num.get('mode')}/{num.get('pattern')} |")
    bd = prof.get("body", {})
    L.append(f"| 正文 | {_cell(bd.get('font_eastasia'))}+{_cell(bd.get('font_ascii'))} "
             f"{_cell(bd.get('size_pt'))}pt 行距 {_cell(bd.get('line_spacing'))} "
             f"缩进 {_cell(bd.get('first_line_indent_chars'))} |")
    L.append("")
    L.append("## 三、图 / 题注 / 表")
    L.append("| 项 | 检测值 |")
    L.append("|---|---|")
    fg = prof.get("figures", {})
    if fg.get("width_rule"):
        L.append(f"| 图片宽度 | {_cell(fg['width_rule'])}，对齐 {_cell(fg.get('align'))} |")
    for kind, name in (("figure", "图题注"), ("table", "表题注")):
        cp = (prof.get("captions") or {}).get(kind)
        if cp:
            L.append(f"| {name} | 位置-{cp.get('position')}，{_cell(cp.get('separator'))} 分隔，"
                     f"编号 {cp.get('numbering', {}).get('mode')}/按章="
                     f"{cp.get('numbering', {}).get('chapter_linked')}，"
                     f"{_cell(cp.get('size_pt'))}pt {_cell(cp.get('bold'))} |")
    tb = prof.get("tables", {})
    if tb.get("three_line"):
        L.append(f"| 三线表 | {'启用' if tb['three_line'].get('enabled') else '未启用'} "
                 f"上下线 {_cell({'value': tb['three_line'].get('top_bottom_pt')})}pt，"
                 f"表头线 {_cell({'value': tb['three_line'].get('header_underline_pt')})}pt |")
    if tb.get("continuation_convention", {}).get("detected"):
        L.append("| 续表惯例 | 检测到（将翻译为：整表不分页/跨页重复表头） |")
    L.append(f"| 表格跨页 | 重复表头 {_cell(tb.get('repeat_header'))}，行不拆分 {_cell(tb.get('cant_split_rows'))} |")
    L.append("")
    L.append("## 四、公式 / 代码 / 参考文献 / 目录")
    L.append("| 项 | 检测值 |")
    L.append("|---|---|")
    fm = prof.get("formulas", {})
    if fm:
        L.append(f"| 公式 | OMML {_cell(fm.get('omml_font'))} {_cell(fm.get('omml_size_pt'))}pt"
                 f"{'，含 OLE 公式（不自动改）' if fm.get('ole_detected') else ''} |")
    cb = prof.get("code_blocks", {})
    if cb.get("detected"):
        L.append(f"| 代码块 | {_cell({'value': cb.get('box_style')})}，"
                 f"{_cell(cb.get('font_ascii'))} {_cell(cb.get('size_pt'))}pt |")
    bib = prof.get("bibliography", {})
    if bib.get("detected"):
        L.append(f"| 参考文献 | {_cell({'value': bib.get('numbering_format')})} 编号，"
                 f"悬挂缩进 {_cell(bib.get('hanging_indent_chars'))}，"
                 f"{_cell(bib.get('size_pt'))}pt |")
    toc = prof.get("toc", {})
    if toc.get("mode"):
        L.append(f"| 目录 | {toc.get('mode')} 域，深度 {toc.get('depth')} |")
    if prof.get("dimensions_absent"):
        L.append("")
        L.append(f"**模板中未出现的维度（应用时跳过）**：{('、'.join(prof['dimensions_absent']))}")
    if prof.get("flags"):
        L.append("")
        L.append("## ⚠️ 模板异常提示")
        for f in prof["flags"]:
            L.append(f"- {f}")
    L.append("")
    st = prof.get("meta", {}).get("doc_stats", {})
    L.append(f"---\n*文档规模：{st.get('paragraphs', 0)} 段 / {st.get('images', 0)} 图 / "
             f"{st.get('tables', 0)} 表 / {st.get('sections', 0)} 节*")
    return "\n".join(x for x in L if x is not None)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("input")
    ap.add_argument("-o", "--out")
    ap.add_argument("--report")
    ap.add_argument("--map", dest="map_out")
    args = ap.parse_args()
    stem = os.path.splitext(args.input)[0]
    prof_out = args.out or stem + ".profile.json"
    rep_out = args.report or stem + ".report.md"
    map_out = args.map_out or stem + ".map.json"

    ex = Extractor(args.input)
    prof = ex.profile()
    with open(prof_out, "w", encoding="utf-8") as f:
        json.dump(prof, f, ensure_ascii=False, indent=2)
    with open(rep_out, "w", encoding="utf-8") as f:
        f.write(render_report(prof))
    with open(map_out, "w", encoding="utf-8") as f:
        json.dump(ex.element_map(), f, ensure_ascii=False, indent=2)
    print(f"profile -> {prof_out}")
    print(f"report -> {rep_out}")
    print(f"map    -> {map_out}")
    print(f"stats  -> {prof['meta']['doc_stats']}")


if __name__ == "__main__":
    main()
