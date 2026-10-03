# -*- coding: utf-8 -*-
"""serve_ui.py — docx-format-transfer 本地界面服务器

用法：python serve_ui.py [端口] [--host 0.0.0.0]
默认 127.0.0.1:8765（仅本机）；--host 0.0.0.0 时局域网内设备可经浏览器访问。
打开浏览器访问启动时打印的地址。
"""
from __future__ import annotations

import io
import json
import os
import re
import socket
import sys
import threading
import time
import urllib.parse
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

from apply_format import (  # noqa: E402
    Applier, RULES, DEFAULT_ORDER, apply_overrides, render_changes,
)
from extract_format import Extractor  # noqa: E402
from verify_format import fmt_check, content_integrity  # noqa: E402

RUN_DIR = os.path.join(os.path.dirname(HERE), "run")
UPLOAD_DIR = os.path.join(RUN_DIR, "uploads")
OUT_DIR = os.path.join(RUN_DIR, "output")
for d in (UPLOAD_DIR, OUT_DIR):
    os.makedirs(d, exist_ok=True)

UI_FILE = os.path.join(HERE, "ui", "index.html")


def app_version() -> str:
    # scripts/ 的上一级：开发态=skill/，发布包=包根；再上一级=项目根。两处都可能放 VERSION
    for p in (os.path.join(HERE, "..", "..", "VERSION"),
              os.path.join(HERE, "..", "VERSION")):
        if os.path.isfile(p):
            return open(p, encoding="utf-8").read().strip()
    return "dev"


APP_VERSION = app_version()


def check_path_to_rule(path: str) -> str | None:
    """把验证检查点归属到规则（最长路径前缀匹配）。"""
    best, best_len = None, -1
    for r in RULES:
        for p in r["paths"]:
            if path == p or path.startswith(p + "."):
                if len(p) > best_len:
                    best, best_len = r["id"], len(p)
    return best


class Handler(BaseHTTPRequestHandler):
    protocol_version = "HTTP/1.1"

    def log_message(self, fmt, *args):
        pass  # 静默访问日志

    # ------------------------------------------------------------- 基础
    def _json(self, obj, code=200):
        body = json.dumps(obj, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _ndjson_begin(self):
        self.send_response(200)
        self.send_header("Content-Type", "application/x-ndjson; charset=utf-8")
        self.send_header("Cache-Control", "no-cache")
        self.close_connection = True  # 流结束即关闭，前端读到 EOF
        self.end_headers()

    def _ndjson(self, obj):
        self.wfile.write((json.dumps(obj, ensure_ascii=False) + "\n").encode("utf-8"))
        self.wfile.flush()

    def _read_json(self):
        n = int(self.headers.get("Content-Length") or 0)
        return json.loads(self.rfile.read(n) or b"{}")

    # ------------------------------------------------------------- 路由
    def do_GET(self):
        u = urllib.parse.urlparse(self.path)
        if u.path in ("/", "/index.html"):
            body = open(UI_FILE, "rb").read()
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
        elif u.path == "/api/version":
            self._json({"version": APP_VERSION, "rules": len(RULES)})
        elif u.path == "/api/rules":
            self._json({"version": APP_VERSION, "rules": RULES,
                        "default_order": DEFAULT_ORDER})
        elif u.path == "/api/file":
            q = urllib.parse.parse_qs(u.query)
            path = os.path.abspath((q.get("path") or [""])[0])
            # 仅提供 .docx 文件字节（供查看器渲染）；服务器默认只绑 127.0.0.1
            if not path.lower().endswith(".docx") or not os.path.isfile(path):
                self._json({"error": "文件不存在或不是 .docx"}, 404)
                return
            data = open(path, "rb").read()
            self.send_response(200)
            self.send_header("Content-Type",
                             "application/vnd.openxmlformats-officedocument."
                             "wordprocessingml.document")
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)
        elif u.path == "/api/listdocs":
            q = urllib.parse.parse_qs(u.query)
            d = (q.get("dir") or [UPLOAD_DIR])[0]
            try:
                files = sorted(f for f in os.listdir(d) if f.lower().endswith(".docx"))
            except OSError:
                files = []
            self._json({"dir": d, "files": files})
        else:
            self._json({"error": "not found"}, 404)

    def do_POST(self):
        u = urllib.parse.urlparse(self.path)
        if u.path == "/api/upload":
            q = urllib.parse.parse_qs(u.query)
            name = os.path.basename((q.get("name") or ["file.docx"])[0])
            n = int(self.headers.get("Content-Length") or 0)
            data = self.rfile.read(n)
            path = os.path.join(UPLOAD_DIR, name)
            with open(path, "wb") as f:
                f.write(data)
            self._json({"ok": True, "path": path, "size": len(data)})
        elif u.path == "/api/detect":
            self._detect(self._read_json())
        elif u.path == "/api/apply":
            self._apply(self._read_json())
        else:
            self._json({"error": "not found"}, 404)

    # ------------------------------------------------------------- 检测（流式）
    def _detect(self, body):
        path = body.get("path", "")
        req_path = body.get("requirement_path") or ""
        req_text = body.get("requirement_text") or ""
        if not os.path.isfile(path):
            self._json({"error": f"文件不存在：{path}"}, 400)
            return
        self._ndjson_begin()
        try:
            t0 = time.time()
            ex = Extractor(path)
            self._ndjson({"type": "start", "file": os.path.basename(path),
                          "paragraphs": sum(1 for b in ex.flow if b["kind"] == "p")})
            prof = {
                "schema_version": "1.0",
                "meta": {"source_file": os.path.basename(path),
                         "extracted_at": "", "generator": "serve_ui",
                         "doc_stats": {}},
                "flags": ex.flags,
                "dimensions_absent": [],
            }
            for step in ex.profile_steps():
                self._ndjson({"type": "step", **step})
                prof[step["key"]] = step["data"]
            ex._instruction_conflict_flags(prof)
            ex._doc_stats(prof)
            ex._absent(prof)
            # 格式要求（可选）：解析 → 合并（要求 > 模板），冲突记 flag
            req_result = None
            if req_text or (req_path and os.path.isfile(req_path)):
                from parse_requirements import (RequirementParser, merge_profile,
                                                extract_text)
                try:
                    if req_path and os.path.isfile(req_path):
                        req_text = (req_text + "\n" + extract_text(req_path))
                    levels = (prof.get("headings") or {}).get("levels_in_use")
                    parsed = RequirementParser(levels).parse(req_text)
                    req_result = merge_profile(prof, parsed)
                    prof = req_result["profile"]
                    self._ndjson({"type": "requirement",
                                  "applied": req_result["applied"],
                                  "conflicts": req_result["conflicts"],
                                  "unparsed": req_result["unparsed"]})
                except Exception as e:
                    self._ndjson({"type": "requirement", "applied": 0,
                                  "conflicts": [], "unparsed": [],
                                  "error": f"要求解析失败：{e}"})
            emap = ex.element_map()
            self._ndjson({"type": "done", "profile": prof,
                          "seconds": round(time.time() - t0, 2),
                          "map_summary": {
                              "headings": len(emap.get("headings", [])),
                              "captions": len(emap.get("captions", [])),
                              "figures": len(emap.get("figures", [])),
                              "tables": len(emap.get("tables", [])),
                              "ambiguities": emap.get("ambiguities", []),
                          }})
        except Exception as e:
            self._ndjson({"type": "error", "message": f"{type(e).__name__}: {e}"})

    # ------------------------------------------------------------- 应用（逐条流式）
    def _apply(self, body):
        target = body.get("target", "")
        profile = body.get("profile") or {}
        order = body.get("order") or DEFAULT_ORDER
        disabled = set(body.get("disabled") or [])
        overrides = body.get("overrides") or {}
        if not os.path.isfile(target):
            self._json({"error": f"文件不存在：{target}"}, 400)
            return
        if not profile:
            self._json({"error": "缺少格式档案（请先检测模板）"}, 400)
            return
        self._ndjson_begin()
        try:
            prof = apply_overrides(profile, overrides)
            applier = Applier(target, prof)
            gen = applier.run_rules(order, disabled)
            for rid in [r for r in order if True]:
                self._ndjson({"type": "rule", "id": rid, "status": "running"})
                try:
                    res = next(gen)
                except StopIteration:
                    break
                self._ndjson({"type": "rule", **res})
            # 落盘
            os.makedirs(OUT_DIR, exist_ok=True)
            base = os.path.splitext(os.path.basename(target))[0]
            out = os.path.join(OUT_DIR, f"{base}.formatted.docx")
            applier.save(out)
            rep = os.path.join(OUT_DIR, f"{base}.changes.md")
            total = render_changes(applier.changes, applier.notes, rep, False)
            self._ndjson({"type": "saved", "output": out, "report": rep,
                          "changes": total, "notes": applier.notes})
            # 验证（逐条映射回规则）
            out_ex = Extractor(out)
            checks = fmt_check(prof, out_ex.profile())
            diffs = content_integrity(target, out)
            for key, oa, ob in diffs:
                checks.append(("FAIL", f"内容完整性-{key}",
                               f"原独有 {oa} / 新独有 {ob}"))
            if not diffs:
                checks.append(("PASS", "内容完整性", "正文文字逐字一致"))
            by_rule: dict[str, dict] = {}
            for status, cpath, detail in checks:
                rid = check_path_to_rule(cpath) or "_other"
                slot = by_rule.setdefault(rid, {"pass": 0, "warn": 0, "fail": 0,
                                                "details": []})
                keyv = {"PASS": "pass", "WARN": "warn", "FAIL": "fail",
                        "SKIP": "pass"}[status]
                slot[keyv] += 1
                if status in ("FAIL", "WARN"):
                    slot["details"].append(f"{cpath}: {detail}")
            n_pass = sum(1 for s, *_ in checks if s == "PASS")
            n_fail = sum(1 for s, *_ in checks if s == "FAIL")
            self._ndjson({"type": "verify", "by_rule": by_rule,
                          "pass": n_pass, "fail": n_fail,
                          "checks": len(checks)})
            self._ndjson({"type": "done", "output": out})
        except Exception as e:
            self._ndjson({"type": "error", "message": f"{type(e).__name__}: {e}"})


def lan_addresses(port: int) -> list[str]:
    addrs = []
    try:
        for info in socket.getaddrinfo(socket.gethostname(), None, socket.AF_INET):
            ip = info[4][0]
            if ip not in ("127.0.0.1", "0.0.0.0") and not ip.startswith("169.254."):
                addrs.append(f"http://{ip}:{port}")
    except OSError:
        pass
    return sorted(set(addrs))


def main():
    args = [a for a in sys.argv[1:]]
    port = 8765
    host = "127.0.0.1"
    for i, a in enumerate(args):
        if a.isdigit():
            port = int(a)
        elif a == "--host" and i + 1 < len(args):
            host = args[i + 1]
    srv = ThreadingHTTPServer((host, port), Handler)
    print(f"docx-format-transfer 网页版 v{APP_VERSION} 已启动")
    print(f"  本机访问：http://127.0.0.1:{port}")
    if host == "0.0.0.0":
        for a in lan_addresses(port):
            print(f"  局域网访问：{a}")
        print("  （局域网设备请确保防火墙放行该端口）")
    print(f"上传目录：{UPLOAD_DIR}\n输出目录：{OUT_DIR}\nCtrl+C 退出")
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        print("\n已退出")


if __name__ == "__main__":
    main()
