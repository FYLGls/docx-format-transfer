# -*- coding: utf-8 -*-
"""build_release.py — 构建 docx-format-transfer 可下载发布包（zip + sha256）

用法：python packaging/build_release.py [--out release]
产物：release/docx-format-transfer-v<VERSION>.zip[.sha256]
"""
from __future__ import annotations

import argparse
import hashlib
import io
import os
import sys
import zipfile

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SKILL = os.path.join(ROOT, "skill")
TPL = os.path.join(ROOT, "packaging", "templates")

EXCLUDE_DIRS = {"__pycache__", "run"}
EXCLUDE_EXT = {".pyc"}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=os.path.join(ROOT, "release"))
    args = ap.parse_args()
    version = open(os.path.join(ROOT, "VERSION"), encoding="utf-8").read().strip()
    top = f"docx-format-transfer-v{version}"
    out_dir = args.out
    os.makedirs(out_dir, exist_ok=True)
    zip_path = os.path.join(out_dir, f"{top}.zip")

    entries: list[tuple[str, str]] = []  # (arcname, src_path)

    def add_file(arc, src):
        entries.append((arc, src))

    # 根：模板文件 + VERSION
    for name in ("启动界面.bat", "install-skill.bat", "check_deps.py", "README.md"):
        add_file(f"{top}/{name}", os.path.join(TPL, name))
    add_file(f"{top}/VERSION", os.path.join(ROOT, "VERSION"))
    # skill 本体：SKILL.md 在根（整目录可直装为 ZCode skill）
    add_file(f"{top}/SKILL.md", os.path.join(SKILL, "SKILL.md"))
    for sub in ("references", "assets"):
        d = os.path.join(SKILL, sub)
        for base, _, files in os.walk(d):
            if "__pycache__" in base:
                continue
            for f in sorted(files):
                if os.path.splitext(f)[1] in EXCLUDE_EXT:
                    continue
                rel = os.path.relpath(os.path.join(base, f), SKILL)
                add_file(f"{top}/{rel.replace(os.sep, '/')}",
                         os.path.join(base, f))
    sd = os.path.join(SKILL, "scripts")
    for base, dirs, files in os.walk(sd):
        dirs[:] = [d for d in dirs if d not in EXCLUDE_DIRS]
        for f in sorted(files):
            if os.path.splitext(f)[1] in EXCLUDE_EXT:
                continue
            full = os.path.join(base, f)
            rel = os.path.relpath(full, SKILL)
            add_file(f"{top}/{rel.replace(os.sep, '/')}", full)

    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as z:
        for arc, src in entries:
            if arc.endswith(".bat"):
                # cmd 批处理要求 ANSI(GBK) 编码 + CRLF 换行；模板是 UTF-8+LF，转码后入包
                text = open(src, encoding="utf-8").read()
                text = text.replace("\r\n", "\n").replace("\n", "\r\n")
                z.writestr(arc, text.encode("gbk"))
            else:
                z.write(src, arc)
    sha = hashlib.sha256(open(zip_path, "rb").read()).hexdigest()
    with open(zip_path + ".sha256", "w", encoding="utf-8") as f:
        f.write(f"{sha}  {os.path.basename(zip_path)}\n")
    n = len(entries)
    size = os.path.getsize(zip_path)
    print(f"构建完成：{zip_path}")
    print(f"  条目 {n}，大小 {size / 1024:.0f} KB，sha256 {sha[:16]}…")
    print(f"  版本 v{version}")


if __name__ == "__main__":
    main()
