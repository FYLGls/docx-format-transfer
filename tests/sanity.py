# -*- coding: utf-8 -*-
"""sanity.py — 输出文件可打开性 + 幂等性检查"""
import io
import os
import subprocess
import sys

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
from docx import Document  # noqa: E402

HERE = os.path.dirname(os.path.abspath(__file__))
FIXDIR = os.path.join(HERE, "fixtures")
SCRIPTS = os.path.join(HERE, "..", "skill", "scripts")
bad = 0
for name in sorted(os.listdir(FIXDIR)):
    d = os.path.join(FIXDIR, name)
    f = os.path.join(d, "target.formatted.docx")
    if not os.path.isfile(f):
        continue
    try:
        doc = Document(f)  # 能否作为合法 docx 打开
        n_par = len(doc.paragraphs)
        print(f"open-ok {name} ({n_par} paras)")
    except Exception as e:
        bad += 1
        print(f"open-FAIL {name}: {e}")
# 幂等性：对 thesis 再跑一次 apply（输出到临时文件）
d = os.path.join(FIXDIR, "thesis_standard")
r = subprocess.run([sys.executable, os.path.join(SCRIPTS, "apply_format.py"),
                    os.path.join(d, "target.docx"),
                    os.path.join(d, "template.profile.json"),
                    "-o", os.path.join(d, "_idem.docx")],
                   capture_output=True, text=True, encoding="utf-8")
print("idempotent rerun:", "OK" if r.returncode == 0 else f"FAIL {r.stderr[-200:]}")
r2 = subprocess.run([sys.executable, os.path.join(SCRIPTS, "apply_format.py"),
                     os.path.join(d, "_idem.docx"),
                     os.path.join(d, "template.profile.json"),
                     "-o", os.path.join(d, "_idem2.docx")],
                    capture_output=True, text=True, encoding="utf-8")
print("second pass:", "OK" if r2.returncode == 0 else f"FAIL {r2.stderr[-200:]}")
sys.exit(1 if bad else 0)
