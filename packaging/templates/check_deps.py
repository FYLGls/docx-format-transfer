# -*- coding: utf-8 -*-
"""check_deps.py — 依赖体检。缺什么打印什么，退出码 0=就绪 1=缺依赖。
pdf/图片格式要求解析为可选能力，缺失仅提示不影响核心功能。"""
import importlib
import sys

ok = True
if sys.version_info < (3, 10):
    print(f"需要 Python 3.10+，当前 {sys.version.split()[0]}")
    ok = False
for mod, pip in (("docx", "python-docx"), ("lxml", "lxml")):
    try:
        importlib.import_module(mod)
    except ImportError:
        print(f"缺少 {pip}：请执行  pip install {pip}")
        ok = False
# 可选：格式要求文件解析（docx 模板功能不需要）
opt_missing = []
try:
    importlib.import_module("pypdf")
except ImportError:
    opt_missing.append("pypdf")
try:
    importlib.import_module("rapidocr_onnxruntime")
except ImportError:
    opt_missing.append("rapidocr-onnxruntime")
if opt_missing:
    print(f"可选（解析 pdf/图片格式的格式要求时需要）：pip install {' '.join(opt_missing)}")
if ok:
    print("依赖就绪 ✓")
sys.exit(0 if ok else 1)
