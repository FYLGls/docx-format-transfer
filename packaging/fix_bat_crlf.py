# -*- coding: utf-8 -*-
"""fix_bat_crlf.py — 把项目根目录 启动界面.bat 转为 GBK+CRLF（cmd 批处理要求）"""
import os

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
p = os.path.join(ROOT, "启动界面.bat")
text = open(p, encoding="utf-8").read()
text = text.replace("\r\n", "\n").replace("\n", "\r\n")
open(p, "wb").write(text.encode("gbk"))
print("root bat -> GBK+CRLF,", len(text), "chars")
