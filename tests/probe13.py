# -*- coding: utf-8 -*-
"""probe13.py — bridge.meta 验证"""
import io
import json
import os
import sys

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "webapp", "app"))
import bridge  # noqa: E402

m = json.loads(bridge.meta())
print("version:", m["version"], "| rules:", len(m["rules"]))
r0 = m["rules"][0]
print("first:", r0["id"], r0["name"], r0["cat"], r0["paths"][:2], "paginated:", r0["paginated"])
