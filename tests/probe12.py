# -*- coding: utf-8 -*-
"""probe12.py — 端口清理 + 启动 serve_ui + 验证 /api/file"""
import io
import os
import subprocess
import sys
import time
import urllib.parse
import urllib.request

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SCRIPTS = os.path.join(ROOT, "skill", "scripts")
PORT = 8933


def kill_port(port):
    r = subprocess.run(["netstat", "-ano"], capture_output=True,
                       text=True, encoding="utf-8", errors="replace")
    if not r.stdout:
        return
    for line in r.stdout.splitlines():
        if f":{port}" in line and "LISTENING" in line:
            pid = line.split()[-1]
            subprocess.run(["taskkill", "/F", "/PID", pid],
                           capture_output=True)


kill_port(PORT)
srv = subprocess.Popen([sys.executable, os.path.join(SCRIPTS, "serve_ui.py"), str(PORT)],
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
for _ in range(40):
    time.sleep(0.25)
    try:
        urllib.request.urlopen(f"http://127.0.0.1:{PORT}/api/version", timeout=2).read()
        break
    except Exception:
        pass
target = os.path.join(ROOT, "tests", "fixtures", "thesis_standard", "target.docx")
q = urllib.parse.quote(target, safe="")
r = urllib.request.urlopen(f"http://127.0.0.1:{PORT}/api/file?path={q}", timeout=10)
data = r.read()
print("api/file:", r.status, len(data), "bytes, PK\\x03\\x04 =",
      data[:2] == b"PK\x03\x04")
assert data[:2] == b"PK\x03\x04"
print("OK — 保持服务器运行（PID 由进程组管理）")
