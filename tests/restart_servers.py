# -*- coding: utf-8 -*-
"""restart_servers.py — 清理端口并以最新代码重启 serve_ui/webapp 两个测试服务器（驻留）"""
import io
import os
import subprocess
import sys
import time
import urllib.request

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SCRIPTS = os.path.join(ROOT, "skill", "scripts")
WEBAPP = os.path.join(ROOT, "webapp")


def kill_port(port):
    r = subprocess.run(["netstat", "-ano"], capture_output=True,
                       text=True, encoding="utf-8", errors="replace")
    if not r.stdout:
        return
    for line in r.stdout.splitlines():
        if f":{port}" in line and "LISTENING" in line:
            pid = line.split()[-1]
            subprocess.run(["taskkill", "/F", "/PID", pid], capture_output=True)


DETACHED = 0x00000008  # DETACHED_PROCESS
for port, cwd, cmd in ((8933, SCRIPTS, [sys.executable, "serve_ui.py", "8933"]),
                       (8934, WEBAPP, [sys.executable, "-m", "http.server", "8934"])):
    kill_port(port)
    subprocess.Popen(cmd, cwd=cwd, creationflags=DETACHED,
                     stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
for port in (8933, 8934):
    for _ in range(40):
        time.sleep(0.25)
        try:
            urllib.request.urlopen(f"http://127.0.0.1:{port}/", timeout=2).read()
            print(f"port {port}: UP")
            break
        except Exception:
            pass
    else:
        print(f"port {port}: DOWN")
