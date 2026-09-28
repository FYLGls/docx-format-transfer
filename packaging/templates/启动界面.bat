@echo off
rem 双击启动 docx-format-transfer 网页版（自动打开浏览器）
cd /d "%~dp0scripts"
python check_deps.py || (echo 按上方提示安装依赖后重试 & pause & exit /b 1)
start "" http://127.0.0.1:8765
python serve_ui.py
pause
