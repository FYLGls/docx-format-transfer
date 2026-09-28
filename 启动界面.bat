@echo off
rem 双击启动 docx-format-transfer 可视化界面
cd /d "%~dp0skill\scripts"
start "docx-format-transfer" /min cmd /c "timeout /t 2 >nul & start http://127.0.0.1:8765"
python serve_ui.py
