@echo off
rem 一键注册为 ZCode skill。可选参数：目标目录（默认 %USERPROFILE%\.agents\skills\docx-format-transfer）
set "DEST=%USERPROFILE%\.agents\skills\docx-format-transfer"
if not "%~1"=="" set "DEST=%~1"
echo 安装 docx-format-transfer 到 %DEST% ...
robocopy "%~dp0." "%DEST%" /E /XF install-skill.bat 启动界面.bat >nul
if errorlevel 8 (
  echo 安装失败。
  exit /b 1
)
echo 完成。已注册为 ZCode skill（重启 ZCode 后生效，说"按模板改格式"即可触发）。
exit /b 0
