@echo off
chcp 65001 >nul
rem 优先启动打包好的单文件 exe（免 Python 环境）
if exist "%~dp0dist\MD5校验工具.exe" (
  start "" "%~dp0dist\MD5校验工具.exe"
  exit /b 0
)
if exist "%~dp0MD5校验工具.exe" (
  start "" "%~dp0MD5校验工具.exe"
  exit /b 0
)
rem 没有 exe 时回退到源码运行
where pythonw >nul 2>nul
if %errorlevel%==0 (
  start "" pythonw "%~dp0MD5校验工具.py"
  exit /b 0
)
echo [X] 未找到 MD5校验工具.exe，也未安装 Python。
echo 请直接运行 dist\MD5校验工具.exe
pause
