@echo off
chcp 65001 >nul
if not exist "D:\python\pythonw.exe" goto nopy
start "" "D:\python\pythonw.exe" "%~dp0MD5校验工具.py"
exit /b 0
:nopy
echo [X] pythonw.exe not found at D:\python\pythonw.exe
echo Edit this bat file and fix the Python path.
pause
