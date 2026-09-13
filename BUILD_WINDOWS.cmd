@echo off
chcp 65001 >nul
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" py -3 -m venv .venv
if errorlevel 1 goto ERROR
".venv\Scripts\python.exe" -m pip install -r requirements-dev.txt
if errorlevel 1 goto ERROR
".venv\Scripts\python.exe" build_windows.py
if errorlevel 1 goto ERROR
echo 완료: dist\RENAME-Windows-x64.zip
pause
exit /b 0
:ERROR
echo 빌드가 실패했습니다. Python 3.12 설치와 인터넷 연결을 확인해 주세요.
pause
exit /b 1
