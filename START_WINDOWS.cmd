@echo off
chcp 65001 >nul
cd /d "%~dp0"
if exist ".venv\Scripts\python.exe" goto RUN
where py >nul 2>nul
if errorlevel 1 (
  echo Python 3.12 64-bit 설치가 필요합니다.
  echo python.org 공식 설치 프로그램을 사용한 뒤 다시 실행해 주세요.
  pause
  exit /b 1
)
py -3 -c "import sys; sys.exit(0 if sys.version_info >= (3,11) else 1)"
if errorlevel 1 (
  echo Python 3.11 이상이 필요합니다. Python 3.12를 권장합니다.
  pause
  exit /b 1
)
py -3 -m venv .venv
if errorlevel 1 goto ERROR
".venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 goto ERROR
:RUN
".venv\Scripts\python.exe" -c "import lxml; import tkinter; import win32com.client"
if errorlevel 1 (
  ".venv\Scripts\python.exe" -m pip install -r requirements.txt
  if errorlevel 1 goto ERROR
)
".venv\Scripts\python.exe" run.py
if errorlevel 1 goto ERROR
exit /b 0
:ERROR
echo 실행 또는 설치에 실패했습니다. 위 오류와 사용 안내를 확인해 주세요.
echo 첫 설치에는 패키지 다운로드를 위한 인터넷 연결이 필요합니다.
pause
exit /b 1
