@echo off
chcp 65001 >nul
cd /d "%~dp0"
py -m pip install --upgrade pip
py -m pip install openpyxl pyinstaller
if errorlevel 1 (
  echo.
  echo 설치에 실패했습니다. Python 설치 여부와 인터넷 연결을 확인해 주세요.
  pause
  exit /b 1
)
echo.
echo 필요한 패키지 설치가 완료되었습니다.
pause
