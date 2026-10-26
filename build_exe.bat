@echo off
chcp 65001 >nul
cd /d "%~dp0"
py -m PyInstaller --noconfirm --clean --onefile --windowed --name "관로번호_검사프로그램" --collect-all openpyxl "관로번호_검사프로그램.py"
if errorlevel 1 (
  echo.
  echo EXE 제작에 실패했습니다. install_requirements.bat를 먼저 실행했는지 확인해 주세요.
  pause
  exit /b 1
)
echo.
echo EXE 제작 완료: dist\관로번호_검사프로그램.exe
pause
