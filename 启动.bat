@echo off
chcp 65001 >nul
cd /d "%~dp0"
where pythonw >nul 2>nul
if errorlevel 1 (
  where python >nul 2>nul
  if errorlevel 1 (
    echo [!] 没有检测到 Python，请直接双击 dist\批量打印.exe
    pause
    exit /b 1
  )
  python gui.py
  exit /b 0
)
start "" pythonw "%~dp0gui.py"
