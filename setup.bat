@echo off
REM ============================================================
REM  AS-BUILT STAMPER — Setup Script
REM  Run this once on your Windows machine before first use.
REM ============================================================

echo.
echo  AS-BUILT STAMPER — Dependency Installer
echo  Hargreaves / Exentec Hargreaves
echo ============================================================
echo.

REM Check Python is available
python --version >nul 2>&1
IF ERRORLEVEL 1 (
    echo  ERROR: Python not found.
    echo  Please install Python 3.9+ from https://python.org
    echo  Make sure to tick "Add Python to PATH" during install.
    pause
    exit /b 1
)

echo  [OK] Python found:
python --version

echo.
echo  Installing required packages...
echo.

pip install --upgrade pip
pip install pywin32

echo.
echo  Running pywin32 post-install (requires admin rights)...
python -c "import site; import os; scripts = [p for p in site.getsitepackages() if 'Scripts' in p or 'scripts' in p]; print(scripts)"

REM Try to run the post-install script
FOR /F "tokens=*" %%G IN ('python -c "import sys; print(sys.prefix)"') DO SET PYPREFIX=%%G

IF EXIST "%PYPREFIX%\Scripts\pywin32_postinstall.py" (
    python "%PYPREFIX%\Scripts\pywin32_postinstall.py" -install
    echo  [OK] pywin32 post-install complete
) ELSE (
    echo  [WARN] Could not find pywin32_postinstall.py
    echo  You may need to run it manually:
    echo    python Scripts\pywin32_postinstall.py -install
)

echo.
echo ============================================================
echo  Setup complete!
echo.
echo  To run the application:
echo    python asbuilt_stamper.py
echo.
echo  IMPORTANT: AutoCAD must be installed and licensed.
echo  The app connects to AutoCAD via COM automation.
echo ============================================================
echo.
pause
