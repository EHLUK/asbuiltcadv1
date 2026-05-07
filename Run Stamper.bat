@echo off
REM ============================================================
REM  AS-BUILT STAMPER — Launcher
REM ============================================================
cd /d "%~dp0"
python asbuilt_stamper.py
IF ERRORLEVEL 1 (
    echo.
    echo  An error occurred. Have you run setup.bat?
    pause
)
