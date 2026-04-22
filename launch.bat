@echo off
REM ============================================================
REM  CDI Outlook Briefing Tool — launch only (no update)
REM
REM  Double-click this file for day-to-day use. It skips the
REM  git pull / pip install steps, so it starts fast.
REM
REM  Use  run.bat  instead when you want to grab the latest
REM  version of the tool.
REM ============================================================

cd /d "%~dp0"
python cdi_outlook_gui.py

if errorlevel 1 (
    echo.
    echo The tool exited with an error. Scroll up to see what went wrong.
    echo Press any key to close this window.
    pause >nul
)
