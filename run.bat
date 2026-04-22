@echo off
REM ============================================================
REM  CDI Outlook Briefing Tool — update + launch
REM
REM  Double-click this file to:
REM    1. Pull the latest code from GitHub
REM    2. Update Python packages if needed
REM    3. Launch the GUI
REM
REM  Use this when you want to make sure you're running the latest
REM  version (e.g. after I push a fix). For day-to-day use, double-
REM  click  launch.bat  instead — it skips the update step and is
REM  much faster.
REM ============================================================

REM Move into this batch file's own folder so relative paths work
REM regardless of where the shortcut is launched from.
cd /d "%~dp0"

echo === Checking for updates ===
git pull
if errorlevel 1 (
    echo.
    echo WARNING: git pull failed. Continuing with the current local copy.
    echo.
)

echo.
echo === Updating Python packages ===
pip install -r requirements.txt
if errorlevel 1 (
    echo.
    echo ERROR: pip install failed. The GUI may not launch. Press any key to try anyway.
    pause >nul
)

echo.
echo === Launching CDI Outlook Briefing Tool ===
python cdi_outlook_gui.py

REM Only pause on error so a clean exit closes the window automatically.
if errorlevel 1 (
    echo.
    echo The tool exited with an error. Scroll up to see what went wrong.
    echo Press any key to close this window.
    pause >nul
)
