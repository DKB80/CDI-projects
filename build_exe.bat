@echo off
REM ============================================================
REM  Build a standalone .exe of the CDI Outlook Briefing Tool.
REM
REM  Output: dist\CDI Outlook Briefing Tool\
REM          -- this whole folder is what you zip and send to
REM             co-workers. They extract it and double-click
REM             "CDI Outlook Briefing Tool.exe".
REM
REM  One-time setup (on the machine that builds the .exe):
REM     pip install -r requirements.txt
REM
REM  Then just double-click this file whenever you want to
REM  package a new version.
REM ============================================================

cd /d "%~dp0"

echo === Making sure PyInstaller is installed ===
python -m pip install --upgrade pyinstaller >nul
if errorlevel 1 (
    echo.
    echo ERROR: couldn't install PyInstaller. Check your Python setup.
    pause
    exit /b 1
)

echo.
echo === Cleaning previous build ===
if exist build rmdir /s /q build
if exist "dist\CDI Outlook Briefing Tool" rmdir /s /q "dist\CDI Outlook Briefing Tool"

echo.
echo === Building .exe (this takes 2-5 minutes the first time) ===
python -m PyInstaller cdi_outlook_gui.spec
if errorlevel 1 (
    echo.
    echo Build failed. Scroll up for the PyInstaller error.
    pause
    exit /b 1
)

echo.
echo === Build complete ===
echo.
echo Output folder: dist\CDI Outlook Briefing Tool\
echo To distribute: right-click that folder -^> Send to -^> Compressed (zipped) folder,
echo then send the .zip to whoever needs it.
echo.
pause
