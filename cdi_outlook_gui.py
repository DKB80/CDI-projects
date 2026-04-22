#!/usr/bin/env python3
"""Entry point for the CDI Outlook Briefing Tool GUI.

Double-click this file to run, or from a terminal:
    python cdi_outlook_gui.py

PyInstaller (one-folder distributable .exe):
    pyinstaller cdi_outlook_gui.spec
"""

from cdi_gui.main_window import main

if __name__ == "__main__":
    main()
