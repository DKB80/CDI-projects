#!/usr/bin/env bash
# ============================================================
#  CDI Outlook Briefing Tool — launch only (macOS)
#
#  Double-click for day-to-day use. Skips git/pip steps.
#  Use  run.command  instead when you want the latest version.
# ============================================================

cd "$(dirname "$0")"
PY=$(command -v python3 || command -v python)
$PY cdi_outlook_gui.py || {
    echo
    echo "The tool exited with an error. Scroll up to see why."
    read -p "Press Enter to close..."
}
