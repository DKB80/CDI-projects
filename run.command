#!/usr/bin/env bash
# ============================================================
#  CDI Outlook Briefing Tool — update + launch (macOS)
#
#  Double-click this file to:
#    1. Pull the latest code from GitHub
#    2. Install / update Python packages
#    3. Launch the GUI
#
#  On first launch macOS will ask to grant "Automation" permission
#  for Python to drive Outlook — click Allow.
#
#  If double-click doesn't work, right-click this file in Finder,
#  choose "Open With → Terminal", and confirm the security prompt.
#
#  For day-to-day use after first run, use  launch.command  — it
#  skips git/pip and opens faster.
# ============================================================

set -e
cd "$(dirname "$0")"

echo "=== Checking for updates ==="
if git pull; then
    true
else
    echo "WARNING: git pull failed — continuing with the current local copy."
fi

echo
echo "=== Updating Python packages ==="
PY=$(command -v python3 || command -v python)
if [ -z "$PY" ]; then
    echo "ERROR: Python 3 not found. Install it with 'brew install python' or from python.org."
    read -p "Press Enter to close..."
    exit 1
fi
$PY -m pip install --user -r requirements.txt || {
    echo "ERROR: pip install failed. See the output above."
    read -p "Press Enter to close..."
    exit 1
}

echo
echo "=== Launching CDI Outlook Briefing Tool ==="
$PY cdi_outlook_gui.py || {
    echo
    echo "The tool exited with an error. Scroll up to see why."
    read -p "Press Enter to close..."
}
