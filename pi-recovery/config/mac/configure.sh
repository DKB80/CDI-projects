#!/bin/bash
# CDI standard configuration for a Mac mini (standalone path). Idempotent; logs
# a report. Run with admin rights:  sudo bash configure.sh
# Not keystroke injection - a transparent script run deliberately on a Mac you service.
set -uo pipefail

TZ_WANT="Australia/Perth"
REPORT="${PWD}/cdi-config_$(scutil --get ComputerName 2>/dev/null || hostname)_$(date +%Y%m%d-%H%M%S).report.txt"
note() { printf '%-16s %-6s %s\n' "$1" "$2" "$3" | tee -a "$REPORT"; }

[ "$(id -u)" -eq 0 ] || { echo "run with sudo"; exit 1; }
echo "CDI Mac config - $(date)" | tee "$REPORT"

# OS - timezone + NTP
if [ "$(systemsetup -gettimezone | awk '{print $NF}')" != "$TZ_WANT" ]; then
  systemsetup -settimezone "$TZ_WANT" >/dev/null && note OS SET "timezone -> $TZ_WANT"
else note OS OK "timezone already $TZ_WANT"; fi
systemsetup -setusingnetworktime on >/dev/null 2>&1 && note OS SET "network time on"

# OS - never sleep (server/kiosk)
pmset -a sleep 0 disksleep 0 >/dev/null 2>&1 && note OS SET "never sleep"

# OS - enable Remote Login (SSH) + Remote Management done via MDM/profile normally
systemsetup -setremotelogin on >/dev/null 2>&1 && note OS SET "remote login (ssh) on"

# Software - Tailscale via Homebrew if present, else point to hosted installer
if command -v brew >/dev/null 2>&1; then
  sudo -u "$(stat -f%Su /dev/console)" brew list --cask tailscale >/dev/null 2>&1 \
    && note Software OK "tailscale present" \
    || { sudo -u "$(stat -f%Su /dev/console)" brew install --cask tailscale >/dev/null 2>&1 && note Software SET "tailscale installed"; }
else
  note Software WARN "no brew - install Tailscale from the Pi: http://cdi-recovery.local:5000/installers/tailscale-macos.pkg"
fi

echo "Report: $REPORT"
