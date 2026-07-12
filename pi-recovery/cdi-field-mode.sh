#!/usr/bin/env bash
#
# cdi-field-mode.sh - toggle wlan0 between WiFi CLIENT and self-hosted AP.
# ----------------------------------------------------------------------
# NetworkManager-native (works on Debian 13 / Trixie). Replaces setup.sh's
# hostapd+dhcpcd AP block, which is dhcpcd/Bookworm-specific and breaks on
# Trixie. Uses `ipv4.method shared`, so NM runs its own DHCP (and NAT if an
# uplink exists) - no hostapd/dnsmasq required.
#
# The Pi Zero 2W has ONE radio: AP mode drops the client WiFi (and therefore
# SSH/Tailscale over WiFi). So `on` schedules an AUTO-REVERT timer; if you do
# nothing, it returns to client WiFi after N minutes. A reboot also returns to
# client mode (the AP profile is autoconnect=no). You cannot strand the Pi.
#
#   sudo cdi-field-mode.sh on [minutes]  # become AP 'CDI-Recovery' @ 192.168.4.1
#                                        #   (auto-reverts after N min, default 10)
#   sudo cdi-field-mode.sh keep          # cancel auto-revert, stay in AP mode
#   sudo cdi-field-mode.sh off           # back to client WiFi now
#   sudo cdi-field-mode.sh prepare       # create/refresh the AP profile, don't activate
#   sudo cdi-field-mode.sh status
#
set -euo pipefail

AP_CON="cdi-ap"
AP_SSID="${AP_SSID:-CDI-Recovery}"
AP_IP="${AP_IP:-192.168.4.1/24}"
AP_CHANNEL="${AP_CHANNEL:-7}"
PSK_FILE="${PSK_FILE:-/etc/cdi-recovery/ap.psk}"
SELF="/usr/local/bin/cdi-field-mode.sh"

[[ $EUID -eq 0 ]] || { echo "run as root" >&2; exit 1; }

ensure_profile() {
  [[ -f "$PSK_FILE" ]] || { echo "missing $PSK_FILE (AP passphrase, 8-63 chars)" >&2; exit 1; }
  local psk; psk="$(head -n1 "$PSK_FILE")"
  if nmcli -t -f NAME connection show | grep -qx "$AP_CON"; then
    nmcli connection modify "$AP_CON" wifi-sec.psk "$psk" >/dev/null
  else
    nmcli connection add type wifi ifname wlan0 con-name "$AP_CON" \
      autoconnect no ssid "$AP_SSID" >/dev/null
    nmcli connection modify "$AP_CON" \
      802-11-wireless.mode ap 802-11-wireless.band bg 802-11-wireless.channel "$AP_CHANNEL" \
      ipv4.method shared ipv4.addresses "$AP_IP" \
      wifi-sec.key-mgmt wpa-psk wifi-sec.psk "$psk" >/dev/null
  fi
}

# pick a non-AP wifi profile to return to (highest priority autoconnect wins anyway)
first_client() {
  nmcli -t -f NAME,TYPE connection show \
    | awk -F: -v ap="$AP_CON" '$2=="802-11-wireless" && $1!=ap {print $1; exit}'
}

client_up() {
  nmcli connection down "$AP_CON" >/dev/null 2>&1 || true
  local c; c="$(first_client)"
  if [[ -n "$c" ]]; then
    nmcli connection up "$c" >/dev/null 2>&1 || nmcli device connect wlan0 >/dev/null 2>&1 || true
  else
    nmcli device connect wlan0 >/dev/null 2>&1 || true
  fi
}

case "${1:-status}" in
  prepare)
    ensure_profile
    echo "AP profile '$AP_CON' ready (SSID '$AP_SSID', ${AP_IP%/*}). Not activated."
    ;;
  on)
    mins="${2:-10}"
    ensure_profile
    systemctl stop cdi-field-revert.timer 2>/dev/null || true
    systemctl reset-failed 'cdi-field-*' 2>/dev/null || true
    # auto-revert safety net
    systemd-run --on-active="${mins}min" --unit=cdi-field-revert --collect \
      "$SELF" off >/dev/null
    # detach the switch so the calling SSH session returns before WiFi drops
    systemd-run --on-active=2 --unit=cdi-field-up --collect \
      nmcli connection up "$AP_CON" >/dev/null
    echo "Switching to AP '$AP_SSID' @ ${AP_IP%/*} in ~2s - WiFi SSH/Tailscale WILL drop."
    echo "Connect a device to '$AP_SSID' and open http://${AP_IP%/*}:5000"
    echo "Auto-reverts to client WiFi in ${mins} min. To stay: sudo cdi-field-mode.sh keep"
    ;;
  keep)
    systemctl stop cdi-field-revert.timer 2>/dev/null || true
    systemctl reset-failed cdi-field-revert.service 2>/dev/null || true
    echo "Auto-revert cancelled - staying in AP mode until 'off' or reboot."
    ;;
  off)
    systemctl stop cdi-field-revert.timer 2>/dev/null || true
    client_up
    echo "Reverted to client WiFi."
    ;;
  status)
    echo "wlan0: $(nmcli -t -f DEVICE,STATE,CONNECTION device status | grep '^wlan0' || echo '?')"
    if nmcli -t -f NAME,STATE connection show --active | grep -q "^${AP_CON}:activated"; then
      echo "mode:  AP ($AP_SSID @ ${AP_IP%/*}:5000)"
      systemctl is-active cdi-field-revert.timer >/dev/null 2>&1 \
        && echo "auto-revert: ARMED" || echo "auto-revert: off (staying)"
    else
      echo "mode:  client"
    fi
    ;;
  *) echo "usage: $0 {on [minutes]|keep|off|prepare|status}" >&2; exit 1 ;;
esac
