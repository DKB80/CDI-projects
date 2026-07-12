#!/usr/bin/env bash
#
# CDI Recovery Tool - Pi Zero 2W provisioning
# -------------------------------------------
# Provisions a Raspberry Pi Zero 2W (Raspberry Pi OS Lite, Bookworm) as a
# field recovery appliance for hardware YOU OWN:
#
#   * Bootable USB (mass-storage gadget) that serves a SystemRescue image,
#     for resetting a LOST LOCAL admin password on your own machines.
#   * Local hosting of the official Tailscale installers, so a managed
#     remote-support agent can be installed on-site without depending on
#     site internet.
#   * A self-hosted WiFi access point + status web panel so a field tech
#     can drive the device from a phone.
#   * Tailscale installed openly on the Pi itself for your own remote mgmt.
#
# Deliberately NOT included (add your own, out of scope for this script):
#   * HID keyboard gadget / keystroke injection.
#   * Any payload that silently configures a target machine.
# Extension points for those are documented in README.md.
#
# This script is idempotent - safe to re-run.
# Run as root:  sudo bash setup.sh

set -euo pipefail

# ---------------------------------------------------------------------------
# Config (override by exporting before running, e.g. AP_SSID=... sudo -E bash setup.sh)
# ---------------------------------------------------------------------------
CDI_ROOT="${CDI_ROOT:-/opt/cdi-recovery}"
CDI_USER="${CDI_USER:-pi}"
AP_SSID="${AP_SSID:-CDI-Recovery}"
AP_PASSPHRASE="${AP_PASSPHRASE:-ChangeMe-SetInEnv}"
AP_IP="${AP_IP:-192.168.4.1}"
AP_DHCP_START="${AP_DHCP_START:-192.168.4.10}"
AP_DHCP_END="${AP_DHCP_END:-192.168.4.50}"
WEB_PORT="${WEB_PORT:-5000}"
INSTALL_TAILSCALE="${INSTALL_TAILSCALE:-yes}"

log()  { printf '\n\033[1;36m==> %s\033[0m\n' "$*"; }
warn() { printf '\033[1;33m[warn]\033[0m %s\n' "$*"; }

if [[ $EUID -ne 0 ]]; then
  echo "This script must run as root (use: sudo bash setup.sh)" >&2
  exit 1
fi

if [[ "$AP_PASSPHRASE" == "ChangeMe-SetInEnv" ]]; then
  warn "AP_PASSPHRASE is the default placeholder. Set a real one:"
  warn "  AP_PASSPHRASE='your-strong-pass' sudo -E bash setup.sh"
fi

# Bookworm moved config.txt to /boot/firmware; keep a fallback for older layouts.
if [[ -f /boot/firmware/config.txt ]]; then
  CONFIG_TXT=/boot/firmware/config.txt
elif [[ -f /boot/config.txt ]]; then
  CONFIG_TXT=/boot/config.txt
else
  CONFIG_TXT=/boot/firmware/config.txt
  warn "config.txt not found in the usual places; will create $CONFIG_TXT"
  touch "$CONFIG_TXT"
fi

# ---------------------------------------------------------------------------
log "[1/8] Enabling USB OTG (dwc2) for mass-storage gadget mode"
# ---------------------------------------------------------------------------
# NOTE: a loose 'grep ^dtoverlay=dwc2' false-matches board-specific defaults in
# stock config.txt (e.g. the [cm5] 'dtoverlay=dwc2,dr_mode=host' line), which
# would leave a Pi Zero 2W with the USB in host mode and no gadget. Key off a
# unique marker instead, and append a fresh [all] section so the overlay applies
# to this board regardless of what precedes it. dr_mode=peripheral keeps the UDC
# available on demand (this appliance is only ever a USB device, never a host).
if ! grep -q 'CDI-Recovery dwc2 gadget' "$CONFIG_TXT"; then
  cat >> "$CONFIG_TXT" <<'CFG'

[all]
# CDI-Recovery dwc2 gadget (peripheral mode; applies to Pi Zero 2W etc.)
dtoverlay=dwc2,dr_mode=peripheral
CFG
  echo "  added dtoverlay=dwc2,dr_mode=peripheral under [all] in $CONFIG_TXT"
else
  echo "  CDI dwc2 gadget overlay already present"
fi
# Ensure modules load at boot (libcomposite is what the gadget scripts use).
grep -qxF 'dwc2'        /etc/modules || echo 'dwc2'        >> /etc/modules
grep -qxF 'libcomposite' /etc/modules || echo 'libcomposite' >> /etc/modules

# ---------------------------------------------------------------------------
log "[2/8] Installing dependencies"
# ---------------------------------------------------------------------------
export DEBIAN_FRONTEND=noninteractive
apt-get update -y
apt-get install -y \
  python3 python3-flask \
  hostapd dnsmasq \
  dosfstools parted util-linux \
  curl ca-certificates rfkill

# ---------------------------------------------------------------------------
log "[3/8] Creating CDI directory structure"
# ---------------------------------------------------------------------------
mkdir -p "$CDI_ROOT"/{installers,logs,web,payloads,gadget}
# server.py and web/index.html are expected to be copied in via scp (see README).
chown -R "$CDI_USER":"$CDI_USER" "$CDI_ROOT"
echo "  $CDI_ROOT ready"

# ---------------------------------------------------------------------------
log "[4/8] Installing Tailscale on the Pi (open, for your own remote mgmt)"
# ---------------------------------------------------------------------------
if [[ "$INSTALL_TAILSCALE" == "yes" ]]; then
  if ! command -v tailscale >/dev/null 2>&1; then
    curl -fsSL https://tailscale.com/install.sh | sh
  else
    echo "  tailscale already installed"
  fi
  echo "  NOTE: authenticate later with:  sudo tailscale up"
else
  echo "  skipped (INSTALL_TAILSCALE=$INSTALL_TAILSCALE)"
fi

# ---------------------------------------------------------------------------
log "[5/8] Configuring WiFi access point (hostapd + dnsmasq)"
# ---------------------------------------------------------------------------
cat > /etc/hostapd/hostapd.conf <<EOF
interface=wlan0
driver=nl80211
ssid=${AP_SSID}
hw_mode=g
channel=7
wmm_enabled=0
macaddr_acl=0
auth_algs=1
ignore_broadcast_ssid=0
wpa=2
wpa_passphrase=${AP_PASSPHRASE}
wpa_key_mgmt=WPA-PSK
rsn_pairwise=CCMP
EOF
chmod 600 /etc/hostapd/hostapd.conf
sed -i 's|^#\?DAEMON_CONF=.*|DAEMON_CONF="/etc/hostapd/hostapd.conf"|' /etc/default/hostapd

# Static IP for wlan0 via dhcpcd (Bookworm Lite still ships dhcpcd).
if [[ -f /etc/dhcpcd.conf ]] && ! grep -q '# CDI-Recovery AP' /etc/dhcpcd.conf; then
  cat >> /etc/dhcpcd.conf <<EOF

# CDI-Recovery AP
interface wlan0
    static ip_address=${AP_IP}/24
    nohook wpa_supplicant
EOF
fi

cat > /etc/dnsmasq.d/cdi-recovery.conf <<EOF
interface=wlan0
dhcp-range=${AP_DHCP_START},${AP_DHCP_END},255.255.255.0,24h
domain-needed
bogus-priv
EOF

# hostapd ships masked on some images.
systemctl unmask hostapd 2>/dev/null || true
rfkill unblock wifi 2>/dev/null || true
systemctl enable hostapd dnsmasq

# ---------------------------------------------------------------------------
log "[6/8] Installing mass-storage gadget script"
# ---------------------------------------------------------------------------
# The gadget script is shipped alongside this repo under gadget/.
# If run from the repo checkout, install it; otherwise expect it copied to $CDI_ROOT/gadget.
SRC_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
GADGET_SRC="$SRC_DIR/gadget/cdi-gadget-massstorage.sh"
if [[ -f "$GADGET_SRC" ]]; then
  install -m 0755 "$GADGET_SRC" /usr/local/bin/cdi-gadget-massstorage.sh
  echo "  installed /usr/local/bin/cdi-gadget-massstorage.sh"
elif [[ -f "$CDI_ROOT/gadget/cdi-gadget-massstorage.sh" ]]; then
  install -m 0755 "$CDI_ROOT/gadget/cdi-gadget-massstorage.sh" /usr/local/bin/cdi-gadget-massstorage.sh
  echo "  installed from $CDI_ROOT/gadget/"
else
  warn "cdi-gadget-massstorage.sh not found; copy it to /usr/local/bin/ manually"
fi

# ---------------------------------------------------------------------------
log "[7/8] Creating systemd service for the web panel"
# ---------------------------------------------------------------------------
cat > /etc/systemd/system/cdi-recovery.service <<EOF
[Unit]
Description=CDI Recovery status web panel
After=network.target

[Service]
Type=simple
User=${CDI_USER}
WorkingDirectory=${CDI_ROOT}
Environment=CDI_ROOT=${CDI_ROOT}
Environment=WEB_PORT=${WEB_PORT}
ExecStart=/usr/bin/python3 ${CDI_ROOT}/server.py
Restart=on-failure
RestartSec=3

[Install]
WantedBy=multi-user.target
EOF

systemctl daemon-reload
systemctl enable cdi-recovery
if [[ -f "$CDI_ROOT/server.py" ]]; then
  systemctl restart cdi-recovery || warn "cdi-recovery failed to start; check: journalctl -u cdi-recovery"
else
  warn "server.py not yet in $CDI_ROOT - copy it over, then: sudo systemctl restart cdi-recovery"
fi

# ---------------------------------------------------------------------------
log "[8/8] Creating placeholder recovery disk image (if absent)"
# ---------------------------------------------------------------------------
if [[ ! -f "$CDI_ROOT/recovery.img" ]]; then
  # 8 MiB placeholder so the gadget has something to expose until you build
  # the real SystemRescue image with build-recovery-image.sh.
  dd if=/dev/zero of="$CDI_ROOT/recovery.img" bs=1M count=8 status=none
  mkfs.vfat -n CDIRECOVERY "$CDI_ROOT/recovery.img" >/dev/null
  chown "$CDI_USER":"$CDI_USER" "$CDI_ROOT/recovery.img"
  echo "  placeholder recovery.img created (build the real one with build-recovery-image.sh)"
else
  echo "  recovery.img already present, leaving it"
fi

log "Setup complete"
cat <<EOF

Next steps:
  1. Copy server.py and web/index.html into ${CDI_ROOT} if not already:
       scp server.py           ${CDI_USER}@<pi>:${CDI_ROOT}/server.py
       scp web/index.html      ${CDI_USER}@<pi>:${CDI_ROOT}/web/index.html
  2. Build the bootable recovery image:  sudo bash build-recovery-image.sh
  3. Fetch Tailscale installers to host:  bash fetch-installers.sh
  4. Reboot:  sudo reboot
  5. Join WiFi "${AP_SSID}" and open http://${AP_IP}:${WEB_PORT}

To present the recovery USB to a target machine you own:
     sudo /usr/local/bin/cdi-gadget-massstorage.sh on
EOF
