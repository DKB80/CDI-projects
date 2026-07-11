#!/usr/bin/env bash
#
# cdi-gadget-massstorage.sh - present recovery.img to a host as a USB drive
# -------------------------------------------------------------------------
# Turns the Pi Zero's OTG port into a USB Mass Storage device backed by
# ${CDI_ROOT}/recovery.img. Plug the Pi's OTG (data) port into a machine
# YOU OWN, then boot that machine from the Pi to run SystemRescue / chntpw
# and reset a lost LOCAL admin password.
#
# This is a single-function gadget: mass storage only. It does NOT present
# an HID keyboard and cannot type into the host. Keystroke-injection is
# intentionally out of scope - see README.md.
#
# Usage:
#   sudo cdi-gadget-massstorage.sh on      # attach the USB drive
#   sudo cdi-gadget-massstorage.sh off     # detach
#   sudo cdi-gadget-massstorage.sh status  # show current state
#
set -euo pipefail

CDI_ROOT="${CDI_ROOT:-/opt/cdi-recovery}"
IMG="${IMG:-$CDI_ROOT/recovery.img}"
G="/sys/kernel/config/usb_gadget/cdi"

require_root() {
  [[ $EUID -eq 0 ]] || { echo "run as root" >&2; exit 1; }
}

ensure_modules() {
  modprobe libcomposite 2>/dev/null || true
  if [[ ! -d /sys/class/udc ]] || [[ -z "$(ls -A /sys/class/udc 2>/dev/null)" ]]; then
    modprobe dwc2 2>/dev/null || true
  fi
}

gadget_on() {
  [[ -f "$IMG" ]] || { echo "image not found: $IMG (build it with build-recovery-image.sh)" >&2; exit 1; }
  ensure_modules

  local udc
  udc="$(ls /sys/class/udc 2>/dev/null | head -n1 || true)"
  [[ -n "$udc" ]] || { echo "no UDC available - is dwc2 loaded and dtoverlay=dwc2 set?" >&2; exit 1; }

  if [[ -d "$G" ]]; then
    echo "gadget already configured; reattaching"
    echo "$udc" > "$G/UDC" 2>/dev/null || true
    echo "on"
    return
  fi

  mkdir -p "$G"
  echo 0x1d6b > "$G/idVendor"   # Linux Foundation
  echo 0x0104 > "$G/idProduct"  # Multifunction Composite Gadget
  echo 0x0100 > "$G/bcdDevice"
  echo 0x0200 > "$G/bcdUSB"

  mkdir -p "$G/strings/0x409"
  echo "CDI0001"       > "$G/strings/0x409/serialnumber"
  echo "CDI"           > "$G/strings/0x409/manufacturer"
  echo "CDI Recovery"  > "$G/strings/0x409/product"

  mkdir -p "$G/configs/c.1/strings/0x409"
  echo "Mass Storage" > "$G/configs/c.1/strings/0x409/configuration"
  echo 250 > "$G/configs/c.1/MaxPower"

  mkdir -p "$G/functions/mass_storage.0"
  echo 1     > "$G/functions/mass_storage.0/stall"
  echo 0     > "$G/functions/mass_storage.0/lun.0/cdrom"
  echo 0     > "$G/functions/mass_storage.0/lun.0/ro"
  echo 0     > "$G/functions/mass_storage.0/lun.0/nofua"
  echo "$IMG" > "$G/functions/mass_storage.0/lun.0/file"

  ln -s "$G/functions/mass_storage.0" "$G/configs/c.1/" 2>/dev/null || true

  echo "$udc" > "$G/UDC"
  echo "on (backing image: $IMG)"
}

gadget_off() {
  if [[ -d "$G" ]]; then
    echo "" > "$G/UDC" 2>/dev/null || true
    # Tear down links and dirs in reverse order.
    rm -f "$G/configs/c.1/mass_storage.0" 2>/dev/null || true
    rmdir "$G/configs/c.1/strings/0x409" 2>/dev/null || true
    rmdir "$G/configs/c.1" 2>/dev/null || true
    rmdir "$G/functions/mass_storage.0" 2>/dev/null || true
    rmdir "$G/strings/0x409" 2>/dev/null || true
    rmdir "$G" 2>/dev/null || true
    echo "off"
  else
    echo "off (nothing configured)"
  fi
}

gadget_status() {
  if [[ -d "$G" ]] && [[ -n "$(cat "$G/UDC" 2>/dev/null || true)" ]]; then
    echo "active - UDC: $(cat "$G/UDC")"
    echo "backing image: $(cat "$G/functions/mass_storage.0/lun.0/file" 2>/dev/null || echo '?')"
  else
    echo "inactive"
  fi
}

require_root
case "${1:-status}" in
  on)     gadget_on ;;
  off)    gadget_off ;;
  status) gadget_status ;;
  *) echo "usage: $0 {on|off|status}" >&2; exit 1 ;;
esac
