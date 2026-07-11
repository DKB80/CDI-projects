#!/usr/bin/env bash
#
# build-recovery-image.sh - build the bootable SystemRescue recovery image
# ------------------------------------------------------------------------
# Builds ${CDI_ROOT}/recovery.img: a Ventoy disk image containing a
# SystemRescue ISO. Boot a machine YOU OWN from this (via the mass-storage
# gadget) to run SystemRescue's tools - including chntpw - to reset a lost
# LOCAL admin password on your own hardware.
#
# Run as root on the Pi:  sudo bash build-recovery-image.sh
#
# Override versions/URLs via env if the pinned releases have moved on:
#   SYSRESCUE_URL=...  VENTOY_VER=1.0.99  sudo -E bash build-recovery-image.sh
#
set -euo pipefail

CDI_ROOT="${CDI_ROOT:-/opt/cdi-recovery}"
IMG="${IMG:-$CDI_ROOT/recovery.img}"
IMG_SIZE_MB="${IMG_SIZE_MB:-2048}"
WORK="${WORK:-/tmp/cdi-build}"

# Check current releases before running:
#   SystemRescue: https://www.system-rescue.org/Download/
#   Ventoy:       https://github.com/ventoy/Ventoy/releases
SYSRESCUE_URL="${SYSRESCUE_URL:-https://fastly-cdn.system-rescue.org/releases/11.02/systemrescue-11.02-amd64.iso}"
VENTOY_VER="${VENTOY_VER:-1.0.99}"

log() { printf '\n\033[1;36m==> %s\033[0m\n' "$*"; }

[[ $EUID -eq 0 ]] || { echo "run as root" >&2; exit 1; }
for t in curl parted mkfs.vfat losetup partprobe; do
  command -v "$t" >/dev/null || { echo "missing tool: $t (apt-get install dosfstools parted util-linux)" >&2; exit 1; }
done

mkdir -p "$WORK" "$(dirname "$IMG")"

LOOP=""
cleanup() {
  mountpoint -q "$WORK/mnt" && umount "$WORK/mnt" 2>/dev/null || true
  [[ -n "$LOOP" ]] && losetup -d "$LOOP" 2>/dev/null || true
}
trap cleanup EXIT

log "Downloading SystemRescue ISO (~800MB, several minutes)"
if [[ ! -f "$WORK/systemrescue.iso" ]]; then
  curl -fL --progress-bar "$SYSRESCUE_URL" -o "$WORK/systemrescue.iso"
else
  echo "  reusing $WORK/systemrescue.iso"
fi

log "Creating ${IMG_SIZE_MB}MB image"
dd if=/dev/zero of="$IMG" bs=1M count="$IMG_SIZE_MB" status=progress

log "Attaching loop device"
LOOP="$(losetup --find --show "$IMG")"
echo "  $LOOP"

log "Partitioning (MBR, single FAT32, bootable)"
parted -s "$LOOP" mklabel msdos
parted -s "$LOOP" mkpart primary fat32 1MiB 100%
parted -s "$LOOP" set 1 boot on
partprobe "$LOOP"; sleep 2
mkfs.vfat -F 32 -n VENTOY "${LOOP}p1"

log "Copying ISO into image"
mkdir -p "$WORK/mnt"
mount "${LOOP}p1" "$WORK/mnt"
mkdir -p "$WORK/mnt/ISO"
cp "$WORK/systemrescue.iso" "$WORK/mnt/ISO/"
umount "$WORK/mnt"

log "Installing Ventoy bootloader (v${VENTOY_VER})"
if [[ ! -d "$WORK/ventoy" ]]; then
  curl -fL --progress-bar \
    "https://github.com/ventoy/Ventoy/releases/download/v${VENTOY_VER}/ventoy-${VENTOY_VER}-linux.tar.gz" \
    -o "$WORK/ventoy.tar.gz"
  mkdir -p "$WORK/ventoy"
  tar xzf "$WORK/ventoy.tar.gz" -C "$WORK/ventoy" --strip-components=1
fi
# -I forces a fresh install to the whole loop device.
( cd "$WORK/ventoy" && echo y | bash Ventoy2Disk.sh -I "$LOOP" -L VENTOY )

log "Detaching and verifying"
losetup -d "$LOOP"; LOOP=""
chown "${CDI_USER:-pi}":"${CDI_USER:-pi}" "$IMG" 2>/dev/null || true

SIZE=$(stat -c%s "$IMG")
echo "  image size: $SIZE bytes"
if [[ "$SIZE" -gt 1000000000 ]]; then
  echo "  OK - recovery.img built"
else
  echo "  ERROR: image looks too small; something went wrong" >&2
  exit 1
fi
