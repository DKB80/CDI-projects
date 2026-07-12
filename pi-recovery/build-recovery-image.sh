#!/usr/bin/env bash
#
# build-recovery-image.sh - build the bootable SystemRescue recovery image
# ------------------------------------------------------------------------
# Builds ${CDI_ROOT}/recovery.img: the SystemRescue ISO itself, used raw.
# Boot a machine YOU OWN from this (via the mass-storage gadget) to run
# SystemRescue's tools - including chntpw - to reset a lost LOCAL admin
# password on your own hardware.
#
# SystemRescue ISOs are isohybrid: written to / exposed as a raw disk they
# boot directly on x86 BIOS and UEFI. So we use the ISO as recovery.img
# directly - no Ventoy. (The previous version used Ventoy, which ships no
# 32-bit-ARM build and therefore cannot run on a 32-bit Raspberry Pi OS.)
#
# Run as root on the Pi:  sudo bash build-recovery-image.sh
#
# Override version/URL via env if the pinned release has moved on:
#   SYSRESCUE_VER=13.01  sudo -E bash build-recovery-image.sh
#   SYSRESCUE_URL=https://.../systemrescue-XX.YY-amd64.iso  sudo -E bash build-recovery-image.sh
#
set -euo pipefail

CDI_ROOT="${CDI_ROOT:-/opt/cdi-recovery}"
IMG="${IMG:-$CDI_ROOT/recovery.img}"

# Check current releases before running: https://www.system-rescue.org/Download/
SYSRESCUE_VER="${SYSRESCUE_VER:-13.01}"
SYSRESCUE_URL="${SYSRESCUE_URL:-https://fastly-cdn.system-rescue.org/releases/${SYSRESCUE_VER}/systemrescue-${SYSRESCUE_VER}-amd64.iso}"

log() { printf '\n\033[1;36m==> %s\033[0m\n' "$*"; }

[[ $EUID -eq 0 ]] || { echo "run as root" >&2; exit 1; }
command -v curl >/dev/null || { echo "missing tool: curl" >&2; exit 1; }

mkdir -p "$(dirname "$IMG")"

log "Checking SystemRescue ${SYSRESCUE_VER} URL"
head1="$(curl -sI -m20 "$SYSRESCUE_URL" | head -1 || true)"
echo "  $SYSRESCUE_URL"
echo "  $head1"
case "$head1" in
  *200*) : ;;
  *) echo "ERROR: ISO URL not reachable (got: $head1). Check the current version at" >&2
     echo "       https://www.system-rescue.org/Download/ and pass SYSRESCUE_VER=..." >&2
     exit 1 ;;
esac

log "Downloading SystemRescue ISO directly to recovery.img (~1.3GB, several minutes)"
curl -fL --progress-bar -m3600 "$SYSRESCUE_URL" -o "$IMG.part"
mv "$IMG.part" "$IMG"
chown "${CDI_USER:-pi}":"${CDI_USER:-pi}" "$IMG" 2>/dev/null || true

log "Verifying image is a bootable isohybrid ISO"
SIZE=$(stat -c%s "$IMG")
echo "  size: $SIZE bytes ($(( SIZE / 1024 / 1024 )) MB)"
# ISO9660 volume descriptor magic "CD001" lives at byte offset 0x8001.
magic=$(dd if="$IMG" bs=1 skip=32769 count=5 2>/dev/null || true)
# isohybrid ISOs carry an MBR with the 0x55AA boot signature at byte 510.
sig=$(dd if="$IMG" bs=1 skip=510 count=2 2>/dev/null | od -An -tx1 | tr -d ' ')
echo "  iso9660 magic: '$magic' (expect CD001)"
echo "  mbr boot sig:  $sig (expect 55aa)"
if [[ "$SIZE" -gt 1000000000 && "$magic" == "CD001" && "$sig" == "55aa" ]]; then
  echo "  OK - recovery.img is a bootable SystemRescue image"
else
  echo "  ERROR: image failed verification (size/magic/signature)" >&2
  exit 1
fi

cat <<EOF

Done. recovery.img = SystemRescue ${SYSRESCUE_VER}.
Present it to a target machine you own with:
    sudo /usr/local/bin/cdi-gadget-massstorage.sh on
For best BIOS compatibility you can expose it read-only as a CD-ROM by
setting the gadget's lun.0/cdrom=1 and lun.0/ro=1 (see cdi-gadget-massstorage.sh).
EOF
