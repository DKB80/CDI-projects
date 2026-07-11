#!/usr/bin/env bash
#
# fetch-installers.sh - mirror the official Tailscale installers locally
# ----------------------------------------------------------------------
# Downloads the current official Tailscale installers into $CDI_ROOT/installers
# so an on-site machine you're servicing can install the managed remote-support
# agent without depending on site internet. These are the vendor's own signed
# installers, served unmodified.
#
# Usage:  bash fetch-installers.sh
#
set -euo pipefail

CDI_ROOT="${CDI_ROOT:-/opt/cdi-recovery}"
DEST="$CDI_ROOT/installers"
mkdir -p "$DEST"

fetch() {
  local url="$1" out="$2"
  echo "==> $out"
  if curl -fL --progress-bar "$url" -o "$DEST/$out.part"; then
    mv "$DEST/$out.part" "$DEST/$out"
    ls -lh "$DEST/$out"
  else
    echo "[warn] failed to fetch $url" >&2
    rm -f "$DEST/$out.part"
    return 1
  fi
}

# Official Tailscale download endpoints (see https://pkgs.tailscale.com/stable/).
fetch "https://pkgs.tailscale.com/stable/tailscale-setup-latest-amd64.msi" "tailscale-windows.msi"  || true
fetch "https://pkgs.tailscale.com/stable/Tailscale-latest-macos.pkg"        "tailscale-macos.pkg"    || true

echo
echo "Done. Hosted installers:"
ls -lh "$DEST" || true
