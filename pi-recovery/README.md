# CDI Recovery Tool — Pi Zero 2W (base build)

A Raspberry Pi Zero 2W provisioned as a **field recovery appliance for
hardware you own or are contracted to service**. This repo contains the
neutral infrastructure: bootable-USB recovery media, local installer hosting,
a self-hosted WiFi AP, and a status web panel.

> **Scope note.** This build covers the parts that stand on their own as
> legitimate IT recovery: booting your own machine from recovery media to
> reset a lost *local* password, and installing a managed remote-support
> agent (Tailscale) openly. The **HID keyboard gadget and any keystroke /
> DuckyScript payloads are intentionally not included** — see
> [Out of scope](#out-of-scope--extension-points). Only ever point this at
> equipment you own or have written authorization to service.

## What's here

| File | Purpose |
|------|---------|
| `setup.sh` | Provisions the Pi: OTG mass-storage gadget, deps, dir structure, Tailscale (open install), WiFi AP, web-panel service. Idempotent. |
| `gadget/cdi-gadget-massstorage.sh` | Presents `recovery.img` to a host as a USB drive (mass storage only — no HID). |
| `build-recovery-image.sh` | Builds the bootable Ventoy + SystemRescue `recovery.img` (SystemRescue includes `chntpw`). |
| `fetch-installers.sh` | Mirrors the official Tailscale installers locally for offline on-site installs. |
| `server.py` | Flask status panel + local installer host + session-note log. |
| `web/index.html` | Status web panel (self-contained, polls `/api/status`). |

## Prerequisites

- Raspberry Pi Zero 2W, Raspberry Pi OS Lite 32-bit (Bookworm), 8GB+ SD card.
- SSH access to the Pi (`ssh pi@cdi-recovery.local`).
- A Tailscale account for your own fleet (auth is done interactively on the Pi;
  no pre-auth key is baked into any file in this build).

## Install

From your machine, copy the build onto the Pi and run setup:

```bash
# 1. Copy this directory to the Pi
scp -r pi-recovery pi@cdi-recovery.local:/tmp/pi-recovery

# 2. Provision (set a real AP passphrase)
ssh pi@cdi-recovery.local \
  "sudo AP_PASSPHRASE='your-strong-pass' bash /tmp/pi-recovery/setup.sh"

# 3. Put the app files where the service expects them
ssh pi@cdi-recovery.local "sudo cp /tmp/pi-recovery/server.py /opt/cdi-recovery/server.py \
  && sudo cp /tmp/pi-recovery/web/index.html /opt/cdi-recovery/web/index.html \
  && sudo systemctl restart cdi-recovery"

# 4. Build the bootable recovery image (check current SystemRescue/Ventoy versions first)
ssh pi@cdi-recovery.local "sudo bash /tmp/pi-recovery/build-recovery-image.sh"

# 5. Mirror the Tailscale installers
ssh pi@cdi-recovery.local "bash /tmp/pi-recovery/fetch-installers.sh"

# 6. Authenticate the Pi to your own tailnet, then reboot
ssh pi@cdi-recovery.local "sudo tailscale up"
ssh pi@cdi-recovery.local "sudo reboot"
```

After reboot, join WiFi **`CDI-Recovery`** and open
`http://192.168.4.1:5000`.

## The recovery workflow (consent-based)

1. **Lost local password on your own machine.** Attach the Pi's OTG (data)
   port to the target and power the Pi from its separate power port:
   ```bash
   ssh pi@cdi-recovery.local "sudo /usr/local/bin/cdi-gadget-massstorage.sh on"
   ```
   Boot the target from the Pi, run SystemRescue, and use `chntpw` (or
   mount + edit) to reset the **local** admin account. Detach with
   `cdi-gadget-massstorage.sh off`.
2. **Re-establish managed remote support.** On the now-accessible machine,
   install Tailscale from the locally hosted installer (grab it from the web
   panel, or `http://192.168.4.1:5000/installers/...`) and sign it in to your
   tailnet openly. This is a visible, managed agent — not a hidden tunnel.

## Verify

```bash
ssh pi@cdi-recovery.local "curl -s http://localhost:5000/api/status | python3 -m json.tool"
ssh pi@cdi-recovery.local "systemctl is-active cdi-recovery hostapd dnsmasq"
```

## Out of scope / extension points

The following are **deliberately not built here** because they are the
keystroke-injection mechanism itself, not recovery infrastructure. If you
develop them yourself, the base build is structured to accommodate them
without changes to what's above:

- **HID keyboard gadget / Phase-2 injection.** `cdi-gadget-massstorage.sh` is
  intentionally single-function. A separate gadget script of your own can add
  an HID and/or USB-ethernet function; it won't collide with this one as long
  as it tears down the `usb_gadget/cdi` config before reconfiguring.
- **Payloads.** `setup.sh` creates an empty `/opt/cdi-recovery/payloads/`
  directory but ships nothing in it and `server.py` does not read from it.
  Anything you place there is yours to manage.
- **Payload-serving routes.** `server.py` exposes only status, installer
  downloads, and a note log. Keep any payload/injection logic in a separate
  module behind an explicit, authenticated route so it stays isolated from the
  neutral panel.

## Notes & caveats

- These scripts were authored for the target environment but **have not been
  run against physical Pi hardware from here** — test on your bench Pi before
  field use.
- `setup.sh` writes a WPA passphrase into `/etc/hostapd/hostapd.conf` (mode
  `600`). Set a strong `AP_PASSPHRASE`; don't ship the placeholder.
- No Tailscale pre-auth key is embedded anywhere in this build. Authenticate
  interactively (`tailscale up`) so credentials aren't baked into files.
