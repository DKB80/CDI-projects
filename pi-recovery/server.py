#!/usr/bin/env python3
"""
CDI Recovery - status web panel + local installer host.

Scope of this server (deliberately narrow and neutral):
  * Serve a status page and a JSON status API for the field tech.
  * Host the official Tailscale installers locally so an on-site machine
    can install a managed remote-support agent without site internet.
  * Report gadget / service / tailscale state.
  * Append session notes to a JSON log for your own record keeping.

It does NOT drive keystroke injection or push anything to a target machine.
If you later add your own payload tooling, keep it in a separate module and
wire it in behind an explicit, authenticated route - see README.md.
"""

import json
import os
import subprocess
import time
from pathlib import Path

from flask import Flask, Response, abort, jsonify, request, send_from_directory

CDI_ROOT = Path(os.environ.get("CDI_ROOT", "/opt/cdi-recovery"))
WEB_PORT = int(os.environ.get("WEB_PORT", "5000"))
WEB_DIR = CDI_ROOT / "web"
INSTALLERS_DIR = CDI_ROOT / "installers"
LOGS_DIR = CDI_ROOT / "logs"
RECOVERY_IMG = CDI_ROOT / "recovery.img"

app = Flask(__name__, static_folder=None)


def _human_size(num: int) -> str:
    for unit in ("B", "KB", "MB", "GB", "TB"):
        if num < 1024:
            return f"{num:.0f}{unit}" if unit == "B" else f"{num:.1f}{unit}"
        num /= 1024
    return f"{num:.1f}PB"


def _service_active(name: str) -> bool:
    try:
        out = subprocess.run(
            ["systemctl", "is-active", name],
            capture_output=True, text=True, timeout=5,
        )
        return out.stdout.strip() == "active"
    except Exception:
        return False


def _tailscale_status() -> str:
    try:
        out = subprocess.run(
            ["tailscale", "status"], capture_output=True, text=True, timeout=5
        )
        if out.returncode == 0 and out.stdout.strip():
            return "authenticated"
        return "not-authenticated"
    except FileNotFoundError:
        return "not-installed"
    except Exception:
        return "unknown"


def _list_installers():
    items = []
    if INSTALLERS_DIR.is_dir():
        for p in sorted(INSTALLERS_DIR.iterdir()):
            if p.is_file():
                items.append({
                    "name": p.name,
                    "size": p.stat().st_size,
                    "size_human": _human_size(p.stat().st_size),
                    "url": f"/installers/{p.name}",
                })
    return items


@app.get("/api/status")
def api_status():
    services = {
        svc: ("active" if _service_active(svc) else "inactive")
        for svc in ("cdi-recovery", "hostapd", "dnsmasq")
    }
    img_ok = RECOVERY_IMG.is_file() and RECOVERY_IMG.stat().st_size > 0
    return jsonify({
        "status": "online",
        "hostname": os.uname().nodename,
        "uptime_seconds": _read_uptime(),
        "services": services,
        "tailscale": _tailscale_status(),
        "recovery_image": {
            "present": img_ok,
            "size_human": _human_size(RECOVERY_IMG.stat().st_size) if img_ok else None,
        },
        "installers": _list_installers(),
    })


def _read_uptime() -> int:
    try:
        return int(float(Path("/proc/uptime").read_text().split()[0]))
    except Exception:
        return -1


@app.get("/installers/<path:name>")
def installers(name):
    if not INSTALLERS_DIR.is_dir():
        abort(404)
    return send_from_directory(INSTALLERS_DIR, name, as_attachment=True)


@app.post("/api/log")
def api_log():
    """Append a free-form session note (for your own record keeping)."""
    LOGS_DIR.mkdir(parents=True, exist_ok=True)
    payload = request.get_json(silent=True) or {}
    entry = {
        "ts": int(time.time()),
        "note": str(payload.get("note", ""))[:2000],
        "site": str(payload.get("site", ""))[:200],
        "remote_addr": request.remote_addr,
    }
    logfile = LOGS_DIR / "sessions.jsonl"
    with logfile.open("a") as fh:
        fh.write(json.dumps(entry) + "\n")
    return jsonify({"ok": True})


@app.get("/")
def index():
    idx = WEB_DIR / "index.html"
    if idx.is_file():
        return Response(idx.read_text(), mimetype="text/html")
    return Response(
        "<h1>CDI Recovery</h1><p>web/index.html not deployed yet.</p>",
        mimetype="text/html",
    )


if __name__ == "__main__":
    # Bind to all interfaces so the AP clients (192.168.4.x) can reach it.
    app.run(host="0.0.0.0", port=WEB_PORT)
