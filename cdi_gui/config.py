"""Config + preset storage for the CDI Outlook Briefing Tool."""

import json
import os
from pathlib import Path


APP_DIR_NAME = "CDI-OutlookScraper"
_local = os.environ.get("LOCALAPPDATA") or str(Path.home() / "AppData" / "Local")
CONFIG_DIR = Path(_local) / APP_DIR_NAME
CONFIG_FILE = CONFIG_DIR / "config.json"
PRESETS_DIR = CONFIG_DIR / "presets"
DOWNLOADS_DIR = Path.home() / "Downloads"
DEFAULT_OUTPUT_ROOT = Path.home() / "Documents" / "CDI outlook scrapes"


def ensure_dirs() -> None:
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    PRESETS_DIR.mkdir(parents=True, exist_ok=True)
    DEFAULT_OUTPUT_ROOT.mkdir(parents=True, exist_ok=True)


def load_config() -> dict:
    ensure_dirs()
    if CONFIG_FILE.exists():
        try:
            return json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
        except Exception:
            return {}
    return {}


def save_config(cfg: dict) -> None:
    ensure_dirs()
    CONFIG_FILE.write_text(json.dumps(cfg, indent=2), encoding="utf-8")


def is_first_run() -> bool:
    cfg = load_config()
    return not cfg.get("setup_complete")


def list_presets() -> list[str]:
    ensure_dirs()
    return sorted(p.stem for p in PRESETS_DIR.glob("*.json"))


def load_preset(name: str) -> dict:
    path = PRESETS_DIR / f"{name}.json"
    if not path.exists():
        raise FileNotFoundError(f"Preset '{name}' not found.")
    return json.loads(path.read_text(encoding="utf-8"))


def save_preset(name: str, data: dict) -> Path:
    ensure_dirs()
    safe = "".join(c if c.isalnum() or c in "-_ " else "_" for c in name).strip() or "preset"
    path = PRESETS_DIR / f"{safe}.json"
    path.write_text(json.dumps(data, indent=2), encoding="utf-8")
    return path


def delete_preset(name: str) -> None:
    path = PRESETS_DIR / f"{name}.json"
    if path.exists():
        path.unlink()
