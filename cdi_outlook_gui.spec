# PyInstaller spec for the CDI Outlook Briefing Tool
# Build with:  pyinstaller cdi_outlook_gui.spec
#
# Produces dist/CDI Outlook Briefing Tool/  — a folder with a single .exe
# that co-workers can run without installing Python.
#
# The folder-based (--onedir) output is preferred over --onefile because:
#   * faster startup (no self-extract to temp)
#   * fewer false antivirus positives on managed laptops
#   * easier to hotfix — swap a single .py file and redistribute

import sys
from pathlib import Path

block_cipher = None

# Pandoc binary is only a fallback renderer now — the CDI-branded Word
# doc is produced by python-docx (pure Python, no external binaries).
# If pypandoc-binary is installed AND its bundled pandoc file is
# actually at the expected location, ship it too for belt-and-braces;
# otherwise skip silently.
datas = []
try:
    import pypandoc
    p = Path(pypandoc.get_pandoc_path())
    # pypandoc on Windows sometimes returns the path without .exe
    if not p.exists() and sys.platform == "win32":
        if p.with_suffix(".exe").exists():
            p = p.with_suffix(".exe")
    if p.exists() and p.is_file():
        datas.append((str(p), "pypandoc/files"))
        print(f"[spec] Bundling pandoc from {p}")
    else:
        print(f"[spec] pandoc not found at {p} — skipping pandoc bundle "
              f"(python-docx handles the normal path).")
except Exception as e:
    print(f"[spec] pypandoc not available ({e}) — skipping pandoc bundle.")

# Ship branding assets (logo, etc.) if present.
assets_dir = Path("cdi_gui/assets")
if assets_dir.exists():
    datas.append((str(assets_dir), "cdi_gui/assets"))

a = Analysis(
    ["cdi_outlook_gui.py"],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=[
        "pythoncom",
        "win32com.client",
        "anthropic",
        "pypandoc",
        "ttkbootstrap",
        "docx",
        "cdi_gui",
        "cdi_gui.main_window",
        "cdi_gui.wizard",
        "cdi_gui.preview",
        "cdi_gui.sender_picker",
        "cdi_gui.config",
        "cdi_gui.outlook_info",
        "cdi_gui.offline_summary",
        "cdi_gui.report_docx",
        "scrape_outlook_local",
        "summarize",
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="CDI Outlook Briefing Tool",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,       # hide the black console window
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    # icon="cdi_gui/assets/cdi_logo.ico",  # uncomment once a CDI .ico is dropped in
)
coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name="CDI Outlook Briefing Tool",
)
