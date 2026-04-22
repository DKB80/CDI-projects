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

# Pandoc binary bundled by pypandoc-binary needs to ship with the app.
datas = []
try:
    import pypandoc
    pandoc_path = Path(pypandoc.get_pandoc_path())
    datas.append((str(pandoc_path), "pypandoc/files"))
except Exception:
    # If pypandoc isn't available at build time, just skip — user can install it at runtime.
    pass

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
