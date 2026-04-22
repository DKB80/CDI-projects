How to add the CDI logo to the tool
====================================

1. Header logo (appears at the top-left of the main window):
   Drop a PNG file named exactly  cdi_logo.png  in this folder.
   Recommended: 64 pixels tall, transparent background.
   Show it in: automatic — the main window loads it if present.

2. Window / taskbar / .exe icon:
   Drop an ICO file named exactly  cdi_logo.ico  in this folder.
   Windows ICOs contain multiple sizes — create with any online
   PNG→ICO converter; include 16x16, 32x32, 48x48, 256x256.
   Show it in: edit cdi_outlook_gui.spec (root of project) and
   uncomment the  icon="cdi_gui/assets/cdi_logo.ico"  line, then
   re-run  pyinstaller cdi_outlook_gui.spec

3. Brand colour / theme:
   Edit cdi_gui/main_window.py, top of file:
     CDI_THEME = "cosmo"     -> try "flatly", "yeti", "journal",
                                "litera", "minty", "pulse",
                                "sandstone", "united", "lumen",
                                "morph", "simplex", "cerculean"
   Each theme is a full palette — preview them at
   https://ttkbootstrap.readthedocs.io/en/latest/themes/

4. Company name + tagline:
   Also in main_window.py top:
     CDI_COMPANY, CDI_APP_TITLE, CDI_TAGLINE
