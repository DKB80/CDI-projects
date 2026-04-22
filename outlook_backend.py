"""Platform dispatcher for the Outlook scraper.

The GUI and every other caller imports scrape_outlook / list_recent_senders
from here — this module picks the right implementation for the current
OS (pywin32 COM on Windows, AppleScript/JXA on macOS) so downstream code
doesn't need any `if sys.platform` checks.
"""

import sys

if sys.platform == "darwin":
    from scrape_outlook_mac import (  # noqa: F401
        scrape_outlook,
        list_recent_senders,
    )
elif sys.platform == "win32":
    from scrape_outlook_local import (  # noqa: F401
        scrape_outlook,
        list_recent_senders,
    )
else:
    raise ImportError(
        f"Unsupported platform: {sys.platform}. "
        "This tool supports Windows (Classic Outlook) and macOS (Outlook for Mac)."
    )
