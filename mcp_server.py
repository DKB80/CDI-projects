#!/usr/bin/env python3
"""
MCP (Model Context Protocol) server exposing the Outlook scraper as a tool
for Claude Desktop.

Once configured, a co-worker can type into Claude Desktop chat:

    "Scrape my Outlook for Project 402. Keywords: 'Project 402',
     'Rio Tinto', 'Port Hedland'. Exclude anything about invoices or
     annual leave. Last 6 months. Call it 'Project 402 – Port Hedland
     Expansion'."

...and Claude will call `scrape_outlook` here, which drives the user's
local Classic Outlook via COM and returns the briefing (markdown) plus
the paths to the saved .msg files, attachments, and Word doc.

INSTALL (per-user, one-time)
----------------------------
1. Windows + Classic Outlook (signed in) + Python 3.10+ + Claude Desktop.
2. Clone this repo to a stable path, e.g. C:\\Tools\\cdi-outlook-scraper.
3. cd into it, then:
     pip install -r requirements.txt
4. Open Claude Desktop's config:
     %APPDATA%\\Claude\\claude_desktop_config.json
   (create it if it doesn't exist). Add the "outlook-scraper" block under
   "mcpServers":

     {
       "mcpServers": {
         "outlook-scraper": {
           "command": "python",
           "args": ["C:\\\\Tools\\\\cdi-outlook-scraper\\\\mcp_server.py"],
           "env": {
             "ANTHROPIC_API_KEY": "sk-ant-..."
           }
         }
       }
     }

5. Quit and re-open Claude Desktop. You should see "outlook-scraper" in
   the tools list (the hammer icon in the input bar).
6. Make sure Classic Outlook is running, then chat as normal.

TEAM ROLLOUT (once CDI has Claude Team)
----------------------------------------
Each co-worker does the 6 steps above on their own Windows laptop. The
Anthropic API key in step 4 can be their individual key or a team-managed
one. Everything else (credentials, mail data) stays local to their machine.
"""

import logging
import os
import sys
from datetime import datetime
from pathlib import Path

from dotenv import load_dotenv

try:
    from mcp.server.fastmcp import Context, FastMCP
except ImportError:
    sys.stderr.write(
        "ERROR: 'mcp' package not installed.\n"
        "Run: pip install -r requirements.txt\n"
    )
    sys.exit(1)

from scrape_outlook_local import scrape_outlook

load_dotenv()

# MCP servers communicate over stdio. Any print() to stdout breaks the
# protocol, so route all logging to stderr.
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
    stream=sys.stderr,
)
log = logging.getLogger("outlook-scraper")

mcp = FastMCP("outlook-scraper")


@mcp.tool()
def scrape_outlook_mailbox(
    keywords: list[str],
    project_label: str,
    exclude_keywords: list[str] | None = None,
    months: int = 12,
    output_dir: str | None = None,
    include_folders_only: list[str] | None = None,
    extra_exclude_folders: list[str] | None = None,
    include_deleted: bool = False,
    include_junk: bool = False,
    skip_attachments: bool = False,
    generate_summary: bool = True,
    max_emails_for_summary: int = 200,
    ctx: Context = None,
) -> dict:
    """Search the user's local Classic Outlook mailbox for emails matching
    keywords over the last N months, save copies (.msg + attachments) to a
    folder, and generate a Claude Opus 4.7 briefing (markdown + Word doc).

    Requires Classic Outlook to be running on the same machine as this
    server. Reads the user's local mail cache — no Azure, no Graph API.

    Args:
        keywords: Search terms joined by OR (case-insensitive, matches
            subject OR body). Example: ["Project 402", "Rio Tinto"].
        project_label: Name used in the summary report's H1 heading.
            Example: "Project 402 – Port Hedland Expansion".
        exclude_keywords: Emails containing any of these phrases are
            dropped even if a search keyword matched. Useful for filtering
            out noise like invoices, timesheets, out-of-office replies.
        months: Look back this many months. Default 12.
        output_dir: Absolute folder path to write results to. If omitted,
            a timestamped folder under ~/Documents/outlook_scrapes is used.
        include_folders_only: If set, only scan folders whose path contains
            any of these substrings (e.g. ["Project Folders", "Inbox"]).
        extra_exclude_folders: Additional folder names to skip beyond the
            built-in system folders (Deleted Items, Junk, Sync Issues, etc.).
        include_deleted: Also search Deleted Items. Default False.
        include_junk: Also search Junk Email. Default False.
        skip_attachments: Skip extracting attachments — they still live
            inside the .msg files. Default False (attachments extracted).
        generate_summary: Run Claude Opus 4.7 to produce the briefing.
            Default True.
        max_emails_for_summary: Cap on the number of emails fed into the
            summary prompt (most recent kept). Default 200.

    Returns:
        {
            "match_count": int,
            "saved_count": int,
            "emails_dir": str,
            "index_path": str | None,
            "summary_markdown": str | None,   # full briefing text
            "summary_md_path": str | None,
            "summary_docx_path": str | None,
            "errors": list[str],
        }
    """
    def progress(msg: str) -> None:
        log.info(msg)
        if ctx is not None:
            try:
                # FastMCP Context.info is async; fire-and-forget since we're sync.
                # In practice, logging to stderr is enough — Claude Desktop
                # doesn't render progress notifications yet.
                pass
            except Exception:
                pass

    if output_dir is None:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        safe_label = "".join(c if c.isalnum() or c in "-_ " else "_" for c in project_label).strip()
        output_dir = str(Path.home() / "Documents" / "outlook_scrapes" / f"{ts}_{safe_label}")

    log.info(f"scrape_outlook_mailbox starting: project={project_label!r}, "
             f"keywords={keywords}, exclude={exclude_keywords}, months={months}")
    try:
        return scrape_outlook(
            keywords=keywords,
            exclude_keywords=exclude_keywords,
            project_label=project_label,
            months=months,
            output=output_dir,
            include_deleted=include_deleted,
            include_junk=include_junk,
            extra_exclude_folders=extra_exclude_folders,
            include_folders_only=include_folders_only,
            skip_attachments=skip_attachments,
            generate_summary=generate_summary,
            max_emails_for_summary=max_emails_for_summary,
            log=progress,
        )
    except Exception as e:
        log.exception("scrape failed")
        return {
            "match_count": 0,
            "saved_count": 0,
            "emails_dir": output_dir,
            "errors": [f"Scrape failed: {e}"],
        }


if __name__ == "__main__":
    if not os.environ.get("ANTHROPIC_API_KEY"):
        log.warning(
            "ANTHROPIC_API_KEY not set — summaries will be skipped. "
            "Set it in claude_desktop_config.json or .env."
        )
    mcp.run()
