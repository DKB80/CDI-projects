#!/usr/bin/env python3
"""
Scrape Classic Outlook desktop (Windows) for project-related emails via COM,
save copies as .msg + attachments to a local folder, and generate a
Claude-powered briefing with action items.

No Azure, no API credentials, no network calls to Microsoft — reads what
Outlook already has cached on your machine.

SETUP
-----
1. Windows with Classic Outlook desktop (File / Home / Send ribbon). Your
   account (dan.bailey@cdienergy.com.au) must be signed in and the mail
   folders visible in Outlook.
2. pip install -r requirements.txt
   (pywin32 installs only on Windows; the Linux/Mac install will skip it.)
3. Copy .env.example to .env and fill in ANTHROPIC_API_KEY. CLIENT_ID and
   TENANT_ID are only needed for the Azure version; leave them blank here.
4. Open Outlook (it must be running), then from the same Windows machine:
     python scrape_outlook_local.py

USAGE
-----
  python scrape_outlook_local.py                       # defaults: 12 months, all folders
  python scrape_outlook_local.py --months 24
  python scrape_outlook_local.py --keywords "Horizon Power" "Hossein"
  python scrape_outlook_local.py --no-summary
  python scrape_outlook_local.py --include-deleted     # also search Deleted Items

OUTPUT
------
  output/
    emails/<date>__<sender>__<subject>.msg   # native Outlook files (open in Outlook)
    attachments/<folder>/<filename>           # extracted attachments per email
    index.json                                # machine-readable index
    SUMMARY.md                                # Claude-generated briefing

NOTES
-----
- Outlook cached mode: if your Outlook is set to keep <12 months offline,
  older emails still search but may pull from the server (slower). Check
  File > Account Settings > Account Settings > Exchange Account > Change >
  "Mail to keep offline" if you see gaps.
- New Outlook (the modern one) does NOT expose COM. You need the Classic
  desktop client for this to work.
"""

import argparse
import json
import os
import re
import sys
from datetime import datetime, timedelta
from pathlib import Path

try:
    import pythoncom  # noqa: F401
    import win32com.client
except ImportError:
    sys.stderr.write(
        "ERROR: pywin32 not installed.\n"
        "This script only works on Windows with Classic Outlook desktop.\n"
        "Run: pip install pywin32\n"
    )
    sys.exit(1)

from dotenv import load_dotenv

from summarize import summarize_emails

load_dotenv()

DEFAULT_KEYWORDS = ["Horizon Power", "307", "Hossein", "remote communities"]

# Folders to skip by default. User can override with --include-deleted / --include-junk.
SKIP_FOLDERS_DEFAULT = {
    "Deleted Items",
    "Junk Email",
    "Sync Issues",
    "Conversation History",
    "RSS Feeds",
    "Conflicts",
    "Local Failures",
    "Server Failures",
    "Recoverable Items",
}

OL_MAIL = 43
OL_SAVE_MSG = 3


def sanitize_filename(name: str, max_len: int = 120) -> str:
    cleaned = re.sub(r"[^\w\-.() ]+", "_", name).strip("._ ")
    return (cleaned or "unnamed")[:max_len]


def walk_folders(folder, skip_names: set[str]):
    if folder.Name in skip_names:
        return
    yield folder
    try:
        subfolders = list(folder.Folders)
    except Exception:
        subfolders = []
    for sub in subfolders:
        yield from walk_folders(sub, skip_names)


def outlook_date_filter(cutoff: datetime) -> str:
    # Outlook DASL expects US-style MM/DD/YYYY hh:mm AM/PM regardless of locale.
    return "[ReceivedTime] >= '" + cutoff.strftime("%m/%d/%Y %I:%M %p") + "'"


def search_matches(
    namespace,
    keywords: list[str],
    cutoff: datetime,
    skip_folders: set[str],
) -> list:
    matches = []
    seen: set[str] = set()
    kw_lower = [k.lower() for k in keywords]
    date_filter = outlook_date_filter(cutoff)

    stores = list(namespace.Folders)
    print(f"Found {len(stores)} mail store(s): {[s.Name for s in stores]}")

    for store in stores:
        print(f"\nScanning store: {store.Name}")
        for folder in walk_folders(store, skip_folders):
            path = getattr(folder, "FolderPath", folder.Name)
            try:
                items = folder.Items
                items.Sort("[ReceivedTime]", True)
                filtered = items.Restrict(date_filter)
            except Exception as e:
                print(f"  [skip] {path}: {e}", file=sys.stderr)
                continue

            folder_hits = 0
            item = filtered.GetFirst()
            while item is not None:
                try:
                    if getattr(item, "Class", 0) == OL_MAIL:
                        subject = item.Subject or ""
                        body = item.Body or ""
                        haystack = (subject + "\n" + body).lower()
                        if any(k in haystack for k in kw_lower):
                            entry_id = item.EntryID
                            if entry_id not in seen:
                                seen.add(entry_id)
                                matches.append(item)
                                folder_hits += 1
                except Exception as e:
                    print(f"  [item error] {path}: {e}", file=sys.stderr)
                try:
                    item = filtered.GetNext()
                except Exception:
                    break

            if folder_hits:
                print(f"  {path}: {folder_hits} match(es)")

    return matches


def save_item(item, emails_dir: Path, attachments_dir: Path, extract_attachments: bool) -> dict:
    received = item.ReceivedTime
    date_str = received.strftime("%Y-%m-%d")
    sender = (
        getattr(item, "SenderEmailAddress", "")
        or getattr(item, "SenderName", "")
        or "unknown"
    )
    subject = item.Subject or "no_subject"

    base_name = sanitize_filename(f"{date_str}__{sender}__{subject}")
    msg_path = emails_dir / (base_name + ".msg")
    if msg_path.exists():
        msg_path = emails_dir / (base_name + "__" + item.EntryID[-8:] + ".msg")
    item.SaveAs(str(msg_path), OL_SAVE_MSG)

    saved_attachments: list[str] = []
    if extract_attachments and item.Attachments.Count > 0:
        att_dir = attachments_dir / sanitize_filename(f"{date_str}__{subject}__{item.EntryID[-8:]}")
        for i in range(1, item.Attachments.Count + 1):
            att = item.Attachments.Item(i)
            if att.Type != 1:  # 1 = olByValue (actual file)
                continue
            att_name = sanitize_filename(att.FileName or f"attachment_{i}")
            att_dir.mkdir(parents=True, exist_ok=True)
            dest = att_dir / att_name
            try:
                att.SaveAsFile(str(dest))
                saved_attachments.append(str(dest))
            except Exception as e:
                print(f"    attachment save failed ({att_name}): {e}", file=sys.stderr)

    return {
        "entry_id": item.EntryID,
        "date": received.isoformat(),
        "from": sender,
        "to": item.To or "",
        "cc": item.CC or "",
        "subject": subject,
        "msg": str(msg_path),
        "attachments": saved_attachments,
        "body": item.Body or "",
    }


def main() -> int:
    parser = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("--keywords", nargs="+", default=DEFAULT_KEYWORDS,
                        help="Search terms (OR, case-insensitive, subject + body). Default: %(default)s")
    parser.add_argument("--months", type=int, default=12,
                        help="Look back this many months (default: %(default)s)")
    parser.add_argument("--output", type=Path, default=Path("output"),
                        help="Output directory (default: %(default)s)")
    parser.add_argument("--include-deleted", action="store_true",
                        help="Also search Deleted Items")
    parser.add_argument("--include-junk", action="store_true",
                        help="Also search Junk Email")
    parser.add_argument("--no-attachments", action="store_true",
                        help="Skip extracting attachments (they remain inside the .msg files)")
    parser.add_argument("--no-summary", action="store_true",
                        help="Skip Claude summary generation")
    parser.add_argument("--max-emails-for-summary", type=int, default=200,
                        help="Cap on emails fed into the summary prompt (default: %(default)s)")
    args = parser.parse_args()

    skip = set(SKIP_FOLDERS_DEFAULT)
    if args.include_deleted:
        skip.discard("Deleted Items")
    if args.include_junk:
        skip.discard("Junk Email")

    print("Connecting to Outlook...")
    try:
        app = win32com.client.Dispatch("Outlook.Application")
        ns = app.GetNamespace("MAPI")
    except Exception as e:
        print(f"Failed to connect to Outlook: {e}", file=sys.stderr)
        print("Make sure Classic Outlook is running.", file=sys.stderr)
        return 1

    cutoff = datetime.now() - timedelta(days=30 * args.months)
    print(f"Searching for {args.keywords} in emails since {cutoff.date().isoformat()}")
    print(f"Skipping folders: {sorted(skip)}")

    matches = search_matches(ns, args.keywords, cutoff, skip)
    print(f"\nTotal unique matches: {len(matches)}")
    if not matches:
        print("Nothing to do.")
        return 0

    # Sort oldest first for output consistency
    matches.sort(key=lambda m: m.ReceivedTime)

    emails_dir = args.output / "emails"
    attachments_dir = args.output / "attachments"
    emails_dir.mkdir(parents=True, exist_ok=True)

    print(f"\nSaving emails to {emails_dir}/ ...")
    index = []
    for i, item in enumerate(matches, 1):
        try:
            entry = save_item(item, emails_dir, attachments_dir, not args.no_attachments)
            index.append(entry)
        except Exception as e:
            print(f"  [{i}/{len(matches)}] save failed: {e}", file=sys.stderr)
            continue
        if i % 10 == 0 or i == len(matches):
            print(f"  [{i}/{len(matches)}] saved")

    # Write index without bodies (bodies are huge; .msg files hold them)
    index_public = [{k: v for k, v in e.items() if k != "body"} for e in index]
    (args.output / "index.json").write_text(json.dumps(index_public, indent=2))
    print(f"\nIndex: {args.output / 'index.json'}")

    if args.no_summary:
        return 0
    if not os.environ.get("ANTHROPIC_API_KEY"):
        print("\nANTHROPIC_API_KEY not set — skipping summary.", file=sys.stderr)
        return 0

    sample = index[-args.max_emails_for_summary:] if len(index) > args.max_emails_for_summary else index
    print(f"\nGenerating summary with Claude Opus 4.7 ({len(sample)} emails)...")
    summary = summarize_emails(sample, args.keywords)
    summary_path = args.output / "SUMMARY.md"
    summary_path.write_text(summary, encoding="utf-8")
    print(f"Summary: {summary_path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
