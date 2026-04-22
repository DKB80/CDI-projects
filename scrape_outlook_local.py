#!/usr/bin/env python3
"""
Scrape Classic Outlook desktop (Windows) for project-related emails via COM,
save copies as .msg + attachments to a local folder, and generate a
Claude Opus 4.7 briefing (markdown + Word doc).

No Azure, no Microsoft credentials, no network calls to Microsoft — reads
what Outlook already has cached on your machine.

The core work lives in `scrape_outlook()` and can be called from:
  - this CLI (`main()`)
  - the MCP server in `mcp_server.py` (for Claude Desktop)
  - any other Python process

SETUP
-----
1. Windows with Classic Outlook desktop (File / Home / Send ribbon). Your
   account must be signed in and the mail folders visible in Outlook.
2. pip install -r requirements.txt
3. Copy .env.example to .env and fill in ANTHROPIC_API_KEY.
4. Open Outlook (it must be running), then:
     python scrape_outlook_local.py --keywords "Project 402" "Rio Tinto"

USAGE
-----
  python scrape_outlook_local.py --keywords "Project 402" "Rio Tinto"
  python scrape_outlook_local.py --keywords "Horizon Power" --exclude "invoice" "leave"
  python scrape_outlook_local.py --keywords X --project "Horizon Power - Remote Communities"
  python scrape_outlook_local.py --keywords X --months 24 --include-deleted
  python scrape_outlook_local.py --keywords X --no-summary

NOTES
-----
- Outlook cached mode: if your Outlook is set to keep <N months offline,
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
from typing import Callable, Iterable

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

from summarize import render_docx, summarize_emails

load_dotenv()

# Defaults the CLI falls back on when --keywords / --project aren't supplied.
# These are tuned for the original Horizon Power / Project 307 brief.
CLI_DEFAULT_KEYWORDS = ["Horizon Power", "307", "Hossein", "remote communities"]
CLI_DEFAULT_PROJECT = "Project 307 – Horizon Power Remote Communities"

# System folders that are almost never what the user wants to search.
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
OL_DEFAULT_MAIL = 0
OL_DEFAULT_POST = 6
# 9 = olMSGUnicode. Do not use 3 (olMSG legacy ASCII) — it fails on Unicode.
OL_SAVE_MSG = 9

MAX_FILENAME_LEN = 120

Logger = Callable[[str], None]


def _noop(msg: str) -> None:
    pass


def sanitize_filename(name: str, max_len: int = MAX_FILENAME_LEN) -> str:
    cleaned = re.sub(r"[\x00-\x1f<>:\"/\\|?*]+", "_", name or "")
    cleaned = re.sub(r"\s+", " ", cleaned).strip("._ ")
    if cleaned.upper() in {"CON", "PRN", "AUX", "NUL"} or re.match(r"^(COM|LPT)\d$", cleaned.upper()):
        cleaned = "_" + cleaned
    return (cleaned or "unnamed")[:max_len]


def format_com_error(e: Exception) -> str:
    args = getattr(e, "args", ())
    if len(args) >= 3 and isinstance(args[2], tuple) and len(args[2]) >= 3:
        return f"{args[2][1]}: {args[2][2]}".strip()
    return str(e)


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


def _folder_path_matches_any(path: str, needles: Iterable[str]) -> bool:
    path_low = path.lower()
    return any(n.lower() in path_low for n in needles)


def search_matches(
    namespace,
    keywords: list[str],
    exclude_keywords: list[str],
    cutoff: datetime,
    skip_folders: set[str],
    include_folders_only: list[str] | None,
    log: Logger,
) -> list:
    matches = []
    seen: set[str] = set()
    kw_lower = [k.lower() for k in keywords]
    ex_lower = [e.lower() for e in (exclude_keywords or [])]
    date_filter = outlook_date_filter(cutoff)

    stores = list(namespace.Folders)
    log(f"Found {len(stores)} mail store(s): {[s.Name for s in stores]}")

    for store in stores:
        log(f"Scanning store: {store.Name}")
        for folder in walk_folders(store, skip_folders):
            path = getattr(folder, "FolderPath", folder.Name)
            default_type = getattr(folder, "DefaultItemType", None)
            if default_type is not None and default_type not in (OL_DEFAULT_MAIL, OL_DEFAULT_POST):
                continue
            if include_folders_only and not _folder_path_matches_any(path, include_folders_only):
                continue

            try:
                items = folder.Items
                items.Sort("[ReceivedTime]", True)
                filtered = items.Restrict(date_filter)
            except Exception as e:
                log(f"  [skip] {path}: {format_com_error(e)}")
                continue

            folder_hits = 0
            item = filtered.GetFirst()
            while item is not None:
                try:
                    if getattr(item, "Class", 0) == OL_MAIL:
                        subject = item.Subject or ""
                        body = item.Body or ""
                        haystack = (subject + "\n" + body).lower()
                        if any(k in haystack for k in kw_lower) and not any(
                            e in haystack for e in ex_lower
                        ):
                            entry_id = item.EntryID
                            if entry_id not in seen:
                                seen.add(entry_id)
                                matches.append(item)
                                folder_hits += 1
                except Exception as e:
                    log(f"  [item error] {path}: {e}")
                try:
                    item = filtered.GetNext()
                except Exception:
                    break

            if folder_hits:
                log(f"  {path}: {folder_hits} match(es)")

    return matches


def save_item(item, emails_dir: Path, attachments_dir: Path, extract_attachments: bool, log: Logger, save_msg: bool = True) -> dict:
    received = item.ReceivedTime
    date_str = received.strftime("%Y-%m-%d")
    sender = (
        getattr(item, "SenderEmailAddress", "")
        or getattr(item, "SenderName", "")
        or "unknown"
    )
    subject = item.Subject or "no_subject"

    msg_path: Path | None = None
    if save_msg:
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
            if att.Type != 1:
                continue
            att_name = sanitize_filename(att.FileName or f"attachment_{i}")
            att_dir.mkdir(parents=True, exist_ok=True)
            dest = att_dir / att_name
            try:
                att.SaveAsFile(str(dest))
                saved_attachments.append(str(dest))
            except Exception as e:
                log(f"    attachment save failed ({att_name}): {format_com_error(e)}")

    return {
        "entry_id": item.EntryID,
        "date": received.isoformat(),
        "from": sender,
        "to": item.To or "",
        "cc": item.CC or "",
        "subject": subject,
        "msg": str(msg_path) if msg_path else None,
        "attachments": saved_attachments,
        "body": item.Body or "",
    }


def scrape_outlook(
    *,
    keywords: list[str],
    exclude_keywords: list[str] | None = None,
    project_label: str = "Project Email Archive",
    months: int = 12,
    output: Path | str = Path("output"),
    include_deleted: bool = False,
    include_junk: bool = False,
    extra_exclude_folders: list[str] | None = None,
    include_folders_only: list[str] | None = None,
    skip_attachments: bool = False,
    save_msg_files: bool = True,
    generate_summary: bool = True,
    max_emails_for_summary: int = 200,
    log: Logger | None = None,
    preview_callback: Callable[[list[dict]], list[str] | None] | None = None,
) -> dict:
    """Core scrape — callable from CLI, MCP server, GUI, or any Python process.

    If `preview_callback` is set, after the scan completes it's called with a
    list of match metadata dicts (entry_id, subject, from, date, body_preview).
    It should return a list of approved entry_ids to keep, or None to cancel.

    Returns a dict with: match_count, saved_count, emails_dir, index_path,
    summary_markdown, summary_md_path, summary_docx_path, errors.
    """
    log = log or _noop
    if not keywords:
        raise ValueError("At least one keyword is required.")

    skip = set(SKIP_FOLDERS_DEFAULT)
    if include_deleted:
        skip.discard("Deleted Items")
    if include_junk:
        skip.discard("Junk Email")
    if extra_exclude_folders:
        skip.update(extra_exclude_folders)

    output = Path(output).absolute()
    emails_dir = output / "emails"
    attachments_dir = output / "attachments"
    emails_dir.mkdir(parents=True, exist_ok=True)
    log(f"Output directory: {output}")

    log("Connecting to Outlook...")
    try:
        app = win32com.client.Dispatch("Outlook.Application")
        ns = app.GetNamespace("MAPI")
    except Exception as e:
        raise RuntimeError(
            f"Failed to connect to Outlook: {format_com_error(e)}. "
            "Make sure Classic Outlook is running on this machine."
        ) from e

    cutoff = datetime.now() - timedelta(days=30 * months)
    log(f"Searching for {keywords} since {cutoff.date().isoformat()} "
        f"(excluding: {exclude_keywords or []})")
    log(f"Skipping folders: {sorted(skip)}")

    matches = search_matches(
        ns, keywords, exclude_keywords or [], cutoff,
        skip, include_folders_only, log,
    )
    log(f"Total unique matches: {len(matches)}")

    result: dict = {
        "match_count": len(matches),
        "saved_count": 0,
        "emails_dir": str(emails_dir),
        "index_path": None,
        "summary_markdown": None,
        "summary_md_path": None,
        "summary_docx_path": None,
        "errors": [],
        "keywords": keywords,
        "exclude_keywords": exclude_keywords or [],
        "project_label": project_label,
        "months": months,
    }
    if not matches:
        return result

    matches.sort(key=lambda m: m.ReceivedTime)

    if preview_callback is not None:
        metadata = [
            {
                "entry_id": m.EntryID,
                "subject": m.Subject or "",
                "from": (
                    getattr(m, "SenderEmailAddress", "")
                    or getattr(m, "SenderName", "")
                    or "unknown"
                ),
                "date": m.ReceivedTime.isoformat(),
                "body_preview": (m.Body or "")[:200].replace("\n", " ").replace("\r", " "),
                "has_attachments": bool(m.Attachments.Count) if hasattr(m, "Attachments") else False,
            }
            for m in matches
        ]
        approved = preview_callback(metadata)
        if approved is None:
            result["errors"].append("Cancelled by user at preview step.")
            log("Preview cancelled — nothing saved.")
            return result
        approved_set = set(approved)
        matches = [m for m in matches if m.EntryID in approved_set]
        log(f"User approved {len(matches)}/{result['match_count']} matches for save.")
        result["approved_count"] = len(matches)
        if not matches:
            return result

    log(f"Saving emails to {emails_dir} ...")

    index = []
    consecutive_failures = 0
    for i, item in enumerate(matches, 1):
        try:
            entry = save_item(item, emails_dir, attachments_dir, not skip_attachments, log, save_msg=save_msg_files)
            index.append(entry)
            consecutive_failures = 0
        except Exception as e:
            consecutive_failures += 1
            err = f"[{i}/{len(matches)}] save failed: {format_com_error(e)}"
            result["errors"].append(err)
            log(err)
            if consecutive_failures >= 3 and len(index) == 0:
                result["errors"].append(
                    f"Aborted after 3 consecutive failures. "
                    f"emails_dir exists={emails_dir.exists()}, "
                    f"writable={os.access(emails_dir, os.W_OK) if emails_dir.exists() else 'n/a'}"
                )
                return result
            continue
        if i % 20 == 0 or i == len(matches):
            log(f"  [{i}/{len(matches)}] saved, {len(index)} successful")

    result["saved_count"] = len(index)
    if not index:
        return result

    index_path = output / "index.json"
    index_public = [{k: v for k, v in e.items() if k != "body"} for e in index]
    index_path.write_text(json.dumps(index_public, indent=2))
    result["index_path"] = str(index_path)
    log(f"Index: {index_path}")

    if not generate_summary:
        return result
    if not os.environ.get("ANTHROPIC_API_KEY"):
        result["errors"].append("ANTHROPIC_API_KEY not set — summary skipped.")
        log("ANTHROPIC_API_KEY not set — skipping summary.")
        return result

    sample = index[-max_emails_for_summary:] if len(index) > max_emails_for_summary else index
    log(f"Generating summary with Claude Opus 4.7 ({len(sample)} emails)...")
    summary = summarize_emails(sample, keywords, project_label=project_label)
    result["summary_markdown"] = summary

    summary_md = output / "SUMMARY.md"
    summary_md.write_text(summary, encoding="utf-8")
    result["summary_md_path"] = str(summary_md)
    log(f"Summary (markdown): {summary_md}")

    summary_docx = output / "SUMMARY.docx"
    if render_docx(summary, summary_docx):
        result["summary_docx_path"] = str(summary_docx)
        log(f"Summary (Word):     {summary_docx}")

    return result


def main() -> int:
    parser = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("--keywords", nargs="+", default=CLI_DEFAULT_KEYWORDS,
                        help="Search terms (OR, case-insensitive, subject + body). Default: %(default)s")
    parser.add_argument("--exclude", nargs="*", default=[],
                        help="Exclusion phrases — emails containing any of these are dropped.")
    parser.add_argument("--project", default=CLI_DEFAULT_PROJECT,
                        help="Project label used in the summary heading (default: %(default)s)")
    parser.add_argument("--months", type=int, default=12,
                        help="Look back this many months (default: %(default)s)")
    parser.add_argument("--output", type=Path, default=Path("output"),
                        help="Output directory (default: %(default)s)")
    parser.add_argument("--include-folder", nargs="*", default=None, dest="include_folders_only",
                        help="Restrict search to folder paths containing any of these substrings.")
    parser.add_argument("--exclude-folder", nargs="*", default=[], dest="extra_exclude_folders",
                        help="Extra folder names to skip in addition to system folders.")
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

    def stdout_log(msg: str) -> None:
        print(msg, flush=True)

    try:
        result = scrape_outlook(
            keywords=args.keywords,
            exclude_keywords=args.exclude,
            project_label=args.project,
            months=args.months,
            output=args.output,
            include_deleted=args.include_deleted,
            include_junk=args.include_junk,
            extra_exclude_folders=args.extra_exclude_folders,
            include_folders_only=args.include_folders_only,
            skip_attachments=args.no_attachments,
            generate_summary=not args.no_summary,
            max_emails_for_summary=args.max_emails_for_summary,
            log=stdout_log,
        )
    except Exception as e:
        print(f"\nERROR: {e}", file=sys.stderr)
        return 1

    print(f"\nDone. {result['saved_count']}/{result['match_count']} emails saved.")
    if result["errors"]:
        print(f"({len(result['errors'])} errors — see above.)", file=sys.stderr)
    return 0 if result["saved_count"] > 0 or result["match_count"] == 0 else 2


if __name__ == "__main__":
    sys.exit(main())
