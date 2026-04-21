#!/usr/bin/env python3
"""
Scrape Outlook (Microsoft 365) mailbox for emails related to a project,
save copies as .eml files + attachments to a local folder, and generate a
Claude-powered summary report with action items.

Defaults are tuned for the Horizon Power / Project 307 brief, but every
parameter can be overridden on the CLI.

SETUP (one-time)
----------------
1. Register an Azure AD app in the cdienergy.com.au tenant:
     Azure Portal > App registrations > New registration
       - Name: anything (e.g. "Outlook scrape – dan.bailey")
       - Supported account types: "Accounts in this organizational directory only"
       - Redirect URI: (leave blank — device code flow)
     After creation:
       - Authentication > "Allow public client flows" = Yes > Save
       - API permissions > Add > Microsoft Graph > Delegated:
           * Mail.Read
           * User.Read
         Click "Grant admin consent" if your tenant requires it — otherwise
         the first sign-in will prompt for user consent, which is usually
         sufficient for delegated Mail.Read.
     Copy:
       - Application (client) ID   -> CLIENT_ID
       - Directory (tenant) ID     -> TENANT_ID

2. pip install -r requirements.txt

3. Copy .env.example to .env and fill in CLIENT_ID, TENANT_ID, ANTHROPIC_API_KEY.

4. Run:
     python scrape_outlook.py

   First run prints a URL + one-time code — open it in a browser, sign in
   as dan.bailey@cdienergy.com.au, approve the permissions. Subsequent runs
   use the cached refresh token silently (stored in ~/.outlook_scrape_token_cache.bin).

OUTPUT
------
  output/
    emails/<date>_<sender>_<subject>.eml    # full MIME copies
    attachments/<message_id>/<filename>      # per-message attachment folders
    index.json                               # machine-readable index
    SUMMARY.md                               # Claude-generated briefing
"""

import argparse
import base64
import json
import os
import re
import sys
import time
from datetime import datetime, timedelta, timezone
from pathlib import Path

import msal
import requests
from dotenv import load_dotenv

from summarize import render_docx, summarize_emails

load_dotenv()

GRAPH = "https://graph.microsoft.com/v1.0"
SCOPES = ["Mail.Read", "User.Read"]
DEFAULT_KEYWORDS = ["Horizon Power", "307", "Hossein", "remote communities"]
TOKEN_CACHE_PATH = Path.home() / ".outlook_scrape_token_cache.bin"


def get_token(client_id: str, tenant_id: str) -> str:
    cache = msal.SerializableTokenCache()
    if TOKEN_CACHE_PATH.exists():
        cache.deserialize(TOKEN_CACHE_PATH.read_text())

    app = msal.PublicClientApplication(
        client_id,
        authority=f"https://login.microsoftonline.com/{tenant_id}",
        token_cache=cache,
    )

    result = None
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])

    if not result:
        flow = app.initiate_device_flow(SCOPES)
        if "user_code" not in flow:
            raise RuntimeError(f"Device flow init failed: {json.dumps(flow, indent=2)}")
        print(flow["message"], flush=True)
        result = app.acquire_token_by_device_flow(flow)

    if "access_token" not in result:
        raise RuntimeError(f"Auth failed: {result.get('error_description', result)}")

    if cache.has_state_changed:
        TOKEN_CACHE_PATH.write_text(cache.serialize())
        try:
            TOKEN_CACHE_PATH.chmod(0o600)
        except OSError:
            pass

    return result["access_token"]


def graph_get(url: str, token: str, params: dict | None = None) -> dict:
    headers = {
        "Authorization": f"Bearer {token}",
        "ConsistencyLevel": "eventual",
    }
    for attempt in range(5):
        r = requests.get(url, headers=headers, params=params, timeout=60)
        if r.status_code == 429 or r.status_code >= 500:
            retry = int(r.headers.get("Retry-After", str(2 ** attempt)))
            print(f"  Graph {r.status_code}, sleeping {retry}s...", file=sys.stderr)
            time.sleep(retry)
            continue
        r.raise_for_status()
        return r.json()
    raise RuntimeError(f"Graph request failed after retries: {url}")


def search_messages(token: str, query: str) -> list[dict]:
    """Search the whole mailbox (all folders, incl. Sent) for one keyword."""
    url = f"{GRAPH}/me/messages"
    params = {
        "$search": f'"{query}"',
        "$select": (
            "id,subject,from,toRecipients,ccRecipients,receivedDateTime,"
            "hasAttachments,bodyPreview,parentFolderId,webLink"
        ),
        "$top": 100,
    }
    results = []
    while url:
        data = graph_get(url, token, params=params)
        results.extend(data.get("value", []))
        url = data.get("@odata.nextLink")
        params = None  # nextLink already carries query string
    return results


def get_message_body(token: str, msg_id: str) -> dict:
    return graph_get(
        f"{GRAPH}/me/messages/{msg_id}",
        token,
        params={"$select": "id,subject,from,toRecipients,receivedDateTime,body,bodyPreview,webLink"},
    )


def download_mime(token: str, msg_id: str, dest: Path) -> None:
    headers = {"Authorization": f"Bearer {token}"}
    for attempt in range(5):
        r = requests.get(f"{GRAPH}/me/messages/{msg_id}/$value", headers=headers, timeout=120)
        if r.status_code == 429 or r.status_code >= 500:
            retry = int(r.headers.get("Retry-After", str(2 ** attempt)))
            time.sleep(retry)
            continue
        r.raise_for_status()
        dest.write_bytes(r.content)
        return
    raise RuntimeError(f"MIME download failed: {msg_id}")


def download_attachments(token: str, msg_id: str, dest_dir: Path) -> list[str]:
    data = graph_get(f"{GRAPH}/me/messages/{msg_id}/attachments", token)
    saved = []
    for att in data.get("value", []):
        if att.get("@odata.type") != "#microsoft.graph.fileAttachment":
            continue  # skip itemAttachment / referenceAttachment
        name = sanitize_filename(att.get("name", "attachment"))
        b64 = att.get("contentBytes")
        if not b64:
            continue
        dest_dir.mkdir(parents=True, exist_ok=True)
        path = dest_dir / name
        path.write_bytes(base64.b64decode(b64))
        saved.append(str(path))
    return saved


def sanitize_filename(name: str, max_len: int = 120) -> str:
    cleaned = re.sub(r"[^\w\-.() ]+", "_", name).strip("._ ")
    return (cleaned or "unnamed")[:max_len]


def strip_html(html: str) -> str:
    text = re.sub(r"<style[^>]*>.*?</style>", "", html, flags=re.DOTALL | re.IGNORECASE)
    text = re.sub(r"<script[^>]*>.*?</script>", "", text, flags=re.DOTALL | re.IGNORECASE)
    text = re.sub(r"<[^>]+>", " ", text)
    replacements = {"&nbsp;": " ", "&amp;": "&", "&lt;": "<", "&gt;": ">", "&quot;": '"', "&#39;": "'"}
    for k, v in replacements.items():
        text = text.replace(k, v)
    return re.sub(r"\s+", " ", text).strip()


def normalize_for_summary(m: dict) -> dict:
    """Convert a Graph message dict to the flat shape expected by summarize_emails."""
    body = m.get("body") or {}
    content = body.get("content", "") if isinstance(body, dict) else ""
    if isinstance(body, dict) and body.get("contentType") == "html":
        content = strip_html(content)
    return {
        "from": m.get("from", {}).get("emailAddress", {}).get("address", "unknown"),
        "to": ", ".join(
            r.get("emailAddress", {}).get("address", "")
            for r in m.get("toRecipients", [])
        ),
        "date": m.get("receivedDateTime", ""),
        "subject": m.get("subject", "(no subject)"),
        "body": content,
    }


def main() -> int:
    parser = argparse.ArgumentParser(
        description=__doc__,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("--keywords", nargs="+", default=DEFAULT_KEYWORDS,
                        help="Search terms (OR). Default: %(default)s")
    parser.add_argument("--exclude", nargs="*", default=[],
                        help="Exclusion phrases — emails containing any of these are dropped.")
    parser.add_argument("--project", default="Project 307 – Horizon Power Remote Communities",
                        help="Project label for the summary heading.")
    parser.add_argument("--months", type=int, default=12,
                        help="Look back this many months (default: %(default)s)")
    parser.add_argument("--output", type=Path, default=Path("output"),
                        help="Output directory (default: %(default)s)")
    parser.add_argument("--no-attachments", action="store_true",
                        help="Skip downloading attachments")
    parser.add_argument("--no-summary", action="store_true",
                        help="Skip Claude summary generation")
    parser.add_argument("--max-emails-for-summary", type=int, default=200,
                        help="Cap on emails fed into the summary prompt (default: %(default)s)")
    args = parser.parse_args()

    client_id = os.environ.get("CLIENT_ID")
    tenant_id = os.environ.get("TENANT_ID")
    if not client_id or not tenant_id:
        print("Missing CLIENT_ID or TENANT_ID in environment. See .env.example.", file=sys.stderr)
        return 1

    print(f"Authenticating via device code flow (tenant: {tenant_id})...")
    token = get_token(client_id, tenant_id)
    me = graph_get(f"{GRAPH}/me", token)
    print(f"Signed in as: {me.get('userPrincipalName') or me.get('mail')}")

    cutoff = datetime.now(timezone.utc) - timedelta(days=30 * args.months)
    print(f"\nSearching emails received since {cutoff.date().isoformat()} for: {args.keywords}")

    by_id: dict[str, dict] = {}
    for kw in args.keywords:
        print(f"  '{kw}' ...", end="", flush=True)
        hits = search_messages(token, kw)
        new = 0
        for m in hits:
            recv = m.get("receivedDateTime")
            if not recv:
                continue
            dt = datetime.fromisoformat(recv.replace("Z", "+00:00"))
            if dt < cutoff:
                continue
            if m["id"] not in by_id:
                by_id[m["id"]] = m
                new += 1
        print(f" {len(hits)} hits, {new} new in range")

    print(f"\nTotal unique emails in date range: {len(by_id)}")
    if not by_id:
        print("Nothing to do.")
        return 0

    messages = sorted(by_id.values(), key=lambda m: m["receivedDateTime"])

    emails_dir = args.output / "emails"
    attachments_dir = args.output / "attachments"
    emails_dir.mkdir(parents=True, exist_ok=True)

    print(f"\nDownloading {len(messages)} emails to {emails_dir}/ ...")
    index = []
    for i, m in enumerate(messages, 1):
        recv = m["receivedDateTime"][:10]
        sender = m.get("from", {}).get("emailAddress", {}).get("address", "unknown")
        subject = m.get("subject", "no_subject")
        filename = sanitize_filename(f"{recv}__{sender}__{subject}") + ".eml"
        eml_path = emails_dir / filename
        try:
            download_mime(token, m["id"], eml_path)
        except Exception as e:
            print(f"  [{i}/{len(messages)}] FAILED: {m['id']}: {e}", file=sys.stderr)
            continue

        atts: list[str] = []
        if m.get("hasAttachments") and not args.no_attachments:
            try:
                atts = download_attachments(token, m["id"], attachments_dir / sanitize_filename(m["id"]))
            except Exception as e:
                print(f"  [{i}/{len(messages)}] attachment download failed: {e}", file=sys.stderr)

        index.append({
            "id": m["id"],
            "date": m["receivedDateTime"],
            "from": sender,
            "to": [r.get("emailAddress", {}).get("address") for r in m.get("toRecipients", [])],
            "subject": subject,
            "has_attachments": m.get("hasAttachments", False),
            "eml": str(eml_path.relative_to(args.output)),
            "attachments": [str(Path(a).relative_to(args.output)) for a in atts],
            "web_link": m.get("webLink"),
        })
        if i % 10 == 0 or i == len(messages):
            print(f"  [{i}/{len(messages)}] downloaded")

    (args.output / "index.json").write_text(json.dumps(index, indent=2))
    print(f"\nIndex: {args.output / 'index.json'}")

    if args.no_summary:
        return 0
    if not os.environ.get("ANTHROPIC_API_KEY"):
        print("\nANTHROPIC_API_KEY not set — skipping summary.", file=sys.stderr)
        return 0

    sample = messages[-args.max_emails_for_summary:] if len(messages) > args.max_emails_for_summary else messages
    print(f"\nFetching bodies for summary ({len(sample)} emails)...")
    bodied: list[dict] = []
    for i, m in enumerate(sample, 1):
        try:
            bodied.append(get_message_body(token, m["id"]))
        except Exception as e:
            print(f"  body fetch failed for {m['id']}: {e}", file=sys.stderr)
        if i % 20 == 0 or i == len(sample):
            print(f"  [{i}/{len(sample)}] bodies fetched")

    # Apply exclusion filter
    if args.exclude:
        ex_lower = [e.lower() for e in args.exclude]
        def _excluded(m):
            body_obj = m.get("body") or {}
            content = body_obj.get("content", "") if isinstance(body_obj, dict) else ""
            hay = ((m.get("subject") or "") + "\n" + content).lower()
            return any(e in hay for e in ex_lower)
        before = len(bodied)
        bodied = [m for m in bodied if not _excluded(m)]
        print(f"Excluded {before - len(bodied)} emails via --exclude.")

    print("\nGenerating summary with Claude Opus 4.7...")
    summary = summarize_emails(
        [normalize_for_summary(m) for m in bodied],
        args.keywords,
        project_label=args.project,
    )
    summary_md = args.output / "SUMMARY.md"
    summary_md.write_text(summary, encoding="utf-8")
    print(f"Summary (markdown): {summary_md}")

    summary_docx = args.output / "SUMMARY.docx"
    if render_docx(summary, summary_docx):
        print(f"Summary (Word):     {summary_docx}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
