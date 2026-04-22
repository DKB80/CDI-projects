"""Scrape Outlook for Mac via AppleScript / JXA.

Public API matches scrape_outlook_local.py so the GUI can swap backends
via outlook_backend.py without any other changes.

Requires macOS with Outlook for Mac running and signed in. The first
time the tool drives Outlook, macOS shows a one-off system prompt asking
you to grant "Automation" permission — click Allow.
"""

from __future__ import annotations

import json
import os
import re
import subprocess
import sys
from datetime import datetime, timedelta
from pathlib import Path
from typing import Callable


Logger = Callable[[str], None]


def _noop(msg: str) -> None:
    pass


# Same skip list + system-sender heuristic as the Windows backend.
SKIP_FOLDERS_DEFAULT = {
    "Deleted Items", "Junk Email", "Junk", "Junk E-mail",
    "Sync Issues", "Conversation History", "RSS Feeds",
    "Conflicts", "Local Failures", "Server Failures",
    "Recoverable Items",
}

_SYSTEM_LOCAL_PARTS = frozenset({
    "no-reply", "noreply", "no_reply", "nr",
    "donotreply", "do-not-reply", "do_not_reply",
    "mailer-daemon", "mailerdaemon", "mailer_daemon",
    "postmaster",
    "bounce", "bounces", "bounce-notice",
    "auto-reply", "autoreply", "automated",
    "notifications", "notification", "notify",
    "alerts", "alert", "alarm", "alarms",
    "news", "newsletter", "digest",
    "system", "sys", "root", "daemon",
    "unsubscribe", "subscribe",
})
_SYSTEM_LOCAL_SUBSTRINGS = (
    "noreply", "no-reply", "no_reply",
    "donotreply", "do-not-reply", "do_not_reply",
    "mailer-daemon", "mailerdaemon", "mailer_daemon",
    "-bounce", "-bounces",
)


def _is_system_sender_email(email: str) -> bool:
    if not email or "@" not in email:
        return False
    local = email.split("@", 1)[0].lower().strip()
    if local in _SYSTEM_LOCAL_PARTS:
        return True
    return any(sub in local for sub in _SYSTEM_LOCAL_SUBSTRINGS)


def sanitize_filename(name: str, max_len: int = 120) -> str:
    cleaned = re.sub(r"[\x00-\x1f<>:\"/\\|?*]+", "_", name or "")
    cleaned = re.sub(r"\s+", " ", cleaned).strip("._ ")
    return (cleaned or "unnamed")[:max_len]


def _run_jxa(script: str, timeout: int = 600) -> str:
    """Execute a JavaScript for Automation script via osascript and
    return stdout. Raises RuntimeError on non-zero exit."""
    try:
        result = subprocess.run(
            ["osascript", "-l", "JavaScript", "-"],
            input=script,
            text=True,
            capture_output=True,
            timeout=timeout,
        )
    except FileNotFoundError as e:
        raise RuntimeError("osascript not found — this backend only runs on macOS.") from e
    if result.returncode != 0:
        raise RuntimeError(f"Outlook automation failed:\n{result.stderr.strip()}")
    return result.stdout.strip()


# ---------- JXA templates ----------

_JXA_PROBE = """
const outlook = Application('Microsoft Outlook');
outlook.includeStandardAdditions = true;
try {
    const running = outlook.running();
    JSON.stringify({ok: true, running: running});
} catch (e) {
    JSON.stringify({ok: false, error: String(e)});
}
"""


_JXA_LIST_ACCOUNTS = """
const outlook = Application('Microsoft Outlook');
const accounts = [];
try {
    outlook.exchangeAccounts().forEach(a => {
        try { accounts.push(String(a.emailAddresses().address())); } catch (_) {}
    });
} catch (_) {}
try {
    outlook.imapAccounts().forEach(a => {
        try { accounts.push(String(a.emailAddresses().address())); } catch (_) {}
    });
} catch (_) {}
try {
    outlook.pop3Accounts().forEach(a => {
        try { accounts.push(String(a.emailAddresses().address())); } catch (_) {}
    });
} catch (_) {}
JSON.stringify(accounts);
"""


_JXA_SCAN = """
const params = PARAMS_JSON;
const outlook = Application('Microsoft Outlook');
const cutoffMs = new Date().getTime() - (params.months * 30 * 86400 * 1000);
const cutoff = new Date(cutoffMs);

const kwLower = params.keywords.map(k => k.toLowerCase());
const exLower = params.excludes.map(k => k.toLowerCase());
const fromLower = params.from_senders.map(k => k.toLowerCase());
const skipSet = new Set(params.skip_folders);
const onlyFolders = params.include_folders_only;

function walk(folder, path) {
    const name = (() => { try { return folder.name(); } catch (_) { return ""; } })();
    if (skipSet.has(name)) return [];
    const thisPath = path ? path + "/" + name : name;
    const out = [{folder: folder, path: thisPath}];
    try {
        const kids = folder.mailFolders();
        for (let i = 0; i < kids.length; i++) {
            out.push(...walk(kids[i], thisPath));
        }
    } catch (_) {}
    return out;
}

const matches = [];
const seen = new Set();
let totalScanned = 0;

// Walk every top-level mail folder of every account
let roots = [];
try { roots = outlook.mailFolders(); } catch (_) {}

for (let i = 0; i < roots.length; i++) {
    const folders = walk(roots[i], "");
    for (let j = 0; j < folders.length; j++) {
        const {folder, path} = folders[j];
        if (onlyFolders && onlyFolders.length) {
            const low = path.toLowerCase();
            if (!onlyFolders.some(n => low.indexOf(n.toLowerCase()) >= 0)) continue;
        }
        let msgs;
        try {
            msgs = folder.messages.whose({timeReceived: {_greaterThan: cutoff}})();
        } catch (_) { continue; }
        for (let k = 0; k < msgs.length; k++) {
            totalScanned++;
            if (totalScanned > params.max_scan) break;
            const m = msgs[k];
            let subject = "", body = "", senderEmail = "", senderName = "", id = "", received = "";
            try { subject = String(m.subject()) || ""; } catch (_) {}
            try { body = String(m.plainTextContent() || ""); } catch (_) {}
            try {
                const s = m.sender();
                if (s) {
                    senderEmail = String(s.address() || "");
                    senderName = String(s.name() || "");
                }
            } catch (_) {}
            try { id = String(m.id()); } catch (_) {}
            try { received = m.timeReceived().toISOString(); } catch (_) {}

            if (seen.has(id)) continue;
            const haySender = (senderEmail + " " + senderName).toLowerCase();
            const hayBody = (subject + "\\n" + body).toLowerCase();
            if (fromLower.length && !fromLower.some(f => haySender.indexOf(f) >= 0)) continue;
            const kwOk = kwLower.length === 0 || kwLower.some(k => hayBody.indexOf(k) >= 0);
            const exOk = !exLower.some(e => hayBody.indexOf(e) >= 0);
            if (!kwOk || !exOk) continue;

            seen.add(id);
            let hasAtt = false;
            try { hasAtt = (m.attachments().length > 0); } catch (_) {}
            matches.push({
                id: id,
                subject: subject,
                body_preview: body.substring(0, 400).replace(/\\s+/g, " "),
                body_full: body,
                from_email: senderEmail,
                from_name: senderName,
                date: received,
                folder: path,
                has_attachments: hasAtt,
            });
        }
        if (totalScanned > params.max_scan) break;
    }
    if (totalScanned > params.max_scan) break;
}

JSON.stringify({
    scanned: totalScanned,
    matches: matches,
});
"""


_JXA_SENDER_SCAN = """
const params = PARAMS_JSON;
const outlook = Application('Microsoft Outlook');
const cutoffMs = new Date().getTime() - (params.months * 30 * 86400 * 1000);
const cutoff = new Date(cutoffMs);
const skipSet = new Set(params.skip_folders);

function walk(folder, path) {
    const name = (() => { try { return folder.name(); } catch (_) { return ""; } })();
    if (skipSet.has(name)) return [];
    const thisPath = path ? path + "/" + name : name;
    const out = [{folder: folder, path: thisPath}];
    try {
        const kids = folder.mailFolders();
        for (let i = 0; i < kids.length; i++) {
            out.push(...walk(kids[i], thisPath));
        }
    } catch (_) {}
    return out;
}

const counts = {};
let scannedItems = 0;
let scannedFolders = 0;
const cap = params.max_items;

let roots = [];
try { roots = outlook.mailFolders(); } catch (_) {}

outer:
for (let i = 0; i < roots.length; i++) {
    const folders = walk(roots[i], "");
    for (let j = 0; j < folders.length; j++) {
        const {folder} = folders[j];
        let msgs;
        try { msgs = folder.messages.whose({timeReceived: {_greaterThan: cutoff}})(); } catch (_) { continue; }
        scannedFolders++;
        for (let k = 0; k < msgs.length; k++) {
            if (scannedItems >= cap) break outer;
            const m = msgs[k];
            try {
                const s = m.sender();
                if (s) {
                    const email = (s.address() || "").trim();
                    const name = (s.name() || "").trim();
                    const key = email.toLowerCase();
                    if (key) {
                        if (!counts[key]) counts[key] = {email: email, name: name, count: 0};
                        counts[key].count += 1;
                        if (name && !counts[key].name) counts[key].name = name;
                    }
                }
            } catch (_) {}
            scannedItems++;
        }
    }
}

JSON.stringify({
    scanned_folders: scannedFolders,
    scanned_items: scannedItems,
    senders: Object.keys(counts).map(k => counts[k]),
});
"""


_JXA_SAVE_ATTACHMENTS = """
const params = PARAMS_JSON;
const outlook = Application('Microsoft Outlook');
const app = Application.currentApplication();
app.includeStandardAdditions = true;

const results = [];
for (let i = 0; i < params.items.length; i++) {
    const item = params.items[i];
    try {
        const msg = outlook.messages.byId(item.id);
        const attDir = item.dest_dir;
        app.doShellScript("mkdir -p " + quoted(attDir));
        const atts = msg.attachments();
        for (let a = 0; a < atts.length; a++) {
            try {
                const aname = atts[a].name();
                const safeName = aname.replace(/[\\/]/g, "_");
                const dest = attDir + "/" + safeName;
                atts[a].save({in: Path(dest)});
                results.push({id: item.id, path: dest, ok: true});
            } catch (e) {
                results.push({id: item.id, ok: false, error: String(e)});
            }
        }
    } catch (e) {
        results.push({id: item.id, ok: false, error: String(e)});
    }
}

function quoted(s) {
    return "'" + String(s).replace(/'/g, "'\\\\''") + "'";
}

JSON.stringify(results);
"""


# ---------- Public API ----------

def list_recent_senders(
    namespace=None,  # accepted for signature compat with Windows backend; ignored.
    months_back: int = 12,
    max_items: int = 1000,
    log: Logger | None = None,
    progress_cb: Callable[[int, int], bool] | None = None,
) -> list[dict]:
    log = log or _noop
    script = _JXA_SENDER_SCAN.replace("PARAMS_JSON", json.dumps({
        "months": months_back,
        "max_items": max_items,
        "skip_folders": list(SKIP_FOLDERS_DEFAULT | {"Sent Items"}),
    }))
    log(f"Scanning recent senders via AppleScript (cap {max_items}) ...")
    try:
        out = _run_jxa(script, timeout=600)
        data = json.loads(out)
    except Exception as e:
        log(f"Sender scan failed: {e}")
        return []

    senders = data.get("senders", [])
    filtered: list[dict] = []
    skipped = 0
    for s in senders:
        if _is_system_sender_email(s.get("email", "")):
            skipped += 1
            continue
        filtered.append(s)

    filtered.sort(key=lambda d: (-d.get("count", 0), d.get("email", "")))
    log(f"Sender scan done: {data.get('scanned_folders', 0)} folder(s), "
        f"{data.get('scanned_items', 0)} email(s), {len(filtered)} unique senders, "
        f"skipped {skipped} automated/no-reply")
    return filtered


def scrape_outlook(
    *,
    keywords: list[str] | None = None,
    exclude_keywords: list[str] | None = None,
    from_senders: list[str] | None = None,
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
    from summarize import render_docx, summarize_emails

    log = log or _noop
    keywords = keywords or []
    from_senders = from_senders or []
    if not keywords and not from_senders:
        raise ValueError("Supply at least one keyword or one sender filter.")

    skip = set(SKIP_FOLDERS_DEFAULT)
    if include_deleted:
        skip.discard("Deleted Items")
    if include_junk:
        for j in ("Junk Email", "Junk", "Junk E-mail"):
            skip.discard(j)
    if extra_exclude_folders:
        skip.update(extra_exclude_folders)

    output = Path(output).expanduser().absolute()
    emails_dir = output / "emails"
    attachments_dir = output / "attachments"
    emails_dir.mkdir(parents=True, exist_ok=True)
    log(f"Output directory: {output}")

    result: dict = {
        "match_count": 0,
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

    log("Connecting to Outlook for Mac...")
    try:
        probe = _run_jxa(_JXA_PROBE, timeout=30)
        probe_data = json.loads(probe)
        if not probe_data.get("ok"):
            raise RuntimeError(probe_data.get("error", "unknown error"))
        if not probe_data.get("running"):
            raise RuntimeError("Outlook is not running. Please open Outlook and try again.")
    except Exception as e:
        raise RuntimeError(f"Failed to connect to Outlook for Mac: {e}") from e

    log(f"Searching for {keywords} since {months} months ago "
        f"(excluding: {exclude_keywords or []}, from: {from_senders or []})")
    log(f"Skipping folders: {sorted(skip)}")

    scan_params = {
        "keywords": keywords,
        "excludes": exclude_keywords or [],
        "from_senders": from_senders,
        "months": months,
        "skip_folders": sorted(skip),
        "include_folders_only": include_folders_only or [],
        "max_scan": 50000,  # generous ceiling to avoid infinite loops
    }
    log("Starting AppleScript scan — this can take several minutes on a large mailbox...")
    try:
        raw = _run_jxa(_JXA_SCAN.replace("PARAMS_JSON", json.dumps(scan_params)),
                       timeout=1800)
        scan_data = json.loads(raw)
    except Exception as e:
        result["errors"].append(f"Scan failed: {e}")
        log(f"Scan failed: {e}")
        return result

    matches = scan_data.get("matches", [])
    matches.sort(key=lambda m: m.get("date", ""))
    result["match_count"] = len(matches)
    log(f"Scanned {scan_data.get('scanned', 0)} email(s), {len(matches)} matches.")

    if not matches:
        return result

    # Preview callback (GUI hook)
    if preview_callback is not None:
        preview_payload = [
            {
                "entry_id": m["id"],
                "subject": m.get("subject", ""),
                "from": m.get("from_email") or m.get("from_name") or "unknown",
                "date": m.get("date", ""),
                "body_preview": m.get("body_preview", ""),
                "has_attachments": m.get("has_attachments", False),
            }
            for m in matches
        ]
        approved = preview_callback(preview_payload)
        if approved is None:
            result["errors"].append("Cancelled by user at preview step.")
            return result
        approved_set = set(approved)
        matches = [m for m in matches if m["id"] in approved_set]
        result["approved_count"] = len(matches)
        if not matches:
            return result
        log(f"User approved {len(matches)}/{result['match_count']} matches.")

    # Write .eml files
    index = []
    save_jobs = []  # for batched attachment saves
    log(f"Saving emails to {emails_dir} ...")
    for i, m in enumerate(matches, 1):
        date_str = (m.get("date") or "")[:10] or "no-date"
        sender = m.get("from_email") or m.get("from_name") or "unknown"
        subject = m.get("subject") or "no_subject"
        base_name = sanitize_filename(f"{date_str}__{sender}__{subject}")
        eml_path: Path | None = None
        if save_msg_files:
            eml_path = emails_dir / (base_name + ".eml")
            if eml_path.exists():
                eml_path = emails_dir / (base_name + "__" + m["id"][-8:] + ".eml")
            _write_eml_file(eml_path, m)

        atts: list[str] = []
        if not skip_attachments and m.get("has_attachments"):
            att_dir = attachments_dir / sanitize_filename(f"{date_str}__{subject}__{m['id'][-8:]}")
            save_jobs.append({"id": m["id"], "dest_dir": str(att_dir)})

        index.append({
            "entry_id": m["id"],
            "date": m.get("date", ""),
            "from": sender,
            "to": "",
            "cc": "",
            "subject": subject,
            "msg": str(eml_path) if eml_path else None,
            "attachments": atts,  # filled in after batch save
            "body": m.get("body_full", ""),
        })
        if i % 20 == 0 or i == len(matches):
            log(f"  [{i}/{len(matches)}] saved")

    # Batch-save attachments via a single AppleScript round-trip.
    if save_jobs:
        log(f"Extracting attachments for {len(save_jobs)} email(s) ...")
        try:
            raw = _run_jxa(
                _JXA_SAVE_ATTACHMENTS.replace("PARAMS_JSON",
                                              json.dumps({"items": save_jobs})),
                timeout=1800,
            )
            att_results = json.loads(raw) if raw else []
        except Exception as e:
            log(f"Attachment save failed: {e}")
            att_results = []
        by_id: dict[str, list[str]] = {}
        for r in att_results:
            if r.get("ok") and r.get("path"):
                by_id.setdefault(r["id"], []).append(r["path"])
        for entry in index:
            entry["attachments"] = by_id.get(entry["entry_id"], [])

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

    sample = index[-max_emails_for_summary:] if len(index) > max_emails_for_summary else index
    summary: str | None = None
    summary_type: str | None = None
    api_key = os.environ.get("ANTHROPIC_API_KEY")

    if api_key:
        try:
            log(f"Generating AI summary with Claude Opus 4.7 ({len(sample)} emails)...")
            summary = summarize_emails(sample, keywords, project_label=project_label)
            summary_type = "ai"
        except Exception as e:
            log(f"AI summary failed ({e}); falling back to offline briefing.")
            result["errors"].append(f"AI summary failed: {e}")
    else:
        log("No ANTHROPIC_API_KEY — generating offline briefing (no AI analysis).")

    if summary is None:
        from cdi_gui.offline_summary import build_offline_briefing
        summary = build_offline_briefing(
            sample, project_label, keywords,
            exclude_keywords=exclude_keywords, months=months,
        )
        summary_type = "offline"

    result["summary_type"] = summary_type
    result["summary_markdown"] = summary

    summary_md = output / "SUMMARY.md"
    summary_md.write_text(summary, encoding="utf-8")
    result["summary_md_path"] = str(summary_md)
    log(f"Summary (markdown, {summary_type}): {summary_md}")

    summary_docx = output / "SUMMARY.docx"
    if render_docx(summary, summary_docx, project=project_label,
                   period=f"last {months} months"):
        result["summary_docx_path"] = str(summary_docx)
        log(f"Summary (Word, {summary_type}): {summary_docx}")

    return result


def _write_eml_file(path: Path, m: dict) -> None:
    """Write a minimal RFC 822 .eml file from message metadata."""
    from email.message import EmailMessage
    from email.utils import formatdate

    eml = EmailMessage()
    name = m.get("from_name") or ""
    addr = m.get("from_email") or ""
    if name and addr:
        eml["From"] = f'"{name}" <{addr}>'
    elif addr:
        eml["From"] = addr
    else:
        eml["From"] = name or "unknown"

    eml["Subject"] = m.get("subject") or ""

    date = m.get("date") or ""
    if date:
        eml["Date"] = date
    else:
        eml["Date"] = formatdate()

    eml.set_content(m.get("body_full") or "")
    path.write_bytes(bytes(eml))
