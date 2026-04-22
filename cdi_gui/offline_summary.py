"""Offline briefing generator — produces a CDI-shaped markdown briefing
from email metadata alone, no LLM call required.

Used as a fallback by scrape_outlook() when the Anthropic API is
unreachable (no key, no internet, or the call errors out). Output has
exactly the same H1 / ## section structure the CDI docx renderer
expects, so the resulting Word doc looks the same except the AI-only
sections say 'AI analysis required — re-run with internet access'.
"""

from __future__ import annotations

from collections import Counter
from datetime import datetime
import re

_AI_REQUIRED = (
    "_AI analysis required — re-run with internet access (and a valid "
    "Anthropic API key) to populate this section._"
)


def _domain(email: str) -> str:
    if "@" in email:
        return email.split("@", 1)[1].lower()
    return ""


def _display_from_sender(sender: str) -> tuple[str, str]:
    """Best-effort name + email split. SenderEmailAddress from Outlook
    is usually just the address; the Claude prompt normally has 'from'
    as the address too, and 'name' isn't stored in the index."""
    sender = (sender or "").strip()
    # 'Name <email@x>' form
    m = re.match(r"(.+?)\s*<([^>]+)>\s*$", sender)
    if m:
        return m.group(1).strip().strip('"'), m.group(2).strip()
    if "@" in sender:
        return "", sender
    return sender, ""


def build_offline_briefing(
    emails: list[dict],
    project_label: str,
    keywords: list[str],
    exclude_keywords: list[str] | None = None,
    months: int = 12,
) -> str:
    """Return a markdown briefing that slots into the CDI docx renderer."""
    if not emails:
        return (f"# {project_label} – Email Briefing\n\n"
                f"## Overview\n_No emails were saved — nothing to summarise._\n")

    dates = sorted(e.get("date", "") for e in emails if e.get("date"))
    start = dates[0][:10] if dates else "?"
    end = dates[-1][:10] if dates else "?"

    senders = Counter(e.get("from", "unknown") for e in emails if e.get("from"))
    subjects = [e.get("subject", "") for e in emails if e.get("subject")]
    attach_count = sum(len(e.get("attachments") or []) for e in emails)

    lines: list[str] = []
    add = lines.append

    add(f"# {project_label} – Email Briefing")
    add("")

    # Overview
    add("## Overview")
    kw_str = ", ".join(keywords) if keywords else "—"
    ex_str = ", ".join(exclude_keywords or []) or "none"
    add(
        f"Offline briefing generated {datetime.now().strftime('%d %b %Y %H:%M')} "
        f"with no AI analysis. Covers **{len(emails)}** saved email(s) received "
        f"between **{start}** and **{end}** (search window: last {months} months). "
        f"Filter keywords: _{kw_str}_. Exclusions: _{ex_str}_. "
        f"Attachments captured: {attach_count}. Unique senders: {len(senders)}."
    )
    add("")
    if senders:
        add("**Most active senders in this set:**")
        for sender, count in senders.most_common(5):
            add(f"- {sender} — {count} email(s)")
        add("")
    add(
        "> This is an offline briefing. The sections below marked "
        "_AI analysis required_ can only be produced by Claude. Re-run "
        "the tool with internet access and a valid API key to get the "
        "full briefing with decisions, action items, risks, and questions."
    )
    add("")

    # Key Decisions — AI-only
    add("## Key Decisions")
    add(_AI_REQUIRED)
    add("")

    # Action Items — AI-only, but we emit a single-row placeholder table so
    # the CDI docx renderer still draws a table + Status chip.
    add("## Action Items")
    add("| Owner | Action | Due Date | Status | Source |")
    add("|---|---|---|---|---|")
    add("| _AI analysis required_ | _Re-run with internet access to extract action items_ | _—_ | Unclear | _Offline mode_ |")
    add("")

    # Open Risks — AI-only
    add("## Open Risks & Blockers")
    add(_AI_REQUIRED)
    add("")

    # Key Contacts — we can do this deterministically
    add("## Key Contacts")
    add("| Name | Email | Role | Relevance |")
    add("|---|---|---|---|")
    for sender, count in senders.most_common():
        name, email = _display_from_sender(sender)
        add(f"| {name or '—'} | {email or sender} | — | {count} email(s) in this set |")
    add("")

    # Outstanding Questions — AI-only
    add("## Outstanding Questions")
    add(_AI_REQUIRED)
    add("")

    return "\n".join(lines)
