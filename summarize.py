"""Shared Claude Opus 4.7 summarization for scraped Outlook emails.

Both scrape_outlook.py (Graph/Azure) and scrape_outlook_local.py (Windows COM)
call summarize_emails() with a list of normalized email dicts.
"""

import sys
from pathlib import Path


EMAIL_SEPARATOR = "\n\n--- EMAIL ---\n\n"


def _format_email(e: dict, body_char_limit: int = 3000) -> str:
    body = (e.get("body") or "").strip()[:body_char_limit]
    return (
        f"From: {e.get('from', 'unknown')}\n"
        f"To: {e.get('to', '')}\n"
        f"Date: {e.get('date', '')}\n"
        f"Subject: {e.get('subject', '(no subject)')}\n\n"
        f"{body}"
    )


def summarize_emails(emails: list[dict], keywords: list[str]) -> str:
    """Generate a markdown briefing.

    Each email dict must have: from, to, date, subject, body.
    """
    import anthropic

    client = anthropic.Anthropic()
    corpus = EMAIL_SEPARATOR.join(_format_email(e) for e in emails)

    system = [
        {
            "type": "text",
            "text": (
                "You are a project analyst producing an action-oriented briefing "
                "from a project email archive. Your output goes to a project "
                "manager who needs to catch up quickly and identify what to do "
                "next. Be specific. Cite sender + date for every finding. Do not "
                "invent facts or assign owners that were not named in the emails."
            ),
        },
        {
            "type": "text",
            "text": f"EMAIL CORPUS (filtered on keywords: {', '.join(keywords)}):\n\n{corpus}",
            "cache_control": {"type": "ephemeral"},
        },
    ]

    prompt = (
        "Produce a markdown report titled "
        "`# Project 307 – Horizon Power Remote Communities – Email Briefing`.\n\n"
        "Include these sections in order:\n\n"
        "## Overview\n"
        "One short paragraph: date range covered, email count, main threads/topics, "
        "key people involved.\n\n"
        "## Key Decisions\n"
        "Bulleted list. For each: what was decided, who decided, date, source "
        "(sender + date + subject).\n\n"
        "## Action Items\n"
        "Markdown table with columns: `Owner | Action | Due Date | Status | Source`. "
        "Include actions explicitly assigned AND actions clearly expected but not "
        "formally assigned. Use `Unassigned` when the owner is unclear. "
        "Status is one of `Open`, `Done`, `Blocked`, `Unclear` based on the thread. "
        "Source format: `<sender>, <YYYY-MM-DD>, \"<subject>\"`.\n\n"
        "## Open Risks & Blockers\n"
        "Bulleted list: risk/blocker, potential impact, suggested next step, source.\n\n"
        "## Key Contacts\n"
        "Table: `Name | Email | Role | Relevance`. Role only if mentioned in emails. "
        "Include Hossein prominently if present.\n\n"
        "## Outstanding Questions\n"
        "Bulleted list of questions asked in emails that do not appear to have "
        "been answered.\n\n"
        "Keep it tight. Prioritise specificity over breadth. If a section has no "
        "material content, write `_None identified._` rather than padding."
    )

    resp = client.messages.create(
        model="claude-opus-4-7",
        max_tokens=16000,
        thinking={"type": "adaptive"},
        output_config={"effort": "high"},
        system=system,
        messages=[{"role": "user", "content": prompt}],
    )

    u = resp.usage
    print(
        f"  Claude usage: input={u.input_tokens}, "
        f"cache_read={u.cache_read_input_tokens}, "
        f"cache_write={u.cache_creation_input_tokens}, "
        f"output={u.output_tokens}",
        file=sys.stderr,
    )
    return "".join(b.text for b in resp.content if b.type == "text")


def render_docx(markdown_text: str, dest: Path) -> bool:
    """Convert markdown briefing to a Word .docx via Pandoc.

    Returns True on success, False if pypandoc/pandoc isn't available
    (caller can keep the .md and move on).
    """
    try:
        import pypandoc
    except ImportError:
        print(
            "  pypandoc not installed — skipping .docx (pip install pypandoc-binary).",
            file=sys.stderr,
        )
        return False

    try:
        pypandoc.convert_text(
            markdown_text,
            to="docx",
            format="gfm",  # GitHub-Flavored Markdown — handles pipe tables
            outputfile=str(dest),
            extra_args=["--standalone"],
        )
        return True
    except OSError as e:
        # pypandoc raises OSError when the pandoc binary itself is missing.
        print(
            f"  Pandoc binary not found: {e}\n"
            "  Install with: pip install pypandoc-binary",
            file=sys.stderr,
        )
        return False
    except Exception as e:
        print(f"  .docx conversion failed: {e}", file=sys.stderr)
        return False
