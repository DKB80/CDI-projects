"""CDI Energy branded Word renderer for the Claude briefing markdown.

Implements the CDI briefing format spec: landscape A4, navy / accent blue
palette, cover block, zebra-striped tables, Status chips, risk cards, and
contacts grouped by email domain.

Called from summarize.render_docx() which handles fallback logic.
"""

from __future__ import annotations

import re
from collections import OrderedDict
from pathlib import Path

from docx import Document
from docx.enum.section import WD_ORIENTATION, WD_SECTION
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor


# ---------- Palette ----------

BRAND_NAVY = RGBColor(0x1F, 0x38, 0x64)
ACCENT_BLUE = RGBColor(0x2E, 0x75, 0xB6)
GREY_MID = RGBColor(0x59, 0x59, 0x59)
GREY_DARK = RGBColor(0x3B, 0x3B, 0x3B)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)

HEX_HEADER_FILL = "1F3864"
HEX_ROW_ALT = "F2F2F2"
HEX_RISK_FILL = "FCE4D6"
HEX_RISK_EDGE = "C00000"

STATUS_COLORS = {
    "open": ("FFF2CC", "7F6000"),      # amber
    "done": ("E2EFDA", "375623"),      # green
    "blocked": ("FCE4D6", "833C0C"),   # red
    "unclear": ("FCE4D6", "833C0C"),   # red
}


# ---------- Low-level XML helpers ----------

def _set_cell_shading(cell, hex_fill: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_fill)
    tc_pr.append(shd)


def _set_cell_borders(cell, *, top=None, left=None, bottom=None, right=None) -> None:
    """Each edge is a tuple (size_in_eighths_of_pt, hex_color) or None."""
    tc_pr = cell._tc.get_or_add_tcPr()
    borders = tc_pr.find(qn("w:tcBorders"))
    if borders is None:
        borders = OxmlElement("w:tcBorders")
        tc_pr.append(borders)
    for edge_name, edge in (("top", top), ("left", left),
                            ("bottom", bottom), ("right", right)):
        if edge is None:
            continue
        size, color = edge
        e = borders.find(qn(f"w:{edge_name}"))
        if e is None:
            e = OxmlElement(f"w:{edge_name}")
            borders.append(e)
        e.set(qn("w:val"), "single")
        e.set(qn("w:sz"), str(size))
        e.set(qn("w:color"), color)


def _set_cell_margins(cell, *, top=100, bottom=100, left=140, right=140) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_mar = OxmlElement("w:tcMar")
    for side, val in (("top", top), ("bottom", bottom), ("left", left), ("right", right)):
        m = OxmlElement(f"w:{side}")
        m.set(qn("w:w"), str(val))
        m.set(qn("w:type"), "dxa")
        tc_mar.append(m)
    tc_pr.append(tc_mar)


def _add_bottom_border_to_para(paragraph, hex_color: str, size: int = 12) -> None:
    p_pr = paragraph._p.get_or_add_pPr()
    pbdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), str(size))
    bottom.set(qn("w:space"), "4")
    bottom.set(qn("w:color"), hex_color)
    pbdr.append(bottom)
    p_pr.append(pbdr)


def _add_top_border_to_para(paragraph, hex_color: str, size: int = 6) -> None:
    p_pr = paragraph._p.get_or_add_pPr()
    pbdr = OxmlElement("w:pBdr")
    top = OxmlElement("w:top")
    top.set(qn("w:val"), "single")
    top.set(qn("w:sz"), str(size))
    top.set(qn("w:space"), "4")
    top.set(qn("w:color"), hex_color)
    pbdr.append(top)
    p_pr.append(pbdr)


def _add_field(paragraph, instr_text: str) -> None:
    """Insert a Word field (e.g. PAGE, NUMPAGES) into the paragraph."""
    run = paragraph.add_run()
    fld_char1 = OxmlElement("w:fldChar")
    fld_char1.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = instr_text
    fld_char2 = OxmlElement("w:fldChar")
    fld_char2.set(qn("w:fldCharType"), "end")
    run._r.append(fld_char1)
    run._r.append(instr)
    run._r.append(fld_char2)


def _para(para, text: str, *, bold=False, italic=False, size_pt: int | None = None,
          color: RGBColor | None = None, font_name: str = "Calibri") -> None:
    run = para.add_run(text)
    run.font.name = font_name
    run.bold = bold
    run.italic = italic
    if size_pt is not None:
        run.font.size = Pt(size_pt)
    if color is not None:
        run.font.color.rgb = color


# ---------- Markdown parsing ----------

SECTION_ORDER = [
    "overview",
    "key decisions",
    "action items",
    "open risks & blockers",
    "key contacts",
    "outstanding questions",
]


def _parse_sections(md: str) -> tuple[str, dict[str, str]]:
    """Split the briefing markdown into {title, section_name -> raw_body}."""
    md = md.replace("\r\n", "\n")
    title_match = re.search(r"^#\s+(.+?)\s*$", md, re.MULTILINE)
    title = title_match.group(1).strip() if title_match else "Project Briefing"

    sections: dict[str, str] = {}
    pattern = re.compile(r"^##\s+(.+?)\s*$", re.MULTILINE)
    matches = list(pattern.finditer(md))
    for i, m in enumerate(matches):
        name = m.group(1).strip().lower()
        start = m.end()
        end = matches[i + 1].start() if i + 1 < len(matches) else len(md)
        sections[name] = md[start:end].strip()
    return title, sections


def _parse_pipe_table(block: str) -> list[list[str]]:
    """Parse a markdown pipe-table into a list of cell rows (header first)."""
    rows: list[list[str]] = []
    for line in block.splitlines():
        line = line.strip()
        if not line.startswith("|"):
            continue
        if re.match(r"^\|[\s\-:|]+\|$", line):  # separator row
            continue
        cells = [c.strip() for c in line.strip("|").split("|")]
        rows.append(cells)
    return rows


def _parse_bullets(block: str) -> list[str]:
    """Return each top-level bullet as one string, preserving inline markdown."""
    bullets: list[str] = []
    current: list[str] = []
    for line in block.splitlines():
        stripped = line.strip()
        if re.match(r"^[-*]\s+", stripped):
            if current:
                bullets.append(" ".join(current).strip())
                current = []
            current.append(re.sub(r"^[-*]\s+", "", stripped))
        elif stripped and current and not stripped.startswith("#"):
            current.append(stripped)
        elif not stripped and current:
            bullets.append(" ".join(current).strip())
            current = []
    if current:
        bullets.append(" ".join(current).strip())
    return [b for b in bullets if b]


def _split_bold(text: str) -> list[tuple[str, bool]]:
    """Split text on **bold** runs. Returns [(fragment, is_bold), ...]."""
    parts = re.split(r"\*\*(.+?)\*\*", text)
    result = []
    for i, p in enumerate(parts):
        if not p:
            continue
        result.append((p, i % 2 == 1))
    return result


def _add_inline_runs(paragraph, text: str, *, base_size: int = 11,
                     base_color: RGBColor | None = None) -> None:
    for frag, is_bold in _split_bold(text):
        run = paragraph.add_run(frag)
        run.font.name = "Calibri"
        run.font.size = Pt(base_size)
        run.bold = is_bold
        if base_color is not None:
            run.font.color.rgb = base_color


# ---------- Page setup, header, footer ----------

def _setup_landscape_a4(section) -> None:
    section.orientation = WD_ORIENTATION.LANDSCAPE
    section.page_width = Cm(29.7)
    section.page_height = Cm(21.0)
    section.top_margin = Cm(1.9)
    section.bottom_margin = Cm(1.9)
    section.left_margin = Cm(1.9)
    section.right_margin = Cm(1.9)


def _add_header(section, project: str) -> None:
    header = section.header
    p = header.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    _para(p, f"CDI Energy  ·  {project} Briefing",
          italic=False, size_pt=9, color=GREY_MID)


def _add_footer(section) -> None:
    footer = section.footer
    p = footer.paragraphs[0]
    # Left text, right pagination — use a tab-stop based layout
    p.paragraph_format.tab_stops.add_tab_stop(Cm(26))  # right tab near page edge
    _para(p, "Confidential — Internal CDI Energy", italic=True, size_pt=9, color=GREY_MID)
    run_tab = p.add_run("\t")
    run_tab.font.size = Pt(9)
    _para(p, "Page ", size_pt=9, color=GREY_MID)
    _add_field(p, "PAGE")
    _para(p, " of ", size_pt=9, color=GREY_MID)
    _add_field(p, "NUMPAGES")


# ---------- Section renderers ----------

def _add_cover(doc: Document, project: str, period: str | None) -> None:
    # Eyebrow
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(48)
    _para(p, "CDI ENERGY  //  PROJECT BRIEFING",
          bold=True, size_pt=10, color=ACCENT_BLUE)

    # Title
    p = doc.add_paragraph()
    _para(p, project, bold=True, size_pt=26, color=BRAND_NAVY)

    # Subtitle + divider
    sub = f"Email Briefing"
    if period:
        sub += f"  ·  Coverage: {period}"
    p = doc.add_paragraph()
    _para(p, sub, italic=True, size_pt=12, color=GREY_MID)
    _add_bottom_border_to_para(p, "2E75B6", size=12)

    # Meta mini-table (borderless)
    meta = doc.add_table(rows=2, cols=4)
    meta.alignment = WD_TABLE_ALIGNMENT.LEFT
    meta.autofit = True
    labels = ["Prepared for", "Period covered", "Issue date", "Classification"]
    values = ["CDI Energy", period or "—",
              _today_str(), "Confidential — Internal"]
    for col, (lab, val) in enumerate(zip(labels, values)):
        lc = meta.cell(0, col)
        vc = meta.cell(1, col)
        lc.paragraphs[0].clear()
        vc.paragraphs[0].clear()
        _para(lc.paragraphs[0], lab, bold=True, size_pt=9, color=GREY_MID)
        _para(vc.paragraphs[0], val, size_pt=10, color=GREY_DARK)
    _remove_table_borders(meta)

    doc.add_page_break()


def _today_str() -> str:
    from datetime import datetime
    return datetime.now().strftime("%d %b %Y")


def _remove_table_borders(table) -> None:
    tbl = table._tbl
    tbl_pr = tbl.find(qn("w:tblPr"))
    if tbl_pr is None:
        tbl_pr = OxmlElement("w:tblPr")
        tbl.insert(0, tbl_pr)
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        e = OxmlElement(f"w:{edge}")
        e.set(qn("w:val"), "none")
        borders.append(e)
    tbl_pr.append(borders)


def _add_h1(doc: Document, text: str) -> None:
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(2)
    _para(p, text, bold=True, size_pt=16, color=BRAND_NAVY)
    _add_bottom_border_to_para(p, "2E75B6", size=12)


def _add_paragraph_text(doc: Document, text: str) -> None:
    for line in text.strip().split("\n\n"):
        p = doc.add_paragraph()
        _add_inline_runs(p, line.strip())


def _add_overview(doc: Document, block: str) -> None:
    # Split into lead prose and bullets
    lines = block.strip().splitlines()
    prose: list[str] = []
    bullet_lines: list[str] = []
    tail_lines: list[str] = []
    seen_bullet = False
    for line in lines:
        if re.match(r"^[-*]\s+", line.strip()):
            seen_bullet = True
            bullet_lines.append(line.strip())
        elif seen_bullet:
            tail_lines.append(line.strip())
        else:
            prose.append(line)

    if prose:
        _add_paragraph_text(doc, "\n".join(prose))

    for b in bullet_lines:
        p = doc.add_paragraph(style="List Bullet")
        _add_inline_runs(p, re.sub(r"^[-*]\s+", "", b))

    tail = "\n".join(t for t in tail_lines if t)
    if tail:
        _add_paragraph_text(doc, tail)


def _add_styled_table(doc: Document, rows: list[list[str]],
                      *, widths_pct: list[float], chip_col: int | None = None) -> None:
    if not rows:
        return
    n_cols = len(rows[0])
    table = doc.add_table(rows=len(rows), cols=n_cols)
    table.autofit = False
    table.alignment = WD_TABLE_ALIGNMENT.LEFT

    # Column widths based on landscape A4 content area (~26 cm)
    content_cm = 26.0
    for col_idx in range(n_cols):
        pct = widths_pct[col_idx] if col_idx < len(widths_pct) else (100 / n_cols) / 100
        width_cm = content_cm * pct
        for r in range(len(rows)):
            table.cell(r, col_idx).width = Cm(width_cm)

    # Header row
    for col_idx, text in enumerate(rows[0]):
        cell = table.cell(0, col_idx)
        cell.paragraphs[0].clear()
        _set_cell_shading(cell, HEX_HEADER_FILL)
        _set_cell_margins(cell)
        _para(cell.paragraphs[0], text.strip(), bold=True, size_pt=10, color=WHITE)

    # Body rows
    for r_idx in range(1, len(rows)):
        for c_idx, text in enumerate(rows[r_idx]):
            cell = table.cell(r_idx, c_idx)
            cell.paragraphs[0].clear()
            _set_cell_margins(cell)
            # Zebra
            if r_idx % 2 == 0:
                _set_cell_shading(cell, HEX_ROW_ALT)
            # Chip column?
            if chip_col is not None and c_idx == chip_col:
                _render_status_chip(cell, text.strip())
            else:
                _add_inline_runs(cell.paragraphs[0], text.strip(), base_size=10,
                                 base_color=GREY_DARK)


def _render_status_chip(cell, raw: str) -> None:
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    normal = raw.strip().strip("*").strip("`").strip()
    key = normal.lower()
    bg_hex, fg_hex = STATUS_COLORS.get(key, (HEX_ROW_ALT, "3B3B3B"))
    _set_cell_shading(cell, bg_hex)
    run = p.add_run(normal or "—")
    run.font.name = "Calibri"
    run.font.size = Pt(10)
    run.bold = True
    run.font.color.rgb = RGBColor.from_string(fg_hex)


def _add_risk_cards(doc: Document, bullets: list[str]) -> None:
    if not bullets:
        p = doc.add_paragraph()
        _para(p, "_None identified._", italic=True, size_pt=10, color=GREY_MID)
        return

    for idx, bullet in enumerate(bullets, start=1):
        table = doc.add_table(rows=1, cols=1)
        table.autofit = False
        cell = table.cell(0, 0)
        cell.width = Cm(26)
        _set_cell_shading(cell, HEX_RISK_FILL)
        _set_cell_margins(cell, top=160, bottom=160, left=200, right=200)
        _set_cell_borders(
            cell,
            top=(12, HEX_RISK_EDGE),
            left=(24, HEX_RISK_EDGE),
            bottom=(4, HEX_RISK_EDGE),
            right=(4, HEX_RISK_EDGE),
        )

        # Title + detail + source split
        title, detail, source = _split_risk_bullet(bullet)

        p_title = cell.paragraphs[0]
        p_title.clear()
        _para(p_title, f"RISK {idx:02d}  ·  {title}",
              bold=True, size_pt=11, color=RGBColor.from_string("833C0C"))

        if detail:
            p_detail = cell.add_paragraph()
            _add_inline_runs(p_detail, detail, base_size=10, base_color=GREY_DARK)

        if source:
            p_src = cell.add_paragraph()
            _para(p_src, f"Source: {source}", italic=True, size_pt=9, color=GREY_MID)

        # Breathing room between cards
        doc.add_paragraph()


def _split_risk_bullet(text: str) -> tuple[str, str, str]:
    """Return (title, detail, source). Best-effort heuristic parse."""
    # Source: strip trailing "Source: ..." or "(Source: ...)" etc.
    source = ""
    m = re.search(r"(?i)(?:^|\s)source[:\-]\s*(.+?)(?:\.|$)", text)
    if m:
        source = m.group(1).strip().rstrip(".")
        text = text[:m.start()].strip().rstrip(".") + "."

    # Title: leading **bold** text if present, else first 10 words
    m = re.match(r"\*\*(.+?)\*\*\s*[.\-—:]*\s*(.*)", text.strip(), flags=re.DOTALL)
    if m:
        title = m.group(1).strip()
        detail = m.group(2).strip()
    else:
        words = text.split()
        title = " ".join(words[:10]).rstrip(".,;:")
        detail = " ".join(words[10:]).strip()
    return title, detail, source


# ---------- Contacts ----------

ORG_HINTS = [
    ("Horizon Power", ["horizonpower.com.au", "horizonpower"]),
    ("CDI Energy", ["cdienergy.com.au", "cdienergy"]),
]


def _classify_contact(email: str, relevance: str = "") -> str:
    email_low = (email or "").lower()
    for name, hints in ORG_HINTS:
        if any(h in email_low for h in hints):
            return name
    if email_low and "@" in email_low:
        return "Engineering Partners & Other"
    return "Engineering Partners & Other"


def _add_contacts(doc: Document, block: str) -> None:
    rows = _parse_pipe_table(block)
    if not rows:
        # No table — emit raw text
        _add_paragraph_text(doc, block)
        return

    header = rows[0]
    body = rows[1:]
    # Expected columns roughly: Name | Email | Role | Relevance
    groups: "OrderedDict[str, list[list[str]]]" = OrderedDict()
    groups["Horizon Power"] = []
    groups["CDI Energy"] = []
    groups["Engineering Partners & Other"] = []

    # Find email column index
    email_idx = next((i for i, h in enumerate(header) if "email" in h.lower()), 1)
    relevance_idx = next((i for i, h in enumerate(header) if "relevance" in h.lower()), -1)

    for row in body:
        email_cell = row[email_idx] if email_idx < len(row) else ""
        rel_cell = row[relevance_idx] if 0 <= relevance_idx < len(row) else ""
        org = _classify_contact(email_cell, rel_cell)
        groups.setdefault(org, []).append(row)

    first = True
    for org, rows_for_org in groups.items():
        if not rows_for_org:
            continue
        p = doc.add_paragraph()
        if not first:
            p.paragraph_format.space_before = Pt(14)
        _para(p, org, bold=True, size_pt=13, color=BRAND_NAVY)
        first = False

        table_rows = [header] + rows_for_org
        widths = [0.18, 0.28, 0.23, 0.31]
        _add_styled_table(doc, table_rows, widths_pct=widths)


# ---------- Outstanding questions ----------

def _add_questions(doc: Document, block: str) -> None:
    bullets = _parse_bullets(block)
    if not bullets:
        p = doc.add_paragraph()
        _para(p, "_None identified._", italic=True, size_pt=10, color=GREY_MID)
        return
    for b in bullets:
        p = doc.add_paragraph(style="List Bullet")
        _add_inline_runs(p, b, base_size=10)


# ---------- Action items ----------

def _add_actions(doc: Document, block: str) -> None:
    rows = _parse_pipe_table(block)
    if not rows:
        _add_paragraph_text(doc, block)
        return
    widths = [0.15, 0.42, 0.13, 0.09, 0.21]
    header = rows[0]
    status_col = next((i for i, h in enumerate(header) if "status" in h.lower()), 3)
    _add_styled_table(doc, rows, widths_pct=widths, chip_col=status_col)


def _add_decisions(doc: Document, block: str) -> None:
    rows = _parse_pipe_table(block)
    if rows:
        widths = [0.30, 0.49, 0.21]
        _add_styled_table(doc, rows, widths_pct=widths)
        return
    # Fallback — render as bullets
    for b in _parse_bullets(block):
        p = doc.add_paragraph(style="List Bullet")
        _add_inline_runs(p, b, base_size=10)


# ---------- Signoff ----------

def _add_signoff(doc: Document) -> None:
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(36)
    _add_top_border_to_para(p, "BFBFBF", size=6)
    _para(p, "— End of briefing —", italic=True, size_pt=10, color=GREY_MID)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER


# ---------- Entry point ----------

def render_cdi_briefing(markdown_text: str, dest: Path, project: str,
                        period: str | None = None) -> None:
    doc = Document()
    section = doc.sections[0]
    _setup_landscape_a4(section)
    _add_header(section, project)
    _add_footer(section)

    title, parsed = _parse_sections(markdown_text)
    _add_cover(doc, project or title, period)

    order = [
        ("Overview", "overview", _add_overview),
        ("Key Decisions", "key decisions", _add_decisions),
        ("Action Items", "action items", _add_actions),
        ("Open Risks & Blockers", "open risks & blockers", _add_risk_cards_wrap),
        ("Key Contacts", "key contacts", _add_contacts),
        ("Outstanding Questions", "outstanding questions", _add_questions),
    ]

    page_break_before = {"Action Items", "Open Risks & Blockers", "Key Contacts"}

    for heading, key, fn in order:
        if heading in page_break_before and any(key in parsed for _, key, _ in order):
            doc.add_page_break()
        _add_h1(doc, heading)
        body = parsed.get(key, "").strip()
        if body:
            fn(doc, body)
        else:
            p = doc.add_paragraph()
            _para(p, "_None identified._", italic=True, size_pt=10, color=GREY_MID)

    _add_signoff(doc)
    doc.save(str(dest))


def _add_risk_cards_wrap(doc: Document, block: str) -> None:
    _add_risk_cards(doc, _parse_bullets(block))
