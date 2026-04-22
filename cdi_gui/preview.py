"""Preview dialog — shown after the scan phase, before saving.

Lists every matching email with a checkbox so the user can untick noise
before the tool commits to saving and summarizing.
"""

import tkinter as tk
import ttkbootstrap as tb
from ttkbootstrap.constants import PRIMARY, SECONDARY, DANGER


class PreviewDialog(tb.Toplevel):
    """Modal dialog that blocks the caller until user clicks Save or Cancel."""

    def __init__(self, parent, matches: list[dict]):
        super().__init__(parent)
        self.title("Preview matches — tick the emails to save")
        self.geometry("1000x620")
        self.matches = matches
        self.approved_ids: list[str] | None = None
        self._vars: dict[str, tk.BooleanVar] = {}
        self._build()
        self.transient(parent)
        self.grab_set()

    def _build(self):
        outer = tb.Frame(self, padding=16)
        outer.pack(fill="both", expand=True)

        tb.Label(
            outer, font=("Segoe UI", 14, "bold"),
            text=f"{len(self.matches)} matching emails found",
        ).pack(anchor="w")
        tb.Label(outer, bootstyle=SECONDARY, text=(
            "Untick anything that isn't relevant (e.g. generic newsletters, "
            "auto-replies). Ticked emails will be saved to disk as .msg files "
            "and fed into the Claude summary."
        )).pack(anchor="w", pady=(0, 12))

        # Bulk-select controls
        top = tb.Frame(outer)
        top.pack(fill="x", pady=(0, 8))
        tb.Button(top, text="Tick all", bootstyle=SECONDARY,
                  command=lambda: self._bulk(True)).pack(side="left")
        tb.Button(top, text="Untick all", bootstyle=SECONDARY,
                  command=lambda: self._bulk(False)).pack(side="left", padx=(6, 0))
        self.count_var = tb.StringVar()
        tb.Label(top, textvariable=self.count_var).pack(side="right")

        # Scrollable list
        container = tb.Frame(outer)
        container.pack(fill="both", expand=True)

        canvas = tk.Canvas(container, highlightthickness=0)
        scroll = tb.Scrollbar(container, orient="vertical", command=canvas.yview)
        self.inner = tb.Frame(canvas)
        self.inner.bind("<Configure>",
                        lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.inner, anchor="nw")
        canvas.configure(yscrollcommand=scroll.set)
        canvas.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")
        canvas.bind_all(
            "<MouseWheel>",
            lambda e: canvas.yview_scroll(int(-1 * (e.delta / 120)), "units"),
        )

        for m in self.matches:
            self._render_row(self.inner, m)

        # Footer
        bar = tb.Frame(outer)
        bar.pack(fill="x", pady=(12, 0))
        tb.Button(bar, text="Cancel — save nothing", bootstyle=DANGER,
                  command=self._cancel).pack(side="left")
        tb.Button(bar, text="Save ticked emails", bootstyle=PRIMARY,
                  command=self._confirm).pack(side="right")

        self._update_count()

    def _render_row(self, parent, m: dict):
        row = tb.Frame(parent, padding=(8, 4))
        row.pack(fill="x", pady=1)

        var = tk.BooleanVar(value=True)
        var.trace_add("write", lambda *_: self._update_count())
        self._vars[m["entry_id"]] = var

        tb.Checkbutton(row, variable=var).pack(side="left")

        text = tb.Frame(row)
        text.pack(side="left", fill="x", expand=True, padx=(6, 0))

        date = (m.get("date") or "")[:10]
        sender = m.get("from") or "unknown"
        subject = (m.get("subject") or "(no subject)").strip()
        attach = " 📎" if m.get("has_attachments") else ""

        line1 = tb.Label(
            text, font=("Segoe UI", 10, "bold"),
            text=f"{subject}{attach}",
            anchor="w", wraplength=820, justify="left",
        )
        line1.pack(fill="x", anchor="w")

        line2 = tb.Label(
            text, bootstyle=SECONDARY,
            text=f"{date}    {sender}",
        )
        line2.pack(fill="x", anchor="w")

        preview = (m.get("body_preview") or "").strip()
        if preview:
            tb.Label(text, bootstyle=SECONDARY, font=("Segoe UI", 9),
                     text=preview[:180], wraplength=820, anchor="w",
                     justify="left").pack(fill="x", anchor="w")

    def _bulk(self, value: bool):
        for v in self._vars.values():
            v.set(value)

    def _update_count(self):
        selected = sum(1 for v in self._vars.values() if v.get())
        self.count_var.set(f"{selected} of {len(self._vars)} selected")

    def _confirm(self):
        self.approved_ids = [eid for eid, v in self._vars.items() if v.get()]
        self.destroy()

    def _cancel(self):
        self.approved_ids = None
        self.destroy()

    def result(self) -> list[str] | None:
        """Call after .wait_window(self) — returns list of approved entry IDs or None."""
        return self.approved_ids
