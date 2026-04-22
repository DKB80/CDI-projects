"""Dialog that shows every sender from the user's mailbox (last N months)
with a search box and multi-select, so they can pick who to filter on.

Opened from the main window's 'Pick from mailbox…' button.
"""

import tkinter as tk
from tkinter import ttk

import ttkbootstrap as tb
from ttkbootstrap.constants import DANGER, PRIMARY, SECONDARY


class SenderPickerDialog(tb.Toplevel):
    """Modal dialog. After wait_window(), call .result() for the picks."""

    def __init__(self, parent, senders: list[dict]):
        super().__init__(parent)
        self.title("Pick senders to filter on")
        self.geometry("840x560")
        self.senders = senders
        self._selected: list[str] | None = None
        self._build()
        self.transient(parent)
        self.grab_set()

    def _build(self):
        outer = tb.Frame(self, padding=16)
        outer.pack(fill="both", expand=True)

        tb.Label(
            outer, font=("Segoe UI", 14, "bold"),
            text=f"{len(self.senders)} unique senders in the last 12 months",
        ).pack(anchor="w")
        tb.Label(outer, bootstyle=SECONDARY, text=(
            "Tick who the scrape should filter on. Matching is a substring "
            "check on email OR display name — e.g. picking 'hossein.sarmadian@…' "
            "also matches emails sent from shared boxes that include 'hossein' "
            "in the display name."
        )).pack(anchor="w", pady=(0, 12))

        # Search box
        search_row = tb.Frame(outer)
        search_row.pack(fill="x", pady=(0, 6))
        tb.Label(search_row, text="Filter:").pack(side="left")
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *_: self._refresh())
        tb.Entry(search_row, textvariable=self.search_var, width=40
                 ).pack(side="left", padx=(6, 0), fill="x", expand=True)
        tb.Button(search_row, text="Tick visible", bootstyle=SECONDARY,
                  command=lambda: self._bulk_visible(True)).pack(side="left", padx=(6, 0))
        tb.Button(search_row, text="Untick visible", bootstyle=SECONDARY,
                  command=lambda: self._bulk_visible(False)).pack(side="left", padx=(6, 0))

        # Treeview
        tree_frame = tb.Frame(outer)
        tree_frame.pack(fill="both", expand=True)

        cols = ("picked", "email", "name", "count")
        self.tree = ttk.Treeview(tree_frame, columns=cols, show="headings",
                                 selectmode="extended")
        self.tree.heading("picked", text="✓")
        self.tree.heading("email", text="Email")
        self.tree.heading("name", text="Name")
        self.tree.heading("count", text="Emails")
        self.tree.column("picked", width=32, anchor="center", stretch=False)
        self.tree.column("email", width=340, stretch=True)
        self.tree.column("name", width=260, stretch=True)
        self.tree.column("count", width=80, anchor="e", stretch=False)

        scroll = tb.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll.set)
        self.tree.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")

        # Click on a row toggles its tick.
        self.tree.bind("<Button-1>", self._on_click)
        self.tree.bind("<space>", self._on_space)
        self.tree.bind("<Return>", self._on_space)

        self._ticked: set[str] = set()
        self._refresh()

        # Footer
        bar = tb.Frame(outer)
        bar.pack(fill="x", pady=(12, 0))
        self.count_label = tb.Label(bar, text="")
        self.count_label.pack(side="left")
        tb.Button(bar, text="Cancel", bootstyle=DANGER,
                  command=self._cancel).pack(side="right")
        tb.Button(bar, text="Use selected", bootstyle=PRIMARY,
                  command=self._confirm).pack(side="right", padx=(0, 6))

        self._update_count()

    # ----- helpers -----

    def _refresh(self):
        q = self.search_var.get().strip().lower()
        self.tree.delete(*self.tree.get_children())
        for s in self.senders:
            hay = f"{s['email']} {s['name']}".lower()
            if q and q not in hay:
                continue
            tick = "✓" if s["email"].lower() in self._ticked else ""
            self.tree.insert("", "end", iid=s["email"],
                             values=(tick, s["email"], s["name"], s["count"]))
        self._update_count()

    def _on_click(self, event):
        row_id = self.tree.identify_row(event.y)
        col = self.tree.identify_column(event.x)
        if not row_id:
            return
        # Clicking the tick column toggles; clicking elsewhere only selects.
        if col == "#1":
            self._toggle(row_id)
            return "break"

    def _on_space(self, _event):
        for row_id in self.tree.selection():
            self._toggle(row_id)
        return "break"

    def _toggle(self, row_id: str):
        key = row_id.lower()
        if key in self._ticked:
            self._ticked.discard(key)
            self.tree.set(row_id, "picked", "")
        else:
            self._ticked.add(key)
            self.tree.set(row_id, "picked", "✓")
        self._update_count()

    def _bulk_visible(self, on: bool):
        for row_id in self.tree.get_children():
            key = row_id.lower()
            if on:
                self._ticked.add(key)
                self.tree.set(row_id, "picked", "✓")
            else:
                self._ticked.discard(key)
                self.tree.set(row_id, "picked", "")
        self._update_count()

    def _update_count(self):
        self.count_label.configure(
            text=f"{len(self._ticked)} ticked of {len(self.senders)} senders"
        )

    def _confirm(self):
        # Return original-case emails for everything ticked.
        self._selected = [
            s["email"] for s in self.senders
            if s["email"].lower() in self._ticked
        ]
        self.destroy()

    def _cancel(self):
        self._selected = None
        self.destroy()

    def result(self) -> list[str] | None:
        return self._selected
