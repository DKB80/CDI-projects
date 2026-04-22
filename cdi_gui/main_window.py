"""Main window of the CDI Outlook Briefing Tool.

Branding scaffolding is in _build_header() — replace CDI_* constants or the
logo image file below with real brand assets once available.
"""

import os
import queue
import shutil
import subprocess
import sys
import threading
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, simpledialog

import ttkbootstrap as tb
from ttkbootstrap.constants import DANGER, INFO, PRIMARY, SECONDARY, SUCCESS, WARNING
from ttkbootstrap.scrolled import ScrolledFrame

from . import config as cfg
from .outlook_info import list_outlook_accounts
from .preview import PreviewDialog
from .wizard import WizardWindow


# ---------- Branding (swap these for real CDI assets) ----------
CDI_COMPANY = "CDI Energy"
CDI_APP_TITLE = "Outlook Briefing Tool"
CDI_TAGLINE = "Summarise a project's email history in one click."
CDI_LOGO_FILE = Path(__file__).parent / "assets" / "cdi_logo.png"  # optional
CDI_THEME = "cosmo"  # ttkbootstrap theme — swap for 'flatly', 'yeti', etc.


class MainWindow(tb.Window):
    def __init__(self):
        super().__init__(themename=CDI_THEME)
        self.title(f"{CDI_COMPANY} — {CDI_APP_TITLE}")
        self.geometry("1000x760")
        self.minsize(860, 680)

        self.app_config = cfg.load_config()
        self.log_queue: queue.Queue = queue.Queue()
        self.preview_event = threading.Event()
        self.preview_response: list[str] | None = None
        self.preview_matches: list[dict] | None = None
        self.running = False

        self._build_header()
        self._build_form()
        self._build_log()
        self._poll_log_queue()

        if cfg.is_first_run():
            self.after(300, self._open_wizard)

    # ==================== UI CONSTRUCTION ====================

    def _build_header(self):
        bar = tb.Frame(self, padding=(20, 14), bootstyle=PRIMARY)
        bar.pack(fill="x")

        if CDI_LOGO_FILE.exists():
            try:
                self.logo_img = tk.PhotoImage(file=str(CDI_LOGO_FILE))
                tb.Label(bar, image=self.logo_img, bootstyle="inverse-primary").pack(side="left", padx=(0, 14))
            except Exception:
                pass

        left = tb.Frame(bar, bootstyle=PRIMARY)
        left.pack(side="left", fill="y")
        tb.Label(left, text=CDI_COMPANY, font=("Segoe UI", 10, "bold"),
                 bootstyle="inverse-primary").pack(anchor="w")
        tb.Label(left, text=CDI_APP_TITLE, font=("Segoe UI", 18, "bold"),
                 bootstyle="inverse-primary").pack(anchor="w")
        tb.Label(left, text=CDI_TAGLINE, bootstyle="inverse-primary").pack(anchor="w")

        right = tb.Frame(bar, bootstyle=PRIMARY)
        right.pack(side="right", fill="y")
        tb.Button(right, text="Settings", bootstyle="light",
                  command=self._open_wizard).pack(side="right")

    def _build_form(self):
        # Scrollable form — keeps the app usable at any window size.
        outer = ScrolledFrame(self, autohide=True, padding=20, height=440)
        outer.pack(fill="both", expand=True)

        # Presets row
        presets_row = tb.Frame(outer)
        presets_row.pack(fill="x", pady=(0, 12))
        tb.Label(presets_row, text="Preset:").pack(side="left")
        self.preset_var = tb.StringVar()
        self.preset_combo = tb.Combobox(presets_row, textvariable=self.preset_var,
                                        values=cfg.list_presets(), state="readonly", width=30)
        self.preset_combo.pack(side="left", padx=(6, 0))
        tb.Button(presets_row, text="Load", bootstyle=SECONDARY,
                  command=self._load_selected_preset).pack(side="left", padx=(6, 0))
        tb.Button(presets_row, text="Save as…", bootstyle=SECONDARY,
                  command=self._save_preset).pack(side="left", padx=(6, 0))
        tb.Button(presets_row, text="Delete", bootstyle=DANGER,
                  command=self._delete_preset).pack(side="left", padx=(6, 0))

        # Project name
        tb.Label(outer, text="Project name  (required — also used as folder name)",
                 font=("Segoe UI", 10, "bold")).pack(anchor="w")
        self.project_var = tb.StringVar()
        tb.Entry(outer, textvariable=self.project_var).pack(fill="x", pady=(2, 10))

        # Keywords
        tb.Label(outer, text="Keywords  (one per line — emails matching ANY of these are found)",
                 font=("Segoe UI", 10, "bold")).pack(anchor="w")
        self.keywords_txt = tb.Text(outer, height=4)
        self.keywords_txt.pack(fill="x", pady=(2, 10))

        # Exclusions
        tb.Label(outer, text="Exclude phrases  (optional — emails containing any of these are dropped)",
                 font=("Segoe UI", 10, "bold")).pack(anchor="w")
        self.exclude_txt = tb.Text(outer, height=3)
        self.exclude_txt.pack(fill="x", pady=(2, 10))

        # Months + attachments + summary
        opts = tb.Frame(outer)
        opts.pack(fill="x", pady=(0, 10))

        tb.Label(opts, text="Months to look back:").pack(side="left")
        self.months_var = tb.IntVar(value=12)
        tb.Spinbox(opts, from_=1, to=120, textvariable=self.months_var, width=6).pack(side="left", padx=(6, 16))

        self.attach_var = tb.BooleanVar(value=True)
        tb.Checkbutton(opts, variable=self.attach_var,
                       text="Extract attachments").pack(side="left", padx=(0, 16))

        self.summary_var = tb.BooleanVar(value=True)
        tb.Checkbutton(opts, variable=self.summary_var,
                       text="Generate Claude summary").pack(side="left", padx=(0, 16))

        self.dl_copy_var = tb.BooleanVar(value=True)
        tb.Checkbutton(opts, variable=self.dl_copy_var,
                       text="Also copy Word doc to Downloads").pack(side="left")

        # Second options row
        opts2 = tb.Frame(outer)
        opts2.pack(fill="x", pady=(0, 10))
        self.save_msg_var = tb.BooleanVar(value=True)
        tb.Checkbutton(opts2, variable=self.save_msg_var,
                       text="Save full .msg copies of every email  (untick for a faster 'summary only' run)"
                       ).pack(side="left")

        # Output folder
        out_row = tb.Frame(outer)
        out_row.pack(fill="x", pady=(0, 10))
        tb.Label(out_row, text="Output folder:").pack(side="left")
        self.output_var = tb.StringVar(value=self.app_config.get("default_output", str(cfg.DEFAULT_OUTPUT_ROOT)))
        tb.Entry(out_row, textvariable=self.output_var).pack(side="left", fill="x", expand=True, padx=(6, 6))
        tb.Button(out_row, text="Browse…", bootstyle=SECONDARY,
                  command=self._browse_output).pack(side="left")

        # Advanced section
        adv = tb.Labelframe(outer, text="Advanced", padding=10)
        adv.pack(fill="x", pady=(0, 10))

        self.deleted_var = tb.BooleanVar(value=False)
        tb.Checkbutton(adv, variable=self.deleted_var,
                       text="Include Deleted Items folder").pack(anchor="w")
        self.junk_var = tb.BooleanVar(value=False)
        tb.Checkbutton(adv, variable=self.junk_var,
                       text="Include Junk Email folder").pack(anchor="w")
        self.exclude_sent_var = tb.BooleanVar(value=False)
        tb.Checkbutton(adv, variable=self.exclude_sent_var,
                       text="Exclude Sent Items (only search emails you received)").pack(anchor="w")

        r = tb.Frame(adv)
        r.pack(fill="x", pady=(6, 0))
        tb.Label(r, text="Only search folders containing (comma-separated, blank = all):").pack(anchor="w")
        self.only_folders_var = tb.StringVar()
        tb.Entry(r, textvariable=self.only_folders_var).pack(fill="x", pady=(2, 0))

        r2 = tb.Frame(adv)
        r2.pack(fill="x", pady=(6, 0))
        tb.Label(r2, text="Max emails to feed the Claude summary:").pack(side="left")
        self.max_sum_var = tb.IntVar(value=200)
        tb.Spinbox(r2, from_=10, to=2000, increment=10, textvariable=self.max_sum_var, width=8).pack(side="left", padx=(6, 0))

        # Run button
        runbar = tb.Frame(outer)
        runbar.pack(fill="x", pady=(6, 0))
        self.run_btn = tb.Button(runbar, text="Run scrape", bootstyle=SUCCESS,
                                 command=self._on_run)
        self.run_btn.pack(side="left")
        self.open_out_btn = tb.Button(runbar, text="Open output folder", bootstyle=SECONDARY,
                                      command=self._open_last_output, state="disabled")
        self.open_out_btn.pack(side="left", padx=(8, 0))
        self.open_dl_btn = tb.Button(runbar, text="Open Downloads folder", bootstyle=SECONDARY,
                                     command=lambda: self._open_folder(cfg.DOWNLOADS_DIR))
        self.open_dl_btn.pack(side="left", padx=(8, 0))
        self.last_output_dir: Path | None = None

    def _build_log(self):
        outer = tb.Frame(self, padding=(20, 0, 20, 20))
        outer.pack(fill="both", expand=True)
        tb.Label(outer, text="Progress", font=("Segoe UI", 10, "bold")).pack(anchor="w")
        self.log_widget = tb.Text(outer, height=10, state="disabled", wrap="word")
        self.log_widget.pack(fill="both", expand=True, pady=(2, 0))

    # ==================== PRESETS ====================

    def _current_form_data(self) -> dict:
        return {
            "project": self.project_var.get().strip(),
            "keywords": self.keywords_txt.get("1.0", "end").strip(),
            "exclude": self.exclude_txt.get("1.0", "end").strip(),
            "months": int(self.months_var.get()),
            "attachments": bool(self.attach_var.get()),
            "summary": bool(self.summary_var.get()),
            "dl_copy": bool(self.dl_copy_var.get()),
            "save_msg": bool(self.save_msg_var.get()),
            "output": self.output_var.get().strip(),
            "include_deleted": bool(self.deleted_var.get()),
            "include_junk": bool(self.junk_var.get()),
            "exclude_sent": bool(self.exclude_sent_var.get()),
            "only_folders": self.only_folders_var.get().strip(),
            "max_summary": int(self.max_sum_var.get()),
        }

    def _apply_preset(self, data: dict) -> None:
        self.project_var.set(data.get("project", ""))
        self.keywords_txt.delete("1.0", "end"); self.keywords_txt.insert("1.0", data.get("keywords", ""))
        self.exclude_txt.delete("1.0", "end"); self.exclude_txt.insert("1.0", data.get("exclude", ""))
        self.months_var.set(int(data.get("months", 12)))
        self.attach_var.set(bool(data.get("attachments", True)))
        self.summary_var.set(bool(data.get("summary", True)))
        self.dl_copy_var.set(bool(data.get("dl_copy", True)))
        self.output_var.set(data.get("output", str(cfg.DEFAULT_OUTPUT_ROOT)))
        self.deleted_var.set(bool(data.get("include_deleted", False)))
        self.junk_var.set(bool(data.get("include_junk", False)))
        self.exclude_sent_var.set(bool(data.get("exclude_sent", False)))
        self.only_folders_var.set(data.get("only_folders", ""))
        self.max_sum_var.set(int(data.get("max_summary", 200)))
        self.save_msg_var.set(bool(data.get("save_msg", True)))

    def _load_selected_preset(self):
        name = self.preset_var.get().strip()
        if not name:
            messagebox.showinfo("Load preset", "Pick a preset from the dropdown first.")
            return
        try:
            self._apply_preset(cfg.load_preset(name))
            self._log(f"Loaded preset '{name}'.")
        except Exception as e:
            messagebox.showerror("Load preset failed", str(e))

    def _save_preset(self):
        default = self.project_var.get().strip() or "preset"
        name = simpledialog.askstring("Save preset", "Preset name:", initialvalue=default, parent=self)
        if not name:
            return
        cfg.save_preset(name, self._current_form_data())
        self.preset_combo.configure(values=cfg.list_presets())
        self.preset_var.set(name)
        self._log(f"Saved preset '{name}'.")

    def _delete_preset(self):
        name = self.preset_var.get().strip()
        if not name:
            return
        if messagebox.askyesno("Delete preset", f"Delete preset '{name}'?"):
            cfg.delete_preset(name)
            self.preset_combo.configure(values=cfg.list_presets())
            self.preset_var.set("")
            self._log(f"Deleted preset '{name}'.")

    # ==================== ACTIONS ====================

    def _browse_output(self):
        d = filedialog.askdirectory(initialdir=self.output_var.get() or str(Path.home()))
        if d:
            self.output_var.set(d)

    def _open_folder(self, path: Path):
        if not path.exists():
            messagebox.showwarning("Folder not found", str(path))
            return
        if sys.platform == "win32":
            os.startfile(str(path))  # noqa: SIM115
        else:
            subprocess.Popen(["xdg-open", str(path)])

    def _open_last_output(self):
        if self.last_output_dir:
            self._open_folder(self.last_output_dir)

    def _open_wizard(self):
        def _done(new_cfg):
            self.app_config = new_cfg
            if new_cfg.get("default_output"):
                self.output_var.set(new_cfg["default_output"])
        WizardWindow(self, _done)

    # ==================== RUN (threaded) ====================

    def _on_run(self):
        if self.running:
            return
        data = self._current_form_data()
        if not data["project"]:
            messagebox.showerror("Missing info", "Enter a project name.")
            return
        keywords = [k.strip() for k in data["keywords"].splitlines() if k.strip()]
        if not keywords:
            messagebox.showerror("Missing info", "Enter at least one keyword.")
            return
        excludes = [e.strip() for e in data["exclude"].splitlines() if e.strip()]
        only_folders = [s.strip() for s in data["only_folders"].split(",") if s.strip()] or None

        if data["summary"] and not (self.app_config.get("anthropic_api_key") or os.environ.get("ANTHROPIC_API_KEY")):
            messagebox.showwarning(
                "No API key",
                "You selected 'Generate Claude summary' but no API key is set. "
                "Open Settings to add one, or untick the summary option to run without it."
            )
            return

        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        safe_proj = "".join(c if c.isalnum() or c in "-_ " else "_" for c in data["project"]).strip()
        output_dir = Path(data["output"]) / f"{safe_proj}_{ts}"

        # Push API key into env for the summary call
        key = self.app_config.get("anthropic_api_key")
        if key:
            os.environ["ANTHROPIC_API_KEY"] = key

        self.running = True
        self.run_btn.configure(state="disabled", text="Running…")
        self.log_widget.configure(state="normal"); self.log_widget.delete("1.0", "end"); self.log_widget.configure(state="disabled")

        threading.Thread(
            target=self._run_worker,
            args=(keywords, excludes, data, only_folders, output_dir),
            daemon=True,
        ).start()

    def _run_worker(self, keywords, excludes, data, only_folders, output_dir: Path):
        try:
            import pythoncom  # COM must be initialised per-thread
            pythoncom.CoInitialize()
        except Exception as e:
            self.log_queue.put(("DONE", {"errors": [f"COM init failed: {e}"]}))
            return

        try:
            # Lazy import to keep GUI startup fast
            from scrape_outlook_local import scrape_outlook

            def log_cb(msg: str) -> None:
                self.log_queue.put(("LOG", msg))

            def preview_cb(matches: list[dict]) -> list[str] | None:
                self.preview_matches = matches
                self.preview_response = None
                self.preview_event.clear()
                self.log_queue.put(("PREVIEW", None))
                self.preview_event.wait()
                return self.preview_response

            extra_exclude = ["Sent Items"] if data["exclude_sent"] else None

            result = scrape_outlook(
                keywords=keywords,
                exclude_keywords=excludes,
                project_label=data["project"],
                months=data["months"],
                output=output_dir,
                include_deleted=data["include_deleted"],
                include_junk=data["include_junk"],
                extra_exclude_folders=extra_exclude,
                include_folders_only=only_folders,
                skip_attachments=not data["attachments"],
                save_msg_files=data["save_msg"],
                generate_summary=data["summary"],
                max_emails_for_summary=data["max_summary"],
                log=log_cb,
                preview_callback=preview_cb,
            )

            # Copy Word doc to Downloads
            if data["dl_copy"] and result.get("summary_docx_path"):
                try:
                    cfg.DOWNLOADS_DIR.mkdir(parents=True, exist_ok=True)
                    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    safe = "".join(c if c.isalnum() or c in "-_ " else "_" for c in data["project"]).strip()
                    dest = cfg.DOWNLOADS_DIR / f"{safe}_briefing_{stamp}.docx"
                    shutil.copyfile(result["summary_docx_path"], dest)
                    result["downloads_copy_path"] = str(dest)
                    log_cb(f"Copied Word doc to {dest}")
                except Exception as e:
                    log_cb(f"Couldn't copy Word doc to Downloads: {e}")

            self.log_queue.put(("DONE", result))
        except Exception as e:
            self.log_queue.put(("DONE", {"errors": [f"Scrape failed: {e}"]}))
        finally:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    # ==================== LOG POLL ====================

    def _poll_log_queue(self):
        try:
            while True:
                kind, payload = self.log_queue.get_nowait()
                if kind == "LOG":
                    self._log(payload)
                elif kind == "PREVIEW":
                    self._show_preview()
                elif kind == "DONE":
                    self._finish(payload)
        except queue.Empty:
            pass
        self.after(100, self._poll_log_queue)

    def _log(self, msg: str):
        self.log_widget.configure(state="normal")
        self.log_widget.insert("end", msg.rstrip() + "\n")
        self.log_widget.see("end")
        self.log_widget.configure(state="disabled")

    def _show_preview(self):
        if not self.preview_matches:
            self.preview_response = []
            self.preview_event.set()
            return
        dialog = PreviewDialog(self, self.preview_matches)
        self.wait_window(dialog)
        self.preview_response = dialog.result()
        self.preview_event.set()

    def _finish(self, result: dict):
        self.running = False
        self.run_btn.configure(state="normal", text="Run scrape")
        errors = result.get("errors", [])
        for e in errors:
            self._log(f"ERROR: {e}")

        if result.get("emails_dir"):
            self.last_output_dir = Path(result["emails_dir"]).parent
            self.open_out_btn.configure(state="normal")

        match_count = result.get("match_count", 0)
        saved_count = result.get("saved_count", 0)
        lines = [
            f"Matches found:       {match_count}",
            f"Emails saved:        {saved_count}",
        ]
        if result.get("summary_md_path"):
            lines.append(f"Markdown summary:    {result['summary_md_path']}")
        if result.get("summary_docx_path"):
            lines.append(f"Word summary:        {result['summary_docx_path']}")
        if result.get("downloads_copy_path"):
            lines.append(f"Copied to Downloads: {result['downloads_copy_path']}")
        msg = "\n".join(lines)

        if saved_count > 0:
            messagebox.showinfo("Scrape complete", msg)
        elif match_count > 0:
            messagebox.showinfo("Scrape finished", msg + "\n\n(Nothing was saved.)")
        else:
            messagebox.showinfo("Scrape finished", "No matching emails found in the date range.")


def main():
    app = MainWindow()
    app.mainloop()


if __name__ == "__main__":
    main()
