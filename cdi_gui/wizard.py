"""First-run setup wizard. Walks a new user through API key entry,
Outlook account verification, and default output folder choice.

Launched automatically the first time the GUI runs, or anytime via
the main window's Settings button.
"""

import webbrowser
from pathlib import Path
from tkinter import filedialog

import ttkbootstrap as tb
from ttkbootstrap.constants import DISABLED, NORMAL, PRIMARY, SECONDARY, SUCCESS

from . import config as cfg
from .outlook_info import list_outlook_accounts


WELCOME_TEXT = (
    "This tool scans your Outlook mailbox for emails related to a project, "
    "saves copies of every match to a folder, and uses Claude (Anthropic's "
    "AI) to produce a concise Word document briefing with decisions, action "
    "items, risks, and open questions.\n\n"
    "Before you start, we need three things:\n\n"
    "  1. Your Anthropic API key — this lets the tool call Claude to write "
    "the briefing.\n\n"
    "  2. Confirmation that Classic Outlook is running, signed in to the "
    "email account you want to search.\n\n"
    "  3. A default folder to save outputs to.\n\n"
    "Click Next to get started. You can re-run this wizard later from the "
    "main window's Settings button."
)

API_KEY_HELP = (
    "What is an API key?\n\n"
    "An API key is a password that lets apps on your computer talk to "
    "Anthropic's Claude service. The tool uses it to send the email contents "
    "to Claude so it can write the briefing. Your key is stored only on this "
    "computer.\n\n"
    "How to get one:\n\n"
    "  • Personal / trial: go to https://console.anthropic.com, sign up, "
    "then click 'Get API keys' in the left sidebar and 'Create Key'.\n\n"
    "  • CDI Energy Team plan (once CDI signs up): your IT admin will "
    "provide a shared team key. Paste that value here.\n\n"
    "The key looks like:  sk-ant-api03-xxxxxxxxxxxxxxxxxxxxxxxx\n\n"
    "Click the button below to open the Anthropic Console in your browser."
)


class WizardWindow(tb.Toplevel):
    def __init__(self, parent, on_finish):
        super().__init__(parent)
        self.title("CDI Energy — Outlook Briefing Tool: First-time setup")
        self.geometry("720x560")
        self.on_finish = on_finish
        self.saved = cfg.load_config()
        self.current_step = 0

        self.container = tb.Frame(self, padding=20)
        self.container.pack(fill="both", expand=True)

        self.steps = [self._step_welcome, self._step_api_key,
                      self._step_outlook, self._step_output, self._step_done]
        self._render()

    def _render(self):
        for w in self.container.winfo_children():
            w.destroy()
        self.steps[self.current_step]()

    def _footer(self, can_back=True, can_next=True, next_label="Next", finish=False):
        bar = tb.Frame(self.container)
        bar.pack(side="bottom", fill="x", pady=(16, 0))

        if can_back:
            tb.Button(bar, text="Back", bootstyle=SECONDARY, command=self._back).pack(side="left")

        if finish:
            tb.Button(bar, text="Finish", bootstyle=SUCCESS, command=self._finish).pack(side="right")
        elif can_next:
            tb.Button(bar, text=next_label, bootstyle=PRIMARY, command=self._next).pack(side="right")

    def _next(self):
        if self.current_step < len(self.steps) - 1:
            self.current_step += 1
            self._render()

    def _back(self):
        if self.current_step > 0:
            self.current_step -= 1
            self._render()

    def _finish(self):
        self.saved["setup_complete"] = True
        cfg.save_config(self.saved)
        self.on_finish(self.saved)
        self.destroy()

    # ---------- Step 1: Welcome ----------
    def _step_welcome(self):
        tb.Label(self.container, text="Welcome to the CDI Outlook Briefing Tool",
                 font=("Segoe UI", 16, "bold")).pack(anchor="w")
        tb.Label(self.container, text="Step 1 of 4 — Introduction",
                 bootstyle=SECONDARY).pack(anchor="w", pady=(0, 12))
        body = tb.Text(self.container, wrap="word", height=18, relief="flat")
        body.insert("1.0", WELCOME_TEXT)
        body.configure(state="disabled")
        body.pack(fill="both", expand=True)
        self._footer(can_back=False)

    # ---------- Step 2: API key ----------
    def _step_api_key(self):
        tb.Label(self.container, text="Anthropic API key",
                 font=("Segoe UI", 16, "bold")).pack(anchor="w")
        tb.Label(self.container, text="Step 2 of 4 — Required",
                 bootstyle=SECONDARY).pack(anchor="w", pady=(0, 12))

        help_frame = tb.Text(self.container, wrap="word", height=12, relief="flat")
        help_frame.insert("1.0", API_KEY_HELP)
        help_frame.configure(state="disabled")
        help_frame.pack(fill="x")

        tb.Button(self.container, text="Open Anthropic Console in browser",
                  bootstyle=SECONDARY,
                  command=lambda: webbrowser.open("https://console.anthropic.com/")
                  ).pack(anchor="w", pady=(8, 12))

        tb.Label(self.container, text="Paste your API key here:").pack(anchor="w")
        self.api_key_var = tb.StringVar(value=self.saved.get("anthropic_api_key", ""))
        entry = tb.Entry(self.container, textvariable=self.api_key_var, show="•", width=60)
        entry.pack(fill="x", pady=(4, 0))

        tb.Label(self.container,
                 text="(The key stays on this computer. It is stored in your user AppData folder.)",
                 bootstyle=SECONDARY, font=("Segoe UI", 9)).pack(anchor="w", pady=(2, 0))

        self._footer(next_label="Next")

        # Save on leaving the step
        def _persist(*_):
            self.saved["anthropic_api_key"] = self.api_key_var.get().strip()
        self.api_key_var.trace_add("write", _persist)

    # ---------- Step 3: Outlook check ----------
    def _step_outlook(self):
        tb.Label(self.container, text="Outlook connection check",
                 font=("Segoe UI", 16, "bold")).pack(anchor="w")
        tb.Label(self.container, text="Step 3 of 4 — Verify mailbox",
                 bootstyle=SECONDARY).pack(anchor="w", pady=(0, 12))

        tb.Label(self.container, wraplength=660, text=(
            "Make sure Classic Outlook is open and signed in to the mailbox you "
            "want to search. Then click the button below — the tool will list "
            "the email accounts it can see. If your work account is in the list, "
            "you're good."
        )).pack(anchor="w")

        self.status_frame = tb.Frame(self.container, padding=12)
        self.status_frame.pack(fill="x", pady=12)

        tb.Button(self.container, text="Check Outlook now", bootstyle=PRIMARY,
                  command=self._do_outlook_check).pack(anchor="w")

        self._footer()

    def _do_outlook_check(self):
        for w in self.status_frame.winfo_children():
            w.destroy()

        connected, emails, err = list_outlook_accounts()
        if not connected:
            tb.Label(self.status_frame, bootstyle="danger",
                     text=f"Could not connect to Outlook.\n{err}\n\n"
                          f"Please open Classic Outlook and wait for it to finish "
                          f"loading, then click 'Check Outlook now' again.").pack(anchor="w")
            return
        if not emails:
            tb.Label(self.status_frame, bootstyle="warning",
                     text="Connected to Outlook, but no email accounts were visible. "
                          "Check that your work account is added in Outlook's Account Settings.").pack(anchor="w")
            return

        tb.Label(self.status_frame, bootstyle=SUCCESS,
                 text="Outlook is running and the following mailbox(es) are available:",
                 font=("Segoe UI", 10, "bold")).pack(anchor="w")
        for email in emails:
            tb.Label(self.status_frame, text=f"   • {email}").pack(anchor="w")
        self.saved["outlook_accounts"] = emails

    # ---------- Step 4: Output folder ----------
    def _step_output(self):
        tb.Label(self.container, text="Default output folder",
                 font=("Segoe UI", 16, "bold")).pack(anchor="w")
        tb.Label(self.container, text="Step 4 of 4 — Where do results go?",
                 bootstyle=SECONDARY).pack(anchor="w", pady=(0, 12))

        tb.Label(self.container, wraplength=660, text=(
            "Each scrape creates its own sub-folder under the location you "
            "choose, named after the project. Every run gets a timestamped "
            "sub-folder so you don't overwrite old results.\n\n"
            "The default is your Documents folder. You can change it per-run "
            "in the main window if you want."
        )).pack(anchor="w")

        default = self.saved.get("default_output", str(cfg.DEFAULT_OUTPUT_ROOT))
        self.output_var = tb.StringVar(value=default)

        row = tb.Frame(self.container)
        row.pack(fill="x", pady=12)
        tb.Entry(row, textvariable=self.output_var).pack(side="left", fill="x", expand=True)
        tb.Button(row, text="Browse…", bootstyle=SECONDARY,
                  command=self._browse).pack(side="left", padx=(8, 0))

        tb.Label(self.container, wraplength=660, bootstyle=SECONDARY,
                 text="Tip: the Word briefing is also automatically copied to your "
                 "Downloads folder when the scrape finishes, so you can find it "
                 "quickly even if the output folder is buried.").pack(anchor="w", pady=(16, 0))

        def _persist(*_):
            self.saved["default_output"] = self.output_var.get().strip()
        self.output_var.trace_add("write", _persist)

        self._footer()

    def _browse(self):
        d = filedialog.askdirectory(initialdir=self.output_var.get() or str(Path.home()))
        if d:
            self.output_var.set(d)

    # ---------- Step 5: Done ----------
    def _step_done(self):
        tb.Label(self.container, text="You're all set.",
                 font=("Segoe UI", 16, "bold"), bootstyle=SUCCESS).pack(anchor="w")
        tb.Label(self.container, bootstyle=SECONDARY,
                 text="Click Finish to open the main window.").pack(anchor="w", pady=(0, 12))

        tb.Label(self.container, wraplength=660, text=(
            "From the main window you can:\n\n"
            "  • Enter keywords and exclusions for a project, click Run, and "
            "review the matching emails before saving.\n\n"
            "  • Save the settings as a preset (e.g. 'Project 307') so you "
            "can re-run it next month in two clicks.\n\n"
            "  • Re-open this wizard anytime via the Settings button if you "
            "rotate your API key or change Outlook accounts."
        )).pack(anchor="w")

        self._footer(can_next=False, finish=True)
