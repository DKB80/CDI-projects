"""Helpers for inspecting the locally running Outlook — used by the
first-run wizard and the main window to show which mailbox will be read."""


def list_outlook_accounts() -> tuple[bool, list[str], str]:
    """Return (connected, account_emails, error_message).

    connected=True + account_emails populated: Outlook running and we can see
    one or more mailboxes.
    connected=False: Outlook isn't running or COM dispatch failed.
    """
    try:
        import pythoncom  # noqa: F401
        import win32com.client
    except ImportError:
        return False, [], "pywin32 not installed."

    try:
        app = win32com.client.Dispatch("Outlook.Application")
        ns = app.GetNamespace("MAPI")
    except Exception as e:
        return False, [], f"Can't connect to Outlook: {e}"

    emails = []
    try:
        for account in ns.Accounts:
            addr = getattr(account, "SmtpAddress", None) or getattr(account, "DisplayName", None)
            if addr:
                emails.append(addr)
    except Exception:
        try:
            for store in ns.Folders:
                emails.append(store.Name)
        except Exception as e:
            return True, [], f"Connected, but couldn't list mailboxes: {e}"

    return True, emails, ""
