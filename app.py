"""CV Emailer - a small desktop app to send your CV to many headhunters /
companies at once, with a personalised email for each recipient.

Run it with:   python app.py
"""

from __future__ import annotations

import queue
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path

import store
import finder
from sender import (
    SmtpConfig,
    SendReport,
    friendly_smtp_error,
    personalise,
    send_bulk,
    verify_connection,
)

APP_TITLE = "CV Emailer"

# Known provider presets so the user can fill SMTP settings in one click.
PRESETS = {
    "Office 365 / University": ("smtp.office365.com", 587, "starttls"),
    "Gmail": ("smtp.gmail.com", 587, "starttls"),
    "Outlook.com": ("smtp-mail.outlook.com", 587, "starttls"),
    "Yahoo": ("smtp.mail.yahoo.com", 465, "ssl"),
}


class CVEmailer(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("860x640")
        self.minsize(760, 560)

        # In-memory state.
        self.recipients: list[dict] = store.load_recipients()
        self.draft = store.load_draft()
        self.settings = store.load_settings()
        self._password = store.load_password(self.settings.get("username", ""))

        # Cross-thread communication for the send worker.
        self._send_queue: "queue.Queue" = queue.Queue()
        self._stop_flag = threading.Event()
        self._sending = False

        self._build_ui()
        self._refresh_recipient_table()
        self._load_draft_into_ui()
        self._load_settings_into_ui()

        self.protocol("WM_DELETE_WINDOW", self._on_close)

    # ------------------------------------------------------------------ #
    # UI construction
    # ------------------------------------------------------------------ #
    def _build_ui(self) -> None:
        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=8, pady=8)
        self.tab_recipients = ttk.Frame(nb)
        self.tab_compose = ttk.Frame(nb)
        self.tab_settings = ttk.Frame(nb)
        self.tab_send = ttk.Frame(nb)
        nb.add(self.tab_recipients, text="  1. Recipients  ")
        nb.add(self.tab_compose, text="  2. Compose  ")
        nb.add(self.tab_settings, text="  3. Settings  ")
        nb.add(self.tab_send, text="  4. Send  ")

        self._build_recipients_tab()
        self._build_compose_tab()
        self._build_settings_tab()
        self._build_send_tab()

    # ---- Recipients tab ---------------------------------------------- #
    def _build_recipients_tab(self) -> None:
        frame = self.tab_recipients
        bar = ttk.Frame(frame)
        bar.pack(fill="x", padx=8, pady=(8, 4))
        ttk.Button(bar, text="Add", command=self._add_recipient).pack(side="left")
        ttk.Button(bar, text="Edit", command=self._edit_recipient).pack(side="left", padx=4)
        ttk.Button(bar, text="Remove", command=self._remove_recipient).pack(side="left")
        ttk.Button(bar, text="Import CSV / Excel…", command=self._import_recipients).pack(side="left", padx=4)
        ttk.Button(bar, text="Find emails (Hunter.io)…", command=self._open_finder).pack(side="left")
        ttk.Button(bar, text="Clear all", command=self._clear_recipients).pack(side="left", padx=4)
        self.lbl_count = ttk.Label(bar, text="")
        self.lbl_count.pack(side="right")

        cols = ("name", "email", "company")
        self.tree = ttk.Treeview(frame, columns=cols, show="headings", selectmode="extended")
        self.tree.heading("name", text="Name")
        self.tree.heading("email", text="Email")
        self.tree.heading("company", text="Company")
        self.tree.column("name", width=200)
        self.tree.column("email", width=300)
        self.tree.column("company", width=240)
        self.tree.pack(fill="both", expand=True, padx=8, pady=(0, 8))
        self.tree.bind("<Double-1>", lambda _e: self._edit_recipient())

        scroll = ttk.Scrollbar(self.tree, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y")

    # ---- Compose tab ------------------------------------------------- #
    def _build_compose_tab(self) -> None:
        frame = self.tab_compose
        pad = {"padx": 8, "pady": 4}

        ttk.Label(frame, text="Subject:").grid(row=0, column=0, sticky="w", **pad)
        self.var_subject = tk.StringVar()
        ttk.Entry(frame, textvariable=self.var_subject).grid(
            row=0, column=1, sticky="ew", **pad
        )

        ttk.Label(frame, text="Body:").grid(row=1, column=0, sticky="nw", **pad)
        self.txt_body = tk.Text(frame, height=14, wrap="word")
        self.txt_body.grid(row=1, column=1, sticky="nsew", **pad)

        hint = ("Placeholders you can use:  {name}   {first_name}   {company}   {email}\n"
                "Each recipient gets their own email with these filled in.")
        ttk.Label(frame, text=hint, foreground="#555").grid(
            row=2, column=1, sticky="w", padx=8
        )

        ttk.Label(frame, text="Attachments\n(your CV):").grid(row=3, column=0, sticky="nw", **pad)
        att_frame = ttk.Frame(frame)
        att_frame.grid(row=3, column=1, sticky="ew", **pad)
        self.lst_attachments = tk.Listbox(att_frame, height=4)
        self.lst_attachments.pack(side="left", fill="both", expand=True)
        att_btns = ttk.Frame(att_frame)
        att_btns.pack(side="left", fill="y", padx=4)
        ttk.Button(att_btns, text="Add file…", command=self._add_attachment).pack(fill="x")
        ttk.Button(att_btns, text="Remove", command=self._remove_attachment).pack(fill="x", pady=4)

        ttk.Button(frame, text="Preview for first recipient",
                   command=self._preview_email).grid(row=4, column=1, sticky="w", padx=8, pady=8)

        frame.columnconfigure(1, weight=1)
        frame.rowconfigure(1, weight=1)

    # ---- Settings tab ------------------------------------------------ #
    def _build_settings_tab(self) -> None:
        frame = self.tab_settings
        pad = {"padx": 8, "pady": 5}
        r = 0

        ttk.Label(frame, text="Quick preset:").grid(row=r, column=0, sticky="w", **pad)
        preset_bar = ttk.Frame(frame)
        preset_bar.grid(row=r, column=1, sticky="w", **pad)
        for name in PRESETS:
            ttk.Button(preset_bar, text=name,
                       command=lambda n=name: self._apply_preset(n)).pack(side="left", padx=2)
        r += 1

        self.var_host = tk.StringVar()
        self.var_port = tk.StringVar()
        self.var_security = tk.StringVar()
        self.var_username = tk.StringVar()
        self.var_fromname = tk.StringVar()
        self.var_password = tk.StringVar()
        self.var_remember = tk.BooleanVar()
        self.var_delay = tk.StringVar()

        def row_entry(label, var, show=None):
            nonlocal r
            ttk.Label(frame, text=label).grid(row=r, column=0, sticky="w", **pad)
            ttk.Entry(frame, textvariable=var, show=show).grid(row=r, column=1, sticky="ew", **pad)
            r += 1

        ttk.Label(frame, text="SMTP host:").grid(row=r, column=0, sticky="w", **pad)
        ttk.Entry(frame, textvariable=self.var_host).grid(row=r, column=1, sticky="ew", **pad)
        r += 1
        ttk.Label(frame, text="Port:").grid(row=r, column=0, sticky="w", **pad)
        ttk.Entry(frame, textvariable=self.var_port, width=10).grid(row=r, column=1, sticky="w", **pad)
        r += 1
        ttk.Label(frame, text="Security:").grid(row=r, column=0, sticky="w", **pad)
        ttk.Combobox(frame, textvariable=self.var_security, state="readonly",
                     values=["starttls", "ssl", "none"], width=12).grid(
            row=r, column=1, sticky="w", **pad)
        r += 1

        row_entry("Your email (username):", self.var_username)
        row_entry("Display name (From):", self.var_fromname)
        row_entry("Password / App password:", self.var_password, show="•")

        remember_text = ("Remember password securely (Windows Credential Manager)"
                         if store.has_secure_storage()
                         else "Secure storage unavailable — password kept for this session only")
        cb = ttk.Checkbutton(frame, text=remember_text, variable=self.var_remember)
        cb.grid(row=r, column=1, sticky="w", **pad)
        if not store.has_secure_storage():
            cb.state(["disabled"])
        r += 1

        ttk.Label(frame, text="Delay between emails (sec):").grid(row=r, column=0, sticky="w", **pad)
        ttk.Entry(frame, textvariable=self.var_delay, width=10).grid(row=r, column=1, sticky="w", **pad)
        r += 1

        btns = ttk.Frame(frame)
        btns.grid(row=r, column=1, sticky="w", **pad)
        ttk.Button(btns, text="Save settings", command=self._save_settings).pack(side="left")
        ttk.Button(btns, text="Test connection", command=self._test_connection).pack(side="left", padx=6)

        frame.columnconfigure(1, weight=1)

    # ---- Send tab ---------------------------------------------------- #
    def _build_send_tab(self) -> None:
        frame = self.tab_send
        bar = ttk.Frame(frame)
        bar.pack(fill="x", padx=8, pady=8)
        self.btn_test_self = ttk.Button(bar, text="Send test to myself", command=self._send_test_to_self)
        self.btn_test_self.pack(side="left")
        self.btn_send_all = ttk.Button(bar, text="Send to ALL recipients", command=self._send_all)
        self.btn_send_all.pack(side="left", padx=6)
        self.btn_stop = ttk.Button(bar, text="Stop", command=self._stop_sending, state="disabled")
        self.btn_stop.pack(side="left")

        self.progress = ttk.Progressbar(frame, mode="determinate")
        self.progress.pack(fill="x", padx=8)
        self.lbl_progress = ttk.Label(frame, text="Idle.")
        self.lbl_progress.pack(anchor="w", padx=8, pady=2)

        self.txt_log = tk.Text(frame, height=18, state="disabled", wrap="word")
        self.txt_log.pack(fill="both", expand=True, padx=8, pady=8)
        self.txt_log.tag_config("ok", foreground="#1a7f37")
        self.txt_log.tag_config("err", foreground="#cf222e")
        self.txt_log.tag_config("info", foreground="#0969da")

    # ------------------------------------------------------------------ #
    # Recipients logic
    # ------------------------------------------------------------------ #
    def _refresh_recipient_table(self) -> None:
        self.tree.delete(*self.tree.get_children())
        for rec in self.recipients:
            self.tree.insert("", "end", values=(rec.get("name", ""),
                                                rec.get("email", ""),
                                                rec.get("company", "")))
        self.lbl_count.config(text=f"{len(self.recipients)} recipient(s)")

    def _add_recipient(self) -> None:
        result = RecipientDialog(self, "Add recipient").result
        if result:
            self.recipients.append(result)
            store.save_recipients(self.recipients)
            self._refresh_recipient_table()

    def _selected_indexes(self) -> list[int]:
        return [self.tree.index(i) for i in self.tree.selection()]

    def _edit_recipient(self) -> None:
        idxs = self._selected_indexes()
        if not idxs:
            return
        idx = idxs[0]
        result = RecipientDialog(self, "Edit recipient", self.recipients[idx]).result
        if result:
            self.recipients[idx] = result
            store.save_recipients(self.recipients)
            self._refresh_recipient_table()

    def _remove_recipient(self) -> None:
        idxs = sorted(self._selected_indexes(), reverse=True)
        if not idxs:
            return
        for idx in idxs:
            del self.recipients[idx]
        store.save_recipients(self.recipients)
        self._refresh_recipient_table()

    def _clear_recipients(self) -> None:
        if self.recipients and messagebox.askyesno(APP_TITLE, "Remove ALL recipients?"):
            self.recipients = []
            store.save_recipients(self.recipients)
            self._refresh_recipient_table()

    def _import_recipients(self) -> None:
        path = filedialog.askopenfilename(
            title="Import recipients",
            filetypes=[("Spreadsheets", "*.csv *.xlsx *.xlsm"),
                       ("CSV", "*.csv"), ("Excel", "*.xlsx *.xlsm"), ("All files", "*.*")],
        )
        if not path:
            return
        try:
            imported = store.import_recipients_from_file(path)
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror(APP_TITLE, f"Could not import:\n\n{exc}")
            return
        if not imported:
            messagebox.showwarning(
                APP_TITLE,
                "No rows with an email address were found.\n\n"
                "Make sure the file has a header row with columns named "
                "'name', 'email' and 'company'.",
            )
            return
        existing = {r["email"].lower() for r in self.recipients}
        added = 0
        for rec in imported:
            if rec["email"].lower() not in existing:
                self.recipients.append(rec)
                existing.add(rec["email"].lower())
                added += 1
        store.save_recipients(self.recipients)
        self._refresh_recipient_table()
        messagebox.showinfo(
            APP_TITLE,
            f"Imported {added} new recipient(s).\n"
            f"({len(imported) - added} duplicate(s) skipped.)",
        )

    def _open_finder(self) -> None:
        api_key = store.load_hunter_key()
        if not api_key:
            messagebox.showwarning(
                APP_TITLE,
                "No Hunter.io API key saved.\n\n"
                "Sign up for a free key at hunter.io, then I'll store it for you.",
            )
            return
        found = FinderDialog(self, api_key).added
        if not found:
            return
        existing = {r["email"].lower() for r in self.recipients}
        added = 0
        for rec in found:
            if rec["email"] and rec["email"].lower() not in existing:
                self.recipients.append(rec)
                existing.add(rec["email"].lower())
                added += 1
        store.save_recipients(self.recipients)
        self._refresh_recipient_table()
        messagebox.showinfo(APP_TITLE, f"Added {added} new contact(s) to your list.")

    # ------------------------------------------------------------------ #
    # Compose logic
    # ------------------------------------------------------------------ #
    def _load_draft_into_ui(self) -> None:
        self.var_subject.set(self.draft.get("subject", ""))
        self.txt_body.delete("1.0", "end")
        self.txt_body.insert("1.0", self.draft.get("body", ""))
        self.lst_attachments.delete(0, "end")
        for path in self.draft.get("attachments", []):
            self.lst_attachments.insert("end", path)

    def _collect_draft(self) -> dict:
        return {
            "subject": self.var_subject.get(),
            "body": self.txt_body.get("1.0", "end-1c"),
            "attachments": list(self.lst_attachments.get(0, "end")),
        }

    def _save_draft(self) -> None:
        self.draft = self._collect_draft()
        store.save_draft(self.draft)

    def _add_attachment(self) -> None:
        paths = filedialog.askopenfilenames(title="Choose attachment(s)")
        for p in paths:
            self.lst_attachments.insert("end", p)
        self._save_draft()

    def _remove_attachment(self) -> None:
        for i in reversed(self.lst_attachments.curselection()):
            self.lst_attachments.delete(i)
        self._save_draft()

    def _preview_email(self) -> None:
        if not self.recipients:
            messagebox.showinfo(APP_TITLE, "Add at least one recipient first.")
            return
        rec = self.recipients[0]
        subject = personalise(self.var_subject.get(), rec)
        body = personalise(self.txt_body.get("1.0", "end-1c"), rec)
        atts = "\n".join(Path(p).name for p in self.lst_attachments.get(0, "end")) or "(none)"
        PreviewDialog(self, rec, subject, body, atts)

    # ------------------------------------------------------------------ #
    # Settings logic
    # ------------------------------------------------------------------ #
    def _load_settings_into_ui(self) -> None:
        s = self.settings
        self.var_host.set(s.get("host", ""))
        self.var_port.set(str(s.get("port", 587)))
        self.var_security.set(s.get("security", "starttls"))
        self.var_username.set(s.get("username", ""))
        self.var_fromname.set(s.get("from_name", ""))
        self.var_remember.set(bool(s.get("remember_password", False)))
        self.var_delay.set(str(s.get("send_delay_seconds", 2.0)))
        self.var_password.set(self._password)

    def _apply_preset(self, name: str) -> None:
        host, port, security = PRESETS[name]
        self.var_host.set(host)
        self.var_port.set(str(port))
        self.var_security.set(security)

    def _collect_settings(self) -> dict:
        try:
            port = int(self.var_port.get())
        except ValueError:
            port = 587
        try:
            delay = float(self.var_delay.get())
        except ValueError:
            delay = 2.0
        return {
            "host": self.var_host.get().strip(),
            "port": port,
            "security": self.var_security.get(),
            "username": self.var_username.get().strip(),
            "from_name": self.var_fromname.get().strip(),
            "remember_password": bool(self.var_remember.get()),
            "send_delay_seconds": delay,
        }

    def _save_settings(self) -> None:
        self.settings = self._collect_settings()
        store.save_settings(self.settings)
        self._password = self.var_password.get()
        username = self.settings["username"]
        if self.settings["remember_password"] and self._password:
            if store.save_password(username, self._password):
                pass
        else:
            store.delete_password(username)
        messagebox.showinfo(APP_TITLE, "Settings saved.")

    def _current_smtp_config(self) -> SmtpConfig:
        s = self._collect_settings()
        return SmtpConfig(
            host=s["host"], port=s["port"], security=s["security"],
            username=s["username"], password=self.var_password.get(),
            from_name=s["from_name"],
        )

    def _test_connection(self) -> None:
        cfg = self._current_smtp_config()
        if not cfg.host or not cfg.username:
            messagebox.showwarning(APP_TITLE, "Fill in host and your email first.")
            return
        try:
            verify_connection(cfg)
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror(APP_TITLE, friendly_smtp_error(exc))
            return
        messagebox.showinfo(APP_TITLE, "Success — connected and logged in. ✅")

    # ------------------------------------------------------------------ #
    # Sending
    # ------------------------------------------------------------------ #
    def _log(self, text: str, tag: str = "") -> None:
        self.txt_log.config(state="normal")
        self.txt_log.insert("end", text + "\n", tag)
        self.txt_log.see("end")
        self.txt_log.config(state="disabled")

    def _validate_before_send(self) -> SmtpConfig | None:
        self._save_draft()
        self._save_settings_silent()
        cfg = self._current_smtp_config()
        if not cfg.host or not cfg.username:
            messagebox.showwarning(APP_TITLE, "Fill in SMTP host and your email (Settings tab).")
            return None
        if not self.var_password.get():
            messagebox.showwarning(APP_TITLE, "Enter your password (Settings tab).")
            return None
        if not self.var_subject.get().strip():
            if not messagebox.askyesno(APP_TITLE, "Subject is empty. Send anyway?"):
                return None
        return cfg

    def _save_settings_silent(self) -> None:
        self.settings = self._collect_settings()
        store.save_settings(self.settings)
        self._password = self.var_password.get()
        if self.settings["remember_password"] and self._password:
            store.save_password(self.settings["username"], self._password)

    def _send_test_to_self(self) -> None:
        cfg = self._validate_before_send()
        if not cfg:
            return
        me = {"name": cfg.from_name or "Me", "email": cfg.username, "company": "Test Co"}
        if not messagebox.askyesno(
            APP_TITLE, f"Send one test email to yourself ({cfg.username})?"
        ):
            return
        self._start_send(cfg, [me], delay=0.0, is_test=True)

    def _send_all(self) -> None:
        cfg = self._validate_before_send()
        if not cfg:
            return
        if not self.recipients:
            messagebox.showinfo(APP_TITLE, "No recipients to send to.")
            return
        draft = self._collect_draft()
        if not draft["attachments"]:
            if not messagebox.askyesno(APP_TITLE, "No CV/attachment added. Send anyway?"):
                return
        if not messagebox.askyesno(
            APP_TITLE,
            f"Send a personalised email to {len(self.recipients)} recipient(s)?\n\n"
            "Tip: run 'Send test to myself' first if you haven't.",
        ):
            return
        self._start_send(cfg, list(self.recipients),
                         delay=self.settings.get("send_delay_seconds", 2.0))

    def _start_send(self, cfg: SmtpConfig, recipients: list[dict],
                    delay: float, is_test: bool = False) -> None:
        if self._sending:
            return
        draft = self._collect_draft()
        self._sending = True
        self._stop_flag.clear()
        self.btn_send_all.config(state="disabled")
        self.btn_test_self.config(state="disabled")
        self.btn_stop.config(state="normal")
        self.progress.config(maximum=len(recipients), value=0)
        self._log(f"--- Starting {'test ' if is_test else ''}send to "
                  f"{len(recipients)} recipient(s) ---", "info")

        def progress_cb(i, total, email, status, detail):
            self._send_queue.put(("progress", i, total, email, status, detail))

        def worker():
            try:
                report = send_bulk(
                    cfg, recipients, draft["subject"], draft["body"],
                    draft["attachments"], delay_seconds=delay,
                    progress=progress_cb, should_stop=self._stop_flag.is_set,
                )
                self._send_queue.put(("done", report))
            except Exception as exc:  # noqa: BLE001
                self._send_queue.put(("fatal", exc))

        threading.Thread(target=worker, daemon=True).start()
        self.after(100, self._poll_send_queue)

    def _poll_send_queue(self) -> None:
        try:
            while True:
                item = self._send_queue.get_nowait()
                kind = item[0]
                if kind == "progress":
                    _, i, total, email, status, detail = item
                    self.progress.config(value=i)
                    self.lbl_progress.config(text=f"{i}/{total} processed")
                    if status == "sent":
                        self._log(f"[{i}/{total}] ✓ sent to {email} {detail}".rstrip(), "ok")
                    else:
                        self._log(f"[{i}/{total}] ✗ FAILED {email}: {detail}", "err")
                elif kind == "done":
                    self._finish_send(item[1])
                    return
                elif kind == "fatal":
                    self._log(friendly_smtp_error(item[1]), "err")
                    messagebox.showerror(APP_TITLE, friendly_smtp_error(item[1]))
                    self._finish_send(None)
                    return
        except queue.Empty:
            pass
        if self._sending:
            self.after(100, self._poll_send_queue)

    def _finish_send(self, report: SendReport | None) -> None:
        self._sending = False
        self.btn_send_all.config(state="normal")
        self.btn_test_self.config(state="normal")
        self.btn_stop.config(state="disabled")
        if report is not None:
            self._log(f"--- Done. {len(report.sent)} sent, "
                      f"{len(report.failed)} failed. ---", "info")
            self.lbl_progress.config(
                text=f"Finished: {len(report.sent)} sent, {len(report.failed)} failed.")
        else:
            self.lbl_progress.config(text="Stopped due to error.")

    def _stop_sending(self) -> None:
        self._stop_flag.set()
        self._log("Stop requested — finishing current email…", "info")

    # ------------------------------------------------------------------ #
    def _on_close(self) -> None:
        if self._sending:
            if not messagebox.askyesno(APP_TITLE, "A send is in progress. Quit anyway?"):
                return
        self._save_draft()
        store.save_settings(self._collect_settings())
        self.destroy()


# ---------------------------------------------------------------------- #
# Small modal dialogs
# ---------------------------------------------------------------------- #
class RecipientDialog(tk.Toplevel):
    def __init__(self, parent, title, initial: dict | None = None):
        super().__init__(parent)
        self.title(title)
        self.result: dict | None = None
        self.transient(parent)
        self.resizable(False, False)
        initial = initial or {}

        self.var_name = tk.StringVar(value=initial.get("name", ""))
        self.var_email = tk.StringVar(value=initial.get("email", ""))
        self.var_company = tk.StringVar(value=initial.get("company", ""))

        for i, (label, var) in enumerate(
            [("Name", self.var_name), ("Email *", self.var_email), ("Company", self.var_company)]
        ):
            ttk.Label(self, text=label).grid(row=i, column=0, sticky="w", padx=10, pady=6)
            ttk.Entry(self, textvariable=var, width=40).grid(row=i, column=1, padx=10, pady=6)

        btns = ttk.Frame(self)
        btns.grid(row=3, column=0, columnspan=2, pady=10)
        ttk.Button(btns, text="OK", command=self._ok).pack(side="left", padx=4)
        ttk.Button(btns, text="Cancel", command=self.destroy).pack(side="left", padx=4)

        self.bind("<Return>", lambda _e: self._ok())
        self.grab_set()
        self.wait_window()

    def _ok(self):
        email = self.var_email.get().strip()
        if not email or "@" not in email:
            messagebox.showwarning("Recipient", "A valid email address is required.")
            return
        self.result = {
            "name": self.var_name.get().strip(),
            "email": email,
            "company": self.var_company.get().strip(),
        }
        self.destroy()


class PreviewDialog(tk.Toplevel):
    def __init__(self, parent, recipient, subject, body, attachments):
        super().__init__(parent)
        self.title("Preview")
        self.geometry("600x460")
        self.transient(parent)
        info = (f"To: {recipient.get('name','')} <{recipient.get('email','')}>\n"
                f"Subject: {subject}\n"
                f"Attachments: {attachments}\n"
                + "-" * 60 + "\n\n" + body)
        txt = tk.Text(self, wrap="word")
        txt.insert("1.0", info)
        txt.config(state="disabled")
        txt.pack(fill="both", expand=True, padx=8, pady=8)
        ttk.Button(self, text="Close", command=self.destroy).pack(pady=6)
        self.grab_set()


class FinderDialog(tk.Toplevel):
    """Search company domains for published recruiter/HR emails via Hunter.io."""

    def __init__(self, parent, api_key: str):
        super().__init__(parent)
        self.title("Find emails (Hunter.io)")
        self.geometry("820x600")
        self.transient(parent)
        self.api_key = api_key
        self.added: list[dict] = []
        self._queue: "queue.Queue" = queue.Queue()
        self._busy = False
        self._row_data: dict[str, dict] = {}  # tree item id -> recipient dict

        pad = {"padx": 8, "pady": 4}
        top = ttk.Frame(self)
        top.pack(fill="x", **pad)
        ttk.Label(
            top,
            text=("Enter company domains or names, one per line "
                  "(e.g. crowdstrike.com, paloaltonetworks.com).\n"
                  "Each company uses ONE of your free Hunter searches."),
            foreground="#444", justify="left",
        ).pack(anchor="w")

        mid = ttk.Frame(self)
        mid.pack(fill="x", **pad)
        self.txt_domains = tk.Text(mid, height=5, wrap="word")
        self.txt_domains.pack(side="left", fill="both", expand=True)
        opts = ttk.Frame(mid)
        opts.pack(side="left", fill="y", padx=6)
        ttk.Label(opts, text="Max per company:").pack(anchor="w")
        self.var_limit = tk.StringVar(value="10")
        ttk.Spinbox(opts, from_=1, to=100, width=6, textvariable=self.var_limit).pack(anchor="w")
        self.var_recruiters_only = tk.BooleanVar(value=False)
        ttk.Checkbutton(opts, text="Recruiter/HR roles only",
                        variable=self.var_recruiters_only).pack(anchor="w", pady=4)
        self.btn_search = ttk.Button(opts, text="Search", command=self._search)
        self.btn_search.pack(anchor="w", fill="x")

        self.lbl_status = ttk.Label(self, text="", foreground="#0969da")
        self.lbl_status.pack(anchor="w", padx=8)

        # Results table
        cols = ("name", "email", "company", "position", "conf")
        self.tree = ttk.Treeview(self, columns=cols, show="headings", selectmode="extended")
        for c, head, w in [("name", "Name", 150), ("email", "Email", 230),
                           ("company", "Company", 140), ("position", "Position", 170),
                           ("conf", "Conf%", 50)]:
            self.tree.heading(c, text=head)
            self.tree.column(c, width=w)
        self.tree.tag_configure("recruiter", background="#e6ffed")
        self.tree.pack(fill="both", expand=True, padx=8, pady=4)

        btns = ttk.Frame(self)
        btns.pack(fill="x", **pad)
        ttk.Button(btns, text="Select all", command=lambda: self.tree.selection_set(self.tree.get_children())).pack(side="left")
        ttk.Button(btns, text="Add selected to recipients", command=self._add_selected).pack(side="left", padx=6)
        ttk.Button(btns, text="Close", command=self.destroy).pack(side="right")
        ttk.Label(btns, text="Tip: green rows look like recruiters/HR.",
                  foreground="#1a7f37").pack(side="right", padx=8)

        self._check_account()
        self.grab_set()

    # -- account quota (shown in the status line) -- #
    def _check_account(self):
        def worker():
            try:
                info = finder.account_info(self.api_key)
                self._queue.put(("account", info))
            except Exception as exc:  # noqa: BLE001
                self._queue.put(("account_err", str(exc)))
        threading.Thread(target=worker, daemon=True).start()
        self.after(100, self._poll)

    def _search(self):
        if self._busy:
            return
        raw = self.txt_domains.get("1.0", "end-1c").splitlines()
        domains = [d.strip() for d in raw if d.strip()]
        if not domains:
            messagebox.showinfo("Find emails", "Type at least one company domain.")
            return
        try:
            limit = max(1, min(100, int(self.var_limit.get())))
        except ValueError:
            limit = 10
        self.tree.delete(*self.tree.get_children())
        self._row_data.clear()
        self._busy = True
        self.btn_search.config(state="disabled")
        self.lbl_status.config(text=f"Searching {len(domains)} company(ies)…")

        recruiters_only = self.var_recruiters_only.get()

        def worker():
            for d in domains:
                try:
                    contacts = finder.domain_search(self.api_key, d, limit=limit)
                    self._queue.put(("found", d, contacts, recruiters_only))
                except Exception as exc:  # noqa: BLE001
                    self._queue.put(("search_err", d, str(exc)))
            self._queue.put(("search_done",))
        threading.Thread(target=worker, daemon=True).start()
        self.after(100, self._poll)

    def _poll(self):
        try:
            while True:
                item = self._queue.get_nowait()
                kind = item[0]
                if kind == "account":
                    info = item[1]
                    self.lbl_status.config(
                        text=f"Hunter plan: {info['plan']} — "
                             f"{info['used']}/{info['available']} searches used.")
                elif kind == "account_err":
                    self.lbl_status.config(text=f"(Could not read Hunter quota: {item[1]})")
                elif kind == "found":
                    _, domain, contacts, recruiters_only = item
                    self._insert_contacts(domain, contacts, recruiters_only)
                elif kind == "search_err":
                    self._insert_error(item[1], item[2])
                elif kind == "search_done":
                    self._busy = False
                    self.btn_search.config(state="normal")
                    n = len(self.tree.get_children())
                    self.lbl_status.config(text=f"Done. {n} contact(s) found. "
                                                "Select the ones you want and click 'Add selected'.")
                    return
        except queue.Empty:
            pass
        if self._busy or not self.tree.get_children():
            self.after(150, self._poll)

    def _insert_contacts(self, domain, contacts, recruiters_only):
        for c in contacts:
            if recruiters_only and not c.is_recruiterish():
                continue
            tags = ("recruiter",) if c.is_recruiterish() else ()
            item = self.tree.insert(
                "", "end",
                values=(c.name, c.email, c.company, c.position, c.confidence),
                tags=tags,
            )
            self._row_data[item] = c.as_recipient()

    def _insert_error(self, domain, msg):
        item = self.tree.insert("", "end", values=("", f"⚠ {domain}: {msg}", "", "", ""))
        # No row data -> can't be added.

    def _add_selected(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Find emails", "Select one or more rows first.")
            return
        for item in sel:
            rec = self._row_data.get(item)
            if rec and rec["email"]:
                self.added.append(rec)
        self.destroy()


if __name__ == "__main__":
    CVEmailer().mainloop()
