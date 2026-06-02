"""Email building and SMTP sending for CV Emailer.

Each recipient gets their *own* individual message (their address only appears
in the To: header), so recipients never see one another and the {name} /
{company} placeholders can be personalised per person.
"""

from __future__ import annotations

import mimetypes
import smtplib
import ssl
import time
from dataclasses import dataclass, field
from email.message import EmailMessage
from email.utils import formataddr
from pathlib import Path


@dataclass
class SmtpConfig:
    host: str
    port: int
    security: str  # "starttls" | "ssl" | "none"
    username: str
    password: str
    from_name: str = ""

    @property
    def from_address(self) -> str:
        return self.username


def personalise(text: str, recipient: dict) -> str:
    """Replace {name}, {company}, {email} placeholders for one recipient.

    A blank name falls back to a neutral greeting word so emails never read
    "Dear ,". Unknown placeholders are left untouched.
    """
    name = (recipient.get("name") or "").strip()
    mapping = {
        "name": name or "there",
        "first_name": (name.split()[0] if name else "there"),
        "company": (recipient.get("company") or "").strip(),
        "email": (recipient.get("email") or "").strip(),
    }
    out = text
    for key, value in mapping.items():
        out = out.replace("{" + key + "}", value)
    return out


def build_message(cfg: SmtpConfig, recipient: dict, subject: str,
                  body: str, attachments: list[str]) -> EmailMessage:
    msg = EmailMessage()
    msg["From"] = formataddr((cfg.from_name or None, cfg.from_address))
    msg["To"] = formataddr((recipient.get("name") or None, recipient["email"]))
    msg["Subject"] = personalise(subject, recipient)
    msg.set_content(personalise(body, recipient))

    for path_str in attachments:
        path = Path(path_str)
        if not path.is_file():
            raise FileNotFoundError(f"Attachment not found: {path}")
        ctype, encoding = mimetypes.guess_type(path.name)
        if ctype is None or encoding is not None:
            ctype = "application/octet-stream"
        maintype, subtype = ctype.split("/", 1)
        with open(path, "rb") as fh:
            msg.add_attachment(
                fh.read(), maintype=maintype, subtype=subtype, filename=path.name
            )
    return msg


def _connect(cfg: SmtpConfig) -> smtplib.SMTP:
    context = ssl.create_default_context()
    if cfg.security == "ssl":
        server = smtplib.SMTP_SSL(cfg.host, cfg.port, timeout=30, context=context)
    else:
        server = smtplib.SMTP(cfg.host, cfg.port, timeout=30)
        server.ehlo()
        if cfg.security == "starttls":
            server.starttls(context=context)
            server.ehlo()
    if cfg.username:
        server.login(cfg.username, cfg.password)
    return server


def verify_connection(cfg: SmtpConfig) -> None:
    """Open + authenticate a connection, then close. Raises on failure."""
    server = _connect(cfg)
    try:
        server.noop()
    finally:
        try:
            server.quit()
        except Exception:
            server.close()


@dataclass
class SendReport:
    sent: list[str] = field(default_factory=list)
    failed: list[tuple[str, str]] = field(default_factory=list)  # (email, error)


def friendly_smtp_error(exc: Exception) -> str:
    """Turn raw SMTP exceptions into guidance, esp. the M365 auth case."""
    if isinstance(exc, smtplib.SMTPAuthenticationError):
        return (
            "Login was rejected by the mail server.\n\n"
            "If this is a university / Microsoft 365 address, the most likely "
            "cause is that 'SMTP basic authentication' is disabled on the "
            "account. Options:\n"
            "  - Ask IT to enable SMTP AUTH for your mailbox, or\n"
            "  - Use a personal Gmail with an App Password instead "
            "(Settings -> Gmail preset).\n\n"
            f"Server said: {exc}"
        )
    if isinstance(exc, smtplib.SMTPConnectError):
        return f"Could not connect to the server. Check host/port.\n\n{exc}"
    if isinstance(exc, (TimeoutError, OSError)):
        return f"Network/connection problem reaching the server.\n\n{exc}"
    return str(exc)


def send_bulk(cfg: SmtpConfig, recipients: list[dict], subject: str, body: str,
              attachments: list[str], delay_seconds: float = 2.0,
              progress=None, should_stop=None) -> SendReport:
    """Send one personalised email per recipient over a single connection.

    ``progress(index, total, email, status, detail)`` is called for each
    recipient; ``status`` is "sent" or "failed". ``should_stop()`` is polled
    between sends so the UI can cancel.
    """
    report = SendReport()
    total = len(recipients)
    server = _connect(cfg)
    try:
        for i, recipient in enumerate(recipients, start=1):
            if should_stop and should_stop():
                break
            email = recipient.get("email", "").strip()
            try:
                msg = build_message(cfg, recipient, subject, body, attachments)
                server.send_message(msg)
                report.sent.append(email)
                if progress:
                    progress(i, total, email, "sent", "")
            except smtplib.SMTPServerDisconnected:
                # Connection dropped mid-run: reconnect once and retry this one.
                server = _connect(cfg)
                try:
                    server.send_message(msg)
                    report.sent.append(email)
                    if progress:
                        progress(i, total, email, "sent", "(reconnected)")
                except Exception as exc:  # noqa: BLE001
                    report.failed.append((email, str(exc)))
                    if progress:
                        progress(i, total, email, "failed", str(exc))
            except Exception as exc:  # noqa: BLE001 - report and keep going
                report.failed.append((email, str(exc)))
                if progress:
                    progress(i, total, email, "failed", str(exc))
            if delay_seconds and i < total:
                time.sleep(delay_seconds)
    finally:
        try:
            server.quit()
        except Exception:
            server.close()
    return report
