"""Build and send application e-mails through Gmail SMTP, then label them."""

from __future__ import annotations

import mimetypes
import smtplib
import ssl
from dataclasses import dataclass, field
from email.headerregistry import Address
from email.message import EmailMessage
from email.utils import formatdate, make_msgid
from pathlib import Path

from common.gmail import GmailClient, GmailError

SMTP_HOST = "smtp.gmail.com"
SMTP_PORT = 465


@dataclass
class OutgoingMail:
    to_address: str
    to_name: str
    subject: str
    body: str
    attachments: list[Path] = field(default_factory=list)
    in_reply_to: str = ""
    references: list[str] = field(default_factory=list)


def build_message(mail: OutgoingMail, from_address: str, from_name: str) -> EmailMessage:
    msg = EmailMessage()
    from_user, from_domain = from_address.split("@", 1)
    to_user, to_domain = mail.to_address.split("@", 1)
    msg["From"] = Address(display_name=from_name, username=from_user, domain=from_domain)
    # The company name as display name lets the watcher name the company without the spreadsheet.
    msg["To"] = Address(display_name=mail.to_name, username=to_user, domain=to_domain)
    msg["Subject"] = mail.subject
    msg["Date"] = formatdate(localtime=True)
    msg["Message-ID"] = make_msgid(domain=from_domain)
    if mail.in_reply_to:
        msg["In-Reply-To"] = mail.in_reply_to
        msg["References"] = " ".join(mail.references or [mail.in_reply_to])
    msg.set_content(mail.body)
    for path in mail.attachments:
        ctype, _ = mimetypes.guess_type(path.name)
        maintype, subtype = (ctype or "application/octet-stream").split("/", 1)
        msg.add_attachment(path.read_bytes(), maintype=maintype, subtype=subtype, filename=path.name)
    return msg


class GmailSender:
    """One SMTP session for a whole batch, plus an IMAP session to label sent mails."""

    def __init__(self, address: str, app_password: str, sender_name: str):
        self.address = address.strip().lower()
        self._password = app_password.replace(" ", "")
        self.sender_name = sender_name
        self._smtp: smtplib.SMTP_SSL | None = None

    def __enter__(self) -> "GmailSender":
        return self

    def __exit__(self, *exc) -> None:
        self.close()

    def _connect(self) -> smtplib.SMTP_SSL:
        if self._smtp is None:
            smtp = smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT, context=ssl.create_default_context(), timeout=60)
            smtp.login(self.address, self._password)
            self._smtp = smtp
        return self._smtp

    def close(self) -> None:
        if self._smtp is not None:
            try:
                self._smtp.quit()
            except Exception:
                pass
            self._smtp = None

    def send(self, mail: OutgoingMail) -> EmailMessage:
        msg = build_message(mail, self.address, self.sender_name)
        try:
            self._connect().send_message(msg)
        except smtplib.SMTPServerDisconnected:
            self._smtp = None  # idle connection dropped by Gmail: retry once
            self._connect().send_message(msg)
        return msg


def test_connection(address: str, app_password: str) -> None:
    """Raise with a readable message if SMTP or IMAP login fails."""
    try:
        smtp = smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT, context=ssl.create_default_context(), timeout=30)
        smtp.login(address, app_password.replace(" ", ""))
        smtp.quit()
    except smtplib.SMTPAuthenticationError as exc:
        raise GmailError("SMTP login refused (check the app password)") from exc
    with GmailClient(address, app_password) as client:
        client.special_folder("\\Sent")
