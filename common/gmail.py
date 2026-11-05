"""Thin wrapper around Gmail's IMAP server and its extensions.

Gmail exposes labels, thread ids and its own search syntax over IMAP
(X-GM-LABELS, X-GM-THRID, X-GM-RAW). Those extensions let both tools share
state through Gmail itself: every application e-mail carries a label, and
the watcher marks what it has already notified with other labels.
"""

from __future__ import annotations

import email
import html
import imaplib
import re
import time
from dataclasses import dataclass, field
from datetime import datetime, timezone
from email.header import decode_header, make_header
from email.message import Message
from email.utils import getaddresses, parseaddr, parsedate_to_datetime

GMAIL_IMAP_HOST = "imap.gmail.com"

# Labels shared by the desktop app and the watcher (ASCII only, so no IMAP
# modified UTF-7 encoding is needed). Nested under one parent in Gmail.
LABEL_ROOT = "Candidatures"
LABEL_SENT = "Candidatures/Envoyees"
LABEL_NOTIFIED = "Candidatures/Notifiees"
LABEL_AUTO_REPLY = "Candidatures/Reponses-auto"
LABEL_REMINDED = "Candidatures/Rappel-fait"
ALL_LABELS = (LABEL_ROOT, LABEL_SENT, LABEL_NOTIFIED, LABEL_AUTO_REPLY, LABEL_REMINDED)

HEADER_FIELDS = (
    "FROM TO SUBJECT DATE MESSAGE-ID IN-REPLY-TO REFERENCES "
    "AUTO-SUBMITTED X-AUTOREPLY X-AUTORESPOND PRECEDENCE"
)
FETCH_CHUNK = 150
BODY_FETCH_LIMIT = 300_000  # bytes; enough for the text part, skips big attachments


class GmailError(Exception):
    """Raised when Gmail refuses a command or the login fails."""


class FolderNotFound(GmailError):
    """Raised when selecting a folder or label that does not exist yet."""


@dataclass
class MessageInfo:
    uid: int
    thread_id: int
    gm_msgid: int
    labels: set[str]
    from_name: str
    from_addr: str
    to: list[tuple[str, str]]
    subject: str
    date: datetime | None
    message_id: str
    in_reply_to: str
    references: list[str] = field(default_factory=list)
    auto_submitted: bool = False


def decode_mime_header(value: str | None) -> str:
    if not value:
        return ""
    try:
        return str(make_header(decode_header(value)))
    except Exception:  # malformed encodings: keep the raw value
        return value


def quote_imap(value: str) -> str:
    """Quote a string argument for an IMAP command."""
    return '"' + value.replace("\\", "\\\\").replace('"', '\\"') + '"'


def _parse_labels(meta: str) -> set[str]:
    match = re.search(r"X-GM-LABELS \(([^)]*)\)", meta)
    if not match:
        return set()
    tokens = re.findall(r'"((?:[^"\\]|\\.)*)"|(\S+)', match.group(1))
    return {quoted.replace('\\"', '"').replace("\\\\", "\\") or bare for quoted, bare in tokens}


def _parse_int(meta: str, key: str) -> int:
    match = re.search(key + r" (\d+)", meta)
    return int(match.group(1)) if match else 0


def _parse_date(value: str | None) -> datetime | None:
    if not value:
        return None
    try:
        parsed = parsedate_to_datetime(value)
    except (TypeError, ValueError):
        return None
    if parsed.tzinfo is None:
        parsed = parsed.replace(tzinfo=timezone.utc)
    return parsed


def _is_auto_submitted(msg: Message) -> bool:
    auto = (msg.get("Auto-Submitted") or "").strip().lower()
    if auto and auto != "no":
        return True
    if msg.get("X-Autoreply") or msg.get("X-Autorespond"):
        return True
    return (msg.get("Precedence") or "").strip().lower() in {"auto_reply", "bulk", "junk"}


def parse_header_block(uid: int, meta: str, raw_headers: bytes) -> MessageInfo:
    msg = email.message_from_bytes(raw_headers)
    from_name, from_addr = parseaddr(decode_mime_header(msg.get("From")))
    to = [
        (decode_mime_header(name), addr.lower())
        for name, addr in getaddresses([msg.get("To", ""), msg.get("Cc", "")])
        if addr
    ]
    references = re.findall(r"<[^>]+>", msg.get("References", ""))
    return MessageInfo(
        uid=uid,
        thread_id=_parse_int(meta, "X-GM-THRID"),
        gm_msgid=_parse_int(meta, "X-GM-MSGID"),
        labels=_parse_labels(meta),
        from_name=from_name,
        from_addr=from_addr.lower(),
        to=to,
        subject=decode_mime_header(msg.get("Subject")),
        date=_parse_date(msg.get("Date")),
        message_id=(msg.get("Message-ID") or "").strip(),
        in_reply_to=(msg.get("In-Reply-To") or "").strip(),
        references=references,
        auto_submitted=_is_auto_submitted(msg),
    )


def _html_to_text(markup: str) -> str:
    markup = re.sub(r"(?is)<(script|style).*?</\1>", " ", markup)
    markup = re.sub(r"(?i)<br\s*/?>|</p>|</div>|</tr>", "\n", markup)
    markup = re.sub(r"(?s)<[^>]+>", " ", markup)
    return html.unescape(markup)


def extract_text(msg: Message) -> str:
    """Return the readable text of an e-mail, preferring the plain-text part."""
    plain, rich = [], []
    for part in msg.walk() if msg.is_multipart() else [msg]:
        if part.get_content_maintype() == "multipart" or part.get_filename():
            continue
        ctype = part.get_content_type()
        if ctype not in ("text/plain", "text/html"):
            continue
        payload = part.get_payload(decode=True) or b""
        charset = part.get_content_charset() or "utf-8"
        try:
            text = payload.decode(charset, errors="replace")
        except LookupError:
            text = payload.decode("utf-8", errors="replace")
        (plain if ctype == "text/plain" else rich).append(text)
    if plain:
        return "\n".join(plain)
    return "\n".join(_html_to_text(t) for t in rich)


class GmailClient:
    """Context manager holding one authenticated IMAP connection."""

    def __init__(self, address: str, app_password: str, host: str = GMAIL_IMAP_HOST):
        self.address = address.strip().lower()
        self._password = app_password.replace(" ", "")
        self._host = host
        self.imap: imaplib.IMAP4_SSL | None = None
        self._special: dict[str, str] | None = None
        self.selected: str | None = None

    # -- connection -----------------------------------------------------
    def __enter__(self) -> "GmailClient":
        self.connect()
        return self

    def __exit__(self, *exc) -> None:
        self.close()

    def connect(self) -> None:
        try:
            self.imap = imaplib.IMAP4_SSL(self._host, timeout=60)
            self.imap.login(self.address, self._password)
        except imaplib.IMAP4.error as exc:
            raise GmailError(f"IMAP login failed: {exc}") from exc

    def close(self) -> None:
        if self.imap is None:
            return
        try:
            if self.selected:
                self.imap.close()
            self.imap.logout()
        except Exception:
            pass
        self.imap = None

    def _ok(self, result: tuple, what: str) -> list:
        typ, data = result
        if typ != "OK":
            raise GmailError(f"{what} failed: {data!r}")
        return data

    # -- folders and labels ----------------------------------------------
    def special_folder(self, flag: str) -> str:
        """Name of a special-use folder such as \\Sent or \\All (localised by Gmail)."""
        if self._special is None:
            self._special = {}
            for line in self._ok(self.imap.list(), "LIST"):
                if not line:
                    continue
                text = line.decode("utf-8", errors="replace") if isinstance(line, bytes) else str(line)
                match = re.match(r'\((?P<flags>[^)]*)\) (?:"[^"]*"|NIL) (?P<name>.+)$', text)
                if not match:
                    continue
                name = match.group("name").strip()
                if name.startswith('"') and name.endswith('"'):
                    name = name[1:-1]
                for item in match.group("flags").split():
                    self._special[item.lower()] = name
        try:
            return self._special[flag.lower()]
        except KeyError:
            raise GmailError(f"Special folder {flag} not found (is IMAP enabled?)") from None

    def ensure_labels(self, labels=ALL_LABELS) -> None:
        for label in labels:
            typ, data = self.imap.create(quote_imap(label))
            if typ != "OK" and b"ALREADYEXISTS" not in b" ".join(d or b"" for d in data).upper():
                # Gmail answers NO when the label already exists; anything else is a real error.
                if b"EXIST" not in b" ".join(d or b"" for d in data).upper():
                    raise GmailError(f"Could not create label {label}: {data!r}")

    def select(self, folder: str, readonly: bool = True) -> int:
        typ, data = self.imap.select(quote_imap(folder), readonly=readonly)
        if typ != "OK":
            missing = b"NONEXISTENT" in b" ".join(d or b"" for d in data).upper()
            raise (FolderNotFound if missing else GmailError)(f"Could not open folder {folder}: {data!r}")
        self.selected = folder
        return int(data[0] or 0)

    # -- search ---------------------------------------------------------
    def _uid_search(self, *criteria: str) -> list[int]:
        data = self._ok(self.imap.uid("SEARCH", *criteria), "SEARCH")
        return [int(x) for x in (data[0] or b"").split()]

    def search_all(self) -> list[int]:
        return self._uid_search("ALL")

    def search_raw(self, query: str) -> list[int]:
        """Search with Gmail's web search syntax (from:, to:, after:, label: ...)."""
        return self._uid_search("X-GM-RAW", quote_imap(query))

    def search_thread(self, thread_id: int) -> list[int]:
        return self._uid_search("X-GM-THRID", str(thread_id))

    def find_by_message_id(self, message_id: str, attempts: int = 6, wait: float = 2.0) -> list[int]:
        """Find a freshly sent message; Gmail may need a few seconds to index it."""
        bare = message_id.strip().strip("<>")
        for attempt in range(attempts):
            uids = self.search_raw(f"rfc822msgid:{bare}")
            if not uids:
                uids = self._uid_search("HEADER", "Message-ID", quote_imap(message_id))
            if uids:
                return uids
            if attempt < attempts - 1:
                time.sleep(wait)
                self.imap.noop()
        return []

    # -- fetch ----------------------------------------------------------
    def fetch_headers(self, uids) -> list[MessageInfo]:
        uids = sorted(set(uids))
        infos: list[MessageInfo] = []
        for start in range(0, len(uids), FETCH_CHUNK):
            chunk = ",".join(str(u) for u in uids[start : start + FETCH_CHUNK])
            data = self._ok(
                self.imap.uid(
                    "FETCH",
                    chunk,
                    f"(UID X-GM-THRID X-GM-MSGID X-GM-LABELS BODY.PEEK[HEADER.FIELDS ({HEADER_FIELDS})])",
                ),
                "FETCH",
            )
            records: list[list] = []
            for item in data:
                if isinstance(item, tuple):
                    records.append([item[0].decode("utf-8", errors="replace"), item[1]])
                elif isinstance(item, bytes) and records:
                    # Attributes sent after the literal end up in a trailing chunk.
                    records[-1][0] += " " + item.decode("utf-8", errors="replace")
            for meta, raw in records:
                uid = _parse_int(meta, "UID")
                if uid:
                    infos.append(parse_header_block(uid, meta, raw or b""))
        return infos

    def fetch_message(self, uid: int) -> Message:
        data = self._ok(self.imap.uid("FETCH", str(uid), f"(BODY.PEEK[]<0.{BODY_FETCH_LIMIT}>)"), "FETCH")
        for item in data:
            if isinstance(item, tuple):
                return email.message_from_bytes(item[1])
        raise GmailError(f"Message {uid} not found")

    def fetch_text(self, uid: int) -> str:
        return extract_text(self.fetch_message(uid))

    # -- modify ---------------------------------------------------------
    def add_label(self, uids, label: str) -> None:
        uids = sorted(set(uids))
        if not uids:
            return
        self._ok(
            self.imap.uid("STORE", ",".join(map(str, uids)), "+X-GM-LABELS", f"({quote_imap(label)})"),
            "STORE",
        )


def gmail_thread_url(account: str, thread_id: int) -> str:
    return f"https://mail.google.com/mail/?authuser={account}#all/{thread_id:x}"
