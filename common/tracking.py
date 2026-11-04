"""Which applications were sent, and which received messages answer them.

Source of truth: the Gmail label LABEL_SENT. The desktop app puts it on every
application it sends (and on older applications found in the Sent folder),
so the watcher does not need access to the spreadsheet.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime, timedelta, timezone

from .gmail import LABEL_SENT, GmailClient, MessageInfo

# Free mailbox providers: matching a reply by domain would be meaningless.
GENERIC_DOMAINS = {
    "gmail.com", "googlemail.com", "outlook.com", "outlook.fr", "hotmail.com", "hotmail.fr",
    "live.com", "live.fr", "msn.com", "yahoo.com", "yahoo.fr", "icloud.com", "me.com",
    "orange.fr", "wanadoo.fr", "free.fr", "sfr.fr", "neuf.fr", "laposte.net", "bbox.fr",
    "aol.com", "gmx.fr", "gmx.com", "protonmail.com", "proton.me",
}
SEARCH_TERMS_PER_QUERY = 25
_EARLIEST = datetime.min.replace(tzinfo=timezone.utc)
_LATEST = datetime.max.replace(tzinfo=timezone.utc)


def domain_of(address: str) -> str:
    return address.rsplit("@", 1)[-1].lower() if "@" in address else ""


@dataclass
class SentApplication:
    """Everything sent to one recipient address (first mail plus follow-ups)."""

    address: str
    display_name: str = ""
    messages: list[MessageInfo] = field(default_factory=list)

    @property
    def domain(self) -> str:
        return domain_of(self.address)

    @property
    def matches_by_domain(self) -> bool:
        return self.domain not in GENERIC_DOMAINS

    @property
    def first(self) -> MessageInfo:
        return min(self.messages, key=lambda m: m.date or _LATEST)

    @property
    def last(self) -> MessageInfo:
        return max(self.messages, key=lambda m: m.date or _EARLIEST)

    @property
    def thread_ids(self) -> set[int]:
        return {m.thread_id for m in self.messages if m.thread_id}

    @property
    def message_ids(self) -> list[str]:
        ordered = sorted(self.messages, key=lambda m: m.date or _EARLIEST)
        return [m.message_id for m in ordered if m.message_id]

    @property
    def name(self) -> str:
        """Best human name for the company: the To display name, else the domain."""
        return self.display_name or self.domain or self.address


@dataclass
class Reply:
    application: SentApplication
    message: MessageInfo
    matched_by: str  # "thread", "address" or "domain"


def load_applications(client: GmailClient) -> dict[str, SentApplication]:
    """Group every message carrying LABEL_SENT by recipient address."""
    client.select(LABEL_SENT)
    infos = client.fetch_headers(client.search_all())
    applications: dict[str, SentApplication] = {}
    for info in infos:
        for name, addr in info.to:
            if not addr or addr == client.address:
                continue
            app = applications.setdefault(addr, SentApplication(address=addr))
            app.messages.append(info)
            if name and not app.display_name:
                app.display_name = name
    return applications


def _match(info: MessageInfo, applications: dict[str, SentApplication], by_thread: dict[int, SentApplication]):
    app = by_thread.get(info.thread_id)
    if app:
        return app, "thread"
    app = applications.get(info.from_addr)
    if app:
        return app, "address"
    domain = domain_of(info.from_addr)
    for candidate in applications.values():
        if candidate.matches_by_domain and candidate.domain == domain:
            return candidate, "domain"
    return None, ""


def find_replies(client: GmailClient, applications: dict[str, SentApplication]) -> list[Reply]:
    """Find received messages that answer one of the sent applications.

    Candidates come from two Gmail searches: messages in the same threads as
    the applications, and messages from the recipients' addresses/domains.
    """
    if not applications:
        return []
    client.select(client.special_folder("\\All"))

    dates = [m.date for app in applications.values() for m in app.messages if m.date]
    earliest = (min(dates) - timedelta(days=1)) if dates else None
    after = f"after:{earliest:%Y/%m/%d} " if earliest else ""

    terms = sorted({app.domain if app.matches_by_domain else app.address for app in applications.values()})
    uids: set[int] = set()
    for start in range(0, len(terms), SEARCH_TERMS_PER_QUERY):
        chunk = terms[start : start + SEARCH_TERMS_PER_QUERY]
        uids.update(client.search_raw(after + "-from:me {" + " ".join(f"from:{t}" for t in chunk) + "}"))

    by_thread: dict[int, SentApplication] = {}
    for app in applications.values():
        for thread_id in app.thread_ids:
            by_thread[thread_id] = app
            uids.update(client.search_thread(thread_id))

    replies: list[Reply] = []
    for info in client.fetch_headers(uids):
        if not info.from_addr or info.from_addr == client.address:
            continue
        app, matched_by = _match(info, applications, by_thread)
        if app is None:
            continue
        first_date = app.first.date
        if first_date and info.date and info.date <= first_date:
            continue  # older correspondence with that company, not an answer
        replies.append(Reply(application=app, message=info, matched_by=matched_by))
    replies.sort(key=lambda r: r.message.date or _EARLIEST)
    return replies
