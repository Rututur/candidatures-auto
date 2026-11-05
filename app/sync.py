"""Synchronise with Gmail (run when the app opens and on demand).

Replies never change the spreadsheet on their own: they are listed with an
estimate, and only the user's decision updates the sheet. A decided mail gets
LABEL_VALIDATED, so it is not proposed again.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, timedelta
from typing import Callable

from common.classify import KIND_AUTO, KIND_BOUNCE, KIND_POSITIVE, KIND_REFUSAL, KIND_UNKNOWN, classify
from common.gmail import LABEL_AUTO_REPLY, LABEL_SENT, LABEL_VALIDATED, GmailClient
from common.tracking import Reply, SentApplication, find_replies, load_applications

from .workbook import (
    DECISION_BOUNCE,
    DECISION_IGNORE,
    STATUS_REFUSED,
    STATUS_REPLIED,
    WAITING_STATUSES,
    ApplicationWorkbook,
)

Log = Callable[[str], None]

SUGGESTED_DECISIONS = {
    KIND_REFUSAL: STATUS_REFUSED,
    KIND_POSITIVE: STATUS_REPLIED,
    KIND_UNKNOWN: STATUS_REPLIED,
    KIND_AUTO: DECISION_IGNORE,
    KIND_BOUNCE: DECISION_BOUNCE,
}


@dataclass
class PendingReply:
    """A reply waiting for the user's decision."""

    uid: int  # in the All Mail folder
    thread_id: int
    address: str  # the application's address, which links the reply to the sheet
    company: str
    sender: str
    subject: str
    received: date
    text: str
    kind: str

    @property
    def suggestion(self) -> str:
        return SUGGESTED_DECISIONS[self.kind]


@dataclass
class SyncReport:
    applications: dict[str, SentApplication] = field(default_factory=dict)
    legacy_labelled: int = 0
    replies_found: int = 0
    pending: list[PendingReply] = field(default_factory=list)


class SentLabeler:
    """Keeps the Sent folder open to label each application right after sending it."""

    def __init__(self, client: GmailClient):
        self.client = client
        client.ensure_labels()
        client.select(client.special_folder("\\Sent"), readonly=False)

    def label(self, message_id: str) -> bool:
        uids = self.client.find_by_message_id(message_id)
        if uids:
            self.client.add_label(uids, LABEL_SENT)
        return bool(uids)


def label_legacy_applications(client: GmailClient, workbook: ApplicationWorkbook,
                              applications: dict[str, SentApplication], log: Log) -> int:
    """Label applications sent before the tool existed (status Envoyée/Relancée in the sheet)."""
    rows = [
        r for r in workbook.rows
        if r.status in WAITING_STATUSES and r.has_valid_email and r.email not in applications
    ]
    if not rows:
        return 0
    client.select(client.special_folder("\\Sent"), readonly=False)
    count = 0
    for row in rows:
        query = f"to:{row.email}"
        if row.contact_date:
            query += f" after:{row.contact_date - timedelta(days=1):%Y/%m/%d}"
        uids = client.search_raw(query)
        if uids:
            client.add_label(uids, LABEL_SENT)
            count += 1
            log(f"Candidature déjà envoyée retrouvée dans Gmail : {row.company}")
        else:
            log(f"Aucun mail envoyé trouvé dans Gmail pour {row.company} ({row.email}) : pas de suivi possible pour celle-ci.")
    return count


def _pending_reply(client: GmailClient, workbook: ApplicationWorkbook, reply: Reply) -> PendingReply:
    msg = reply.message
    address = reply.application.address
    rows = workbook.by_email(address)
    text = client.fetch_text(msg.uid)
    return PendingReply(
        uid=msg.uid,
        thread_id=msg.thread_id,
        address=address,
        company=rows[0].company if rows else reply.application.name,
        sender=f"{msg.from_name} <{msg.from_addr}>" if msg.from_name else msg.from_addr,
        subject=msg.subject,
        received=msg.date.astimezone().date() if msg.date else date.today(),
        text=text,
        kind=classify(text, msg.from_addr, msg.subject, msg.auto_submitted).kind,
    )


def synchronize(address: str, password: str, workbook: ApplicationWorkbook, log: Log) -> SyncReport:
    report = SyncReport()
    with GmailClient(address, password) as client:
        client.ensure_labels()
        applications = load_applications(client)
        report.legacy_labelled = label_legacy_applications(client, workbook, applications, log)
        if report.legacy_labelled:
            applications = load_applications(client)
        report.applications = applications

        replies = find_replies(client, applications)  # leaves All Mail selected
        report.replies_found = len(replies)
        report.pending = [
            _pending_reply(client, workbook, reply)
            for reply in replies
            if LABEL_VALIDATED not in reply.message.labels
        ]
    return report


def record_decision(address: str, password: str, workbook: ApplicationWorkbook,
                    reply: PendingReply, decision: str) -> None:
    """Update the sheet, then label the mail in Gmail so it is not proposed again."""
    workbook.check_writable()
    for row in workbook.by_email(reply.address):
        workbook.apply_decision(row.row, decision, reply.received, reply.subject)
    with GmailClient(address, password) as client:
        client.select(client.special_folder("\\All"), readonly=False)
        client.add_label([reply.uid], LABEL_VALIDATED)
        # The watcher keeps sending follow-up reminders for mails labelled as automatic.
        if decision == DECISION_IGNORE:
            client.add_label([reply.uid], LABEL_AUTO_REPLY)
        else:
            client.remove_label([reply.uid], LABEL_AUTO_REPLY)
