"""Synchronise the spreadsheet with Gmail (run when the app opens and on demand)."""

from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass, field
from datetime import date, timedelta
from typing import Callable

from common.classify import KIND_AUTO, KIND_BOUNCE, classify
from common.gmail import LABEL_AUTO_REPLY, LABEL_SENT, GmailClient
from common.tracking import SentApplication, find_replies, load_applications

from .workbook import SENDABLE_STATUSES, WAITING_STATUSES, ApplicationWorkbook

Log = Callable[[str], None]


@dataclass
class SyncReport:
    applications: dict[str, SentApplication] = field(default_factory=dict)
    legacy_labelled: int = 0
    replies_found: int = 0
    updates: list[str] = field(default_factory=list)


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


_kind_cache: dict[tuple[int, int], str] = {}


def synchronize(address: str, password: str, workbook: ApplicationWorkbook, log: Log) -> SyncReport:
    report = SyncReport()
    with GmailClient(address, password) as client:
        client.ensure_labels()
        applications = load_applications(client)
        report.legacy_labelled = label_legacy_applications(client, workbook, applications, log)
        if report.legacy_labelled:
            applications = load_applications(client)
        report.applications = applications

        replies = find_replies(client, applications)
        report.replies_found = len(replies)
        grouped = defaultdict(list)
        for reply in replies:
            grouped[reply.application.address].append(reply)

        for address_, items in grouped.items():
            for row in workbook.by_email(address_):
                if row.reply_date is not None or row.status not in WAITING_STATUSES | SENDABLE_STATUSES:
                    continue
                for reply in items:
                    msg = reply.message
                    received = msg.date.astimezone().date() if msg.date else date.today()
                    if row.contact_date and received < row.contact_date:
                        continue
                    key = (msg.gm_msgid, msg.uid)
                    if LABEL_AUTO_REPLY in msg.labels:
                        kind = KIND_AUTO
                    elif key in _kind_cache:
                        kind = _kind_cache[key]
                    else:
                        text = client.fetch_text(msg.uid)
                        kind = _kind_cache[key] = classify(text, msg.from_addr, msg.subject, msg.auto_submitted).kind
                    result = workbook.apply_reply(row.row, kind, received, msg.subject)
                    if result:
                        report.updates.append(f"{row.company} : {result}")
                        log(f"Tableur mis à jour — {row.company} : {result}")
                        row = workbook.by_row(row.row)
                    if kind not in (KIND_AUTO, KIND_BOUNCE) and result:
                        break
    return report
