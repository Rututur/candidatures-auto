"""Spreadsheet adaptation and Gmail synchronisation (needs openpyxl)."""

from __future__ import annotations

import tempfile
import unittest
from datetime import date
from pathlib import Path
from unittest import mock

import openpyxl
from openpyxl.worksheet.datavalidation import DataValidation

from app import sync
from app.templates import render, salutation
from app.workbook import STATUS_REPLIED, ApplicationWorkbook
from tests.fake_gmail import FakeGmail
from tests.test_core import ME, mailbox

HEADERS = ["N°", "Entreprise", "Service visé", "Email", "Interlocuteur", "Mode de contact",
           "Date de contact", "Date de relance", "Statut", "Date de réponse", "Notes"]


def make_sheet(path: Path) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Suivi"
    ws["A1"] = "Suivi des candidatures"
    ws.append([])
    ws.append([])
    ws.append(HEADERS)  # row 4
    rows = [
        ("Cyberscope", "contact@cyberscope.fr", "Email", date(2026, 10, 25), "Envoyée"),
        ("Ceralis", "contact@ceralis.fr", "Email", date(2026, 10, 26), "Envoyée"),
        ("UP2TECH", "contact@up2tech.fr", "", None, "Lettre prête"),
        ("Sans mail", "", "Courrier", None, "Lettre prête"),
    ]
    for i, (company, email, mode, sent, status) in enumerate(rows, start=5):
        ws.append([i - 4, company, "Dev", email, "M. le Président", mode, sent, f'=IF(G{i}="","",G{i}+10)', status, None, None])
    dv = DataValidation(type="list", formula1='"À contacter,Lettre prête,Envoyée,Relancée,Refusée,Sans réponse"')
    dv.add("I5:I20")
    ws.add_data_validation(dv)
    summary = wb.create_sheet("Synthèse")
    for r, status in enumerate(["Envoyée", "Relancée", "Refusée", "Sans réponse"], start=3):
        summary.cell(r, 1, status)
        summary.cell(r, 2, f"=COUNTIF(Suivi!$I$5:$I$20,A{r})")
    summary["A7"], summary["B7"] = "Total", "=SUM(B3:B6)"
    wb.save(path)


class WorkbookTest(unittest.TestCase):
    def setUp(self):
        self.dir = tempfile.TemporaryDirectory()
        self.path = Path(self.dir.name) / "suivi.xlsx"
        make_sheet(self.path)

    def tearDown(self):
        self.dir.cleanup()

    def test_reads_and_upgrades(self):
        wb = ApplicationWorkbook(self.path, reminder_days=7)
        self.assertEqual([r.company for r in wb.rows if r.can_be_sent], ["UP2TECH"])
        self.assertEqual([r.company for r in wb.rows if r.needs_followup(7, date(2026, 11, 1))], ["Cyberscope"])
        wb.mark_sent(wb.by_email("contact@up2tech.fr")[0].row, date(2026, 11, 4), "C:/LM_UP2TECH.pdf")

        book = openpyxl.load_workbook(self.path)
        ws = book["Suivi"]
        headers = [c.value for c in ws[4]]
        self.assertIn("Fichier lettre (PDF)", headers)
        self.assertIn("Relance envoyée le", headers)
        self.assertIn(STATUS_REPLIED, ws.data_validations.dataValidation[0].formula1)
        self.assertEqual(ws["H5"].value, '=IF(G5="","",IF(M5="",G5,M5)+7)')
        summary = book["Synthèse"]
        self.assertEqual(summary["A7"].value, STATUS_REPLIED)
        self.assertEqual(summary["B8"].value, "=SUM(B3:B7)")
        self.assertTrue((self.path.parent / "suivi (sauvegarde auto).xlsx").exists())

    def test_sync_updates_rows(self):
        wb = ApplicationWorkbook(self.path)
        box = mailbox()
        with mock.patch.object(sync, "GmailClient", lambda *a: box):
            report = sync.synchronize(ME, "x", wb, log=lambda m: None)
        statuses = {r.company: (r.status, r.reply_date is not None) for r in wb.rows}
        self.assertEqual(statuses["Cyberscope"], (STATUS_REPLIED, True))
        self.assertEqual(statuses["Ceralis"][0], "Envoyée")  # bounce: note only
        self.assertIn("non distribué", wb.by_email("contact@ceralis.fr")[0].notes)
        self.assertEqual(len(report.updates), 2)

    def test_legacy_applications_get_labelled(self):
        wb = ApplicationWorkbook(self.path)
        from common.gmail import MessageInfo
        from datetime import datetime, timezone
        old = MessageInfo(uid=50, thread_id=5000, gm_msgid=50, labels=set(), from_name="", from_addr=ME,
                          to=[("", "contact@cyberscope.fr")], subject="Candidature",
                          date=datetime(2026, 10, 25, tzinfo=timezone.utc), message_id="<50@t>", in_reply_to="")
        box = FakeGmail(ME, [(old, "")])
        box.search_raw = lambda q: [50] if "to:contact@cyberscope.fr" in q else []
        logs = []
        count = sync.label_legacy_applications(box, wb, {}, logs.append)
        self.assertEqual(count, 1)
        self.assertIn("Candidatures/Envoyees", box.messages[50][0].labels)


class TemplateTest(unittest.TestCase):
    def test_salutation(self):
        self.assertEqual(salutation("M. le Président"), "Monsieur le Président,")
        self.assertEqual(salutation("David Jovanneau (resp. CSN)"), "Bonjour David Jovanneau,")
        self.assertEqual(salutation("Service recrutement"), "Madame, Monsieur,")
        self.assertEqual(salutation(""), "Madame, Monsieur,")

    def test_render_keeps_unknown(self):
        self.assertEqual(render("{entreprise} {inconnu}", {"entreprise": "X"}), "X {inconnu}")


if __name__ == "__main__":
    unittest.main()
