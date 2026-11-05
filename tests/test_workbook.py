"""Spreadsheet adaptation and Gmail synchronisation (needs openpyxl)."""

from __future__ import annotations

import os
import tempfile
import unittest
from datetime import date
from pathlib import Path
from unittest import mock

import openpyxl
from openpyxl.worksheet.datavalidation import DataValidation

from app import sync
from app.templates import render, salutation
from app.workbook import (
    DECISION_BOUNCE,
    DECISION_IGNORE,
    STATUS_INTERVIEW,
    STATUS_REFUSED,
    STATUS_REPLIED,
    ApplicationWorkbook,
    WorkbookChanged,
    WorkbookError,
)
from common.gmail import LABEL_AUTO_REPLY, LABEL_VALIDATED
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

    def synchronize(self, wb, box):
        with mock.patch.object(sync, "GmailClient", lambda *a: box):
            return sync.synchronize(ME, "x", wb, log=lambda m: None)

    def decide(self, wb, box, reply, decision):
        with mock.patch.object(sync, "GmailClient", lambda *a: box):
            sync.record_decision(ME, "x", wb, reply, decision)

    def test_sync_proposes_replies_without_changing_the_sheet(self):
        wb = ApplicationWorkbook(self.path)
        before = self.path.read_bytes()
        report = self.synchronize(wb, mailbox())
        self.assertEqual(self.path.read_bytes(), before)
        self.assertEqual({r.uid: (r.company, r.suggestion) for r in report.pending}, {
            10: ("Cyberscope", STATUS_REPLIED),
            11: ("Ceralis", DECISION_BOUNCE),
            12: ("Evolusite", DECISION_IGNORE),
        })

    def test_correcting_the_estimate_updates_the_sheet_and_the_labels(self):
        wb = ApplicationWorkbook(self.path)
        box = mailbox()
        box.messages[10][0].labels.add(LABEL_AUTO_REPLY)
        reply = next(r for r in self.synchronize(wb, box).pending if r.uid == 10)
        self.assertEqual(reply.suggestion, STATUS_REPLIED)

        self.decide(wb, box, reply, STATUS_REFUSED)

        row = wb.by_email("contact@cyberscope.fr")[0]
        self.assertEqual((row.status, row.reply_date), (STATUS_REFUSED, reply.received))
        self.assertIn("[Validé", row.notes)
        self.assertIn(LABEL_VALIDATED, box.messages[10][0].labels)
        self.assertNotIn(LABEL_AUTO_REPLY, box.messages[10][0].labels)
        self.assertNotIn(10, [r.uid for r in self.synchronize(wb, box).pending])

    def test_second_decision_keeps_the_first_reply_date(self):
        wb = ApplicationWorkbook(self.path)
        row = wb.by_email("contact@cyberscope.fr")[0].row
        wb.apply_decision(row, STATUS_REPLIED, date(2026, 11, 1), "Premier mail")
        wb.apply_decision(row, STATUS_INTERVIEW, date(2026, 11, 3), "Second mail")
        item = wb.by_row(row)
        self.assertEqual((item.status, item.reply_date), (STATUS_INTERVIEW, date(2026, 11, 1)))

    def test_bounce_only_adds_a_note(self):
        wb = ApplicationWorkbook(self.path)
        row = wb.by_email("contact@ceralis.fr")[0].row
        wb.apply_decision(row, DECISION_BOUNCE, date(2026, 11, 1), "Échec")
        item = wb.by_row(row)
        self.assertEqual((item.status, item.reply_date), ("Envoyée", None))
        self.assertIn("non distribué", item.notes)

    def test_ignoring_a_reply_keeps_reminders_and_the_sheet(self):
        wb = ApplicationWorkbook(self.path)
        box = mailbox()
        reply = next(r for r in self.synchronize(wb, box).pending if r.uid == 10)
        before = self.path.read_bytes()
        self.decide(wb, box, reply, DECISION_IGNORE)
        self.assertEqual(self.path.read_bytes(), before)
        self.assertTrue({LABEL_VALIDATED, LABEL_AUTO_REPLY} <= box.messages[10][0].labels)

    def test_ignoring_undoes_an_estimate_applied_by_an_older_version(self):
        wb = ApplicationWorkbook(self.path)
        row = wb.by_email("contact@cyberscope.fr")[0].row
        when = date(2026, 11, 1)
        wb.update(row, {"status": STATUS_REFUSED, "reply_date": when,
                        "notes": "[Auto 01/11/2026] Réponse reçue, estimation : refus probable. À vérifier."})
        wb.apply_decision(row, DECISION_IGNORE, when, "Accusé")
        item = wb.by_row(row)
        self.assertEqual((item.status, item.reply_date), ("Envoyée", None))
        self.assertIn("Estimation annulée", item.notes)

    def column(self, wb, title):
        return next(c for c in wb.table_columns if c.title == title)

    def test_table_shows_every_column_and_computes_known_formulas(self):
        book = openpyxl.load_workbook(self.path)
        ws = book["Suivi"]
        ws.cell(4, 12, "Action")
        ws.cell(5, 12, '=IF(I5="Envoyée","À relancer","")')
        ws.cell(6, 12, '=IF(I6="Envoyée","À relancer","")')
        book.save(self.path)

        wb = ApplicationWorkbook(self.path, reminder_days=7)
        self.assertEqual([c.title for c in wb.table_columns], HEADERS + ["Action"])
        row = wb.by_email("contact@cyberscope.fr")[0].row
        show = lambda title: wb.display_value(row, self.column(wb, title))  # noqa: E731
        self.assertEqual(show("Entreprise"), "Cyberscope")
        self.assertEqual(show("Date de contact"), "25/10/2026")
        self.assertEqual(show("Date de relance"), "01/11/2026")  # computed: no cached formula result
        self.assertEqual(show("Action"), "À relancer")
        self.assertIn("Envoyée", self.column(wb, "Statut").choices)

    def test_edit_cell_replaces_the_value_in_the_file(self):
        wb = ApplicationWorkbook(self.path)
        row = wb.by_email("contact@cyberscope.fr")[0].row
        wb.edit_cell(row, self.column(wb, "Notes"), "Première note")
        wb.edit_cell(row, self.column(wb, "Notes"), "Note corrigée")
        wb.edit_cell(row, self.column(wb, "Statut"), "Refusée")
        wb.edit_cell(row, self.column(wb, "Date de réponse"), "03/11/2026")
        wb.edit_cell(row, self.column(wb, "N°"), "42")

        ws = openpyxl.load_workbook(self.path)["Suivi"]
        self.assertEqual(ws.cell(row, 11).value, "Note corrigée")
        self.assertEqual(ws.cell(row, 9).value, "Refusée")
        self.assertEqual(ws.cell(row, 10).value.date(), date(2026, 11, 3))
        self.assertEqual(ws.cell(row, 10).number_format, "dd/mm/yyyy")
        self.assertEqual(ws.cell(row, 1).value, 42)
        self.assertEqual(ws.cell(row, 8).value, f'=IF(G{row}="","",G{row}+10)')  # formula untouched
        self.assertEqual(wb.by_row(row).status, "Refusée")

    def test_edit_cell_can_empty_a_cell(self):
        wb = ApplicationWorkbook(self.path)
        row = wb.by_email("contact@cyberscope.fr")[0].row
        wb.edit_cell(row, self.column(wb, "Date de contact"), "")
        self.assertIsNone(wb.by_row(row).contact_date)

    def test_edit_cell_refuses_invalid_input_without_writing(self):
        wb = ApplicationWorkbook(self.path)
        row = wb.by_email("contact@cyberscope.fr")[0].row
        before = self.path.read_bytes()
        for title, text in (("Date de contact", "demain"), ("Date de relance", "01/01/2027"), ("Entreprise", "  ")):
            with self.subTest(title=title), self.assertRaises(WorkbookError):
                wb.edit_cell(row, self.column(wb, title), text)
        self.assertEqual(self.path.read_bytes(), before)

    def test_edit_cell_refuses_a_file_changed_elsewhere(self):
        wb = ApplicationWorkbook(self.path)
        row = wb.by_email("contact@cyberscope.fr")[0].row
        # Sorted in Excel: the row now holds another company.
        book = openpyxl.load_workbook(self.path)
        book["Suivi"].cell(row, 2, "Ceralis")
        book.save(self.path)
        os.utime(self.path, ns=(0, wb._loaded_mtime + 1_000_000))

        with self.assertRaises(WorkbookChanged):
            wb.edit_cell(row, self.column(wb, "Notes"), "Mauvaise ligne")
        self.assertIsNone(openpyxl.load_workbook(self.path)["Suivi"].cell(row, 11).value)
        self.assertEqual(wb.by_row(row).company, "Ceralis")  # reloaded for the user

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
