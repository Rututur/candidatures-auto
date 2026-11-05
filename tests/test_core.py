"""Tests for the shared logic and the watcher (no network needed).

Run from the project folder:  python -m unittest
"""

from __future__ import annotations

import io
import os
import unittest
from contextlib import redirect_stdout
from datetime import datetime, timedelta, timezone
from unittest import mock

from common import gmail
from common.classify import KIND_AUTO, KIND_BOUNCE, KIND_POSITIVE, KIND_REFUSAL, KIND_UNKNOWN, classify, strip_quoted
from common.gmail import LABEL_AUTO_REPLY, LABEL_NOTIFIED, LABEL_REMINDED, LABEL_SENT, LABEL_VALIDATED, MessageInfo
from common.tracking import find_replies, load_applications
from tests.fake_gmail import FakeGmail

ME = "arthur@gmail.com"
NOW = datetime.now(timezone.utc)


def msg(uid, sender, to, subject, days_ago, thread, labels=(), auto=False):
    return MessageInfo(
        uid=uid, thread_id=thread, gm_msgid=uid, labels=set(labels), from_name="", from_addr=sender,
        to=[(name, addr) for name, addr in to], subject=subject, date=NOW - timedelta(days=days_ago),
        message_id=f"<{uid}@test>", in_reply_to="", auto_submitted=auto,
    )


QUOTED = "\n\nLe mer. 4 nov. 2026 à 10:00, Arthur <arthur@gmail.com> a\nécrit :\n> Je reste disponible pour un entretien."


class ClassifyTest(unittest.TestCase):
    def test_refusal(self):
        text = "Bonjour,\nMalheureusement nous ne pouvons pas donner suite à votre demande de stage." + QUOTED
        self.assertEqual(classify(text).kind, KIND_REFUSAL)

    def test_refusal_wins_over_favorable(self):
        self.assertEqual(classify("Nous ne pouvons pas réserver une suite favorable à votre candidature.").kind, KIND_REFUSAL)

    def test_positive(self):
        text = "Bonjour Arthur, votre profil nous intéresse. Pouvez-vous m'indiquer vos disponibilités pour un entretien ?"
        self.assertEqual(classify(text).kind, KIND_POSITIVE)

    def test_quoted_text_is_ignored(self):
        # "entretien" only appears in our own quoted mail: must not count as positive.
        self.assertEqual(classify("Bonjour, merci pour votre message." + QUOTED).kind, KIND_UNKNOWN)
        self.assertNotIn("entretien", strip_quoted("Merci." + QUOTED))

    def test_auto_reply(self):
        self.assertEqual(classify("Je suis absente jusqu'au 12 novembre.", auto_submitted=True).kind, KIND_AUTO)
        self.assertEqual(classify("Nous accusons réception de votre candidature.", "noreply@rh.fr").kind, KIND_AUTO)

    def test_bounce(self):
        self.assertEqual(classify("Address not found", "mailer-daemon@googlemail.com").kind, KIND_BOUNCE)


class ImapParsingTest(unittest.TestCase):
    def test_fetch_headers_parses_gmail_extensions(self):
        raw = (b"From: =?utf-8?q?Service_RH?= <rh@sodebo.fr>\r\nTo: Arthur <arthur@gmail.com>\r\n"
               b"Subject: =?utf-8?q?R=C3=A9ponse_candidature?=\r\nDate: Wed, 04 Nov 2026 10:00:00 +0100\r\n"
               b"Message-ID: <abc@sodebo.fr>\r\nAuto-Submitted: auto-replied\r\n\r\n")
        data = [(b'7 (X-GM-THRID 1790000000000001 X-GM-MSGID 1790000000000002 X-GM-LABELS ("\\\\Important" '
                 b'Candidatures/Notifiees) UID 42 BODY[HEADER.FIELDS (FROM)] {120}', raw), b")"]
        client = gmail.GmailClient(ME, "x")
        client.imap = mock.Mock()
        client.imap.uid.return_value = ("OK", data)
        info = client.fetch_headers([42])[0]
        self.assertEqual(info.uid, 42)
        self.assertEqual(info.thread_id, 1790000000000001)
        self.assertIn(LABEL_NOTIFIED, info.labels)
        self.assertEqual(info.from_addr, "rh@sodebo.fr")
        self.assertEqual(info.subject, "Réponse candidature")
        self.assertTrue(info.auto_submitted)

    def test_sender_with_special_characters_in_encoded_name(self):
        meta = "1 (X-GM-THRID 5 X-GM-MSGID 6 X-GM-LABELS () UID 42 BODY[HEADER.FIELDS (FROM)] {1}"
        senders = {
            "=?utf-8?q?Lef=C3=A8vre=2C_Marie?= <rh@sodebo.fr>": "Lefèvre, Marie",
            "=?utf-8?q?rh=40sodebo=2Efr_=C3=A9quipe?= <rh@sodebo.fr>": "rh@sodebo.fr équipe",
            "=?utf-8?q?Sodebo=3B_RH_=C3=A9quipe?= <rh@sodebo.fr>": "Sodebo; RH équipe",
            "=?utf-8?q?Mich=C3=A8le?= =?utf-8?q?_Dupont=2C_RH?= <rh@sodebo.fr>": "Michèle Dupont, RH",
            "Marie Dupont <rh@sodebo.fr>": "Marie Dupont",
            "rh@sodebo.fr": "",
        }
        for header, name in senders.items():
            with self.subTest(header=header):
                info = gmail.parse_header_block(42, meta, f"From: {header}\r\n\r\n".encode())
                self.assertEqual(info.from_addr, "rh@sodebo.fr")
                self.assertEqual(info.from_name, name)

    def test_special_folder_from_list(self):
        client = gmail.GmailClient(ME, "x")
        client.imap = mock.Mock()
        client.imap.list.return_value = ("OK", [
            b'(\\HasNoChildren) "/" "INBOX"',
            b'(\\All \\HasNoChildren) "/" "[Gmail]/Tous les messages"',
            b'(\\HasNoChildren \\Sent) "/" "[Gmail]/Messages envoy&AOk-s"',
        ])
        self.assertEqual(client.special_folder("\\Sent"), "[Gmail]/Messages envoy&AOk-s")
        self.assertEqual(client.special_folder("\\All"), "[Gmail]/Tous les messages")


def mailbox():
    return FakeGmail(ME, [
        # Sent applications
        (msg(1, ME, [("Cyberscope", "contact@cyberscope.fr")], "Candidature", 10, 100, [LABEL_SENT]), ""),
        (msg(2, ME, [("Ceralis", "contact@ceralis.fr")], "Candidature", 9, 200, [LABEL_SENT]), ""),
        (msg(3, ME, [("UP2TECH", "contact@up2tech.fr")], "Candidature", 1, 300, [LABEL_SENT]), ""),
        (msg(4, ME, [("Evolusite", "contact@evolusite.fr")], "Candidature", 8, 400, [LABEL_SENT]), ""),
        # Answer from another address of the same domain, new thread
        (msg(10, "rh@cyberscope.fr", [("", ME)], "Votre candidature", 1, 900),
         "Bonjour, nous serions ravis de vous rencontrer en entretien."),
        # Answer in the same thread (bounce)
        (msg(11, "mailer-daemon@googlemail.com", [("", ME)], "Delivery Status Notification", 0, 200), "not found"),
        # Out-of-office reply for Evolusite: still needs a reminder
        (msg(12, "contact@evolusite.fr", [("", ME)], "Absence", 7, 400, auto=True), "Je suis absent."),
        # Unrelated mail
        (msg(13, "news@shop.fr", [("", ME)], "Promo", 0, 500), "Promo"),
        # Old mail from Ceralis before the application: not an answer
        (msg(14, "contact@ceralis.fr", [("", ME)], "Ancien échange", 30, 600), "Bonjour"),
    ])


class BatchDoneTest(unittest.TestCase):
    def finish_batch(self, sent, test):
        from app import gui

        app = mock.Mock()
        with mock.patch.object(gui, "messagebox"):
            gui.App._batch_done(app, gui.MODE_SEND, sent, 0, test)
        return app

    def test_real_batch_syncs_to_catch_up_missing_labels(self):
        self.finish_batch(sent=3, test=False).run_sync.assert_called_once_with(silent=True)

    def test_no_sync_after_test_batch_or_when_nothing_was_sent(self):
        self.finish_batch(sent=3, test=True).run_sync.assert_not_called()
        self.finish_batch(sent=0, test=False).run_sync.assert_not_called()


class ReviewFlowTest(unittest.TestCase):
    def test_failed_decision_keeps_the_reply_pending(self):
        from app import gui

        reply = mock.Mock(uid=1, company="Sodebo")
        app = mock.Mock(pending=[reply])
        with mock.patch.object(gui, "messagebox") as box:
            gui.App._decision_done(app, reply, gui.STATUS_REFUSED, WorkbookLocked("Excel"))
        box.showerror.assert_called_once()
        self.assertEqual(app.pending, [reply])

    def test_recorded_decision_removes_the_reply(self):
        from app import gui

        reply = mock.Mock(uid=1, company="Sodebo")
        app = mock.Mock(pending=[reply])
        gui.App._decision_done(app, reply, gui.STATUS_REFUSED, None)
        self.assertEqual(app.pending, [])
        app.refresh_tabs.assert_called_once()


class WorkbookLocked(Exception):
    pass


class TableTabTest(unittest.TestCase):
    def test_click_on_the_edge_of_a_cell_selects_that_cell_column(self):
        from app import gui

        tab = mock.Mock()
        tab.app.workbook.table_columns = ["first", "second"]
        tab.tree.bbox.side_effect = lambda row, column: {"#1": (0, 25, 100, 20), "#2": (100, 25, 100, 20)}[column]
        self.assertEqual(gui.TableTab._column_at(tab, "5", 99), "first")
        self.assertEqual(gui.TableTab._column_at(tab, "5", 100), "second")
        self.assertEqual(gui.TableTab._column_at(tab, "5", 103), "second")
        self.assertIsNone(gui.TableTab._column_at(tab, "5", 200))  # right of the last column


class SelectTest(unittest.TestCase):
    def client_with_select_answer(self, data):
        client = gmail.GmailClient(ME, "x")
        client.imap = mock.Mock()
        client.imap.select.return_value = ("NO", data)
        return client

    def test_missing_label_raises_folder_not_found(self):
        client = self.client_with_select_answer([b"[NONEXISTENT] Unknown Mailbox: Candidatures/Envoyees (Failure)"])
        with self.assertRaises(gmail.FolderNotFound):
            client.select(LABEL_SENT)

    def test_other_refusal_stays_a_generic_error(self):
        client = self.client_with_select_answer([b"[UNAVAILABLE] Temporary failure"])
        with self.assertRaises(gmail.GmailError) as caught:
            client.select(LABEL_SENT)
        self.assertNotIsInstance(caught.exception, gmail.FolderNotFound)


class DiscordSetupTest(unittest.TestCase):
    def test_invalid_token_gives_a_readable_error(self):
        from tools import setup_discord

        ws = mock.Mock()
        # Discord greets the client, then closes the connection on a bad token.
        ws.recv.side_effect = ['{"op": 10, "d": {"heartbeat_interval": 41250}}', ""]
        with mock.patch.object(setup_discord.websocket, "create_connection", return_value=ws):
            with self.assertRaisesRegex(RuntimeError, "token"):
                setup_discord.identify_once("bad-token")
        ws.close.assert_called_once()


class TrackingTest(unittest.TestCase):
    def test_find_replies(self):
        box = mailbox()
        apps = load_applications(box)
        self.assertEqual(set(apps), {"contact@cyberscope.fr", "contact@ceralis.fr", "contact@up2tech.fr", "contact@evolusite.fr"})
        self.assertEqual(apps["contact@cyberscope.fr"].name, "Cyberscope")
        replies = {r.message.uid: r for r in find_replies(box, apps)}
        self.assertEqual(set(replies), {10, 11, 12})
        self.assertEqual(replies[10].matched_by, "domain")
        self.assertEqual(replies[11].matched_by, "thread")
        self.assertEqual(replies[12].application.address, "contact@evolusite.fr")


class WatcherTest(unittest.TestCase):
    def run_watcher(self, box):
        from watcher import watch

        sent = []
        fake_discord = mock.Mock()
        fake_discord.return_value.send.side_effect = lambda embed=None, content="": sent.append(embed)
        env = {"GMAIL_ADDRESS": ME, "GMAIL_APP_PASSWORD": "x", "DISCORD_BOT_TOKEN": "t", "DISCORD_USER_ID": "1"}
        with mock.patch.dict(os.environ, env), mock.patch.object(watch, "GmailClient", lambda *a: box), \
                mock.patch.object(watch, "DiscordDM", fake_discord), redirect_stdout(io.StringIO()) as out:
            watch.run()
        return sent, out.getvalue()

    def test_alerts_and_reminders(self):
        box = mailbox()
        sent, log = self.run_watcher(box)
        titles = [e["title"] for e in sent]
        self.assertIn("📬 Réponse de Cyberscope", titles)
        self.assertIn("⚠️ Mail non distribué : Ceralis", titles)
        # Evolusite only sent an out-of-office (older than 72 h: silent) -> reminder.
        self.assertIn("⏰ À relancer : Evolusite", titles)
        self.assertNotIn("⏰ À relancer : Cyberscope", titles)  # answered
        self.assertNotIn("⏰ À relancer : UP2TECH", titles)  # sent yesterday
        estimations = {e["title"]: e["fields"][1]["value"] for e in sent if "fields" in e and len(e["fields"]) > 1}
        self.assertIn("Plutôt positif", estimations["📬 Réponse de Cyberscope"])
        self.assertTrue(box.messages[4][0].labels >= {LABEL_REMINDED})
        # Public logs must not leak names or addresses.
        self.assertNotIn("@", log)
        self.assertNotIn("Cyberscope", log)

        # Second run: nothing new to say.
        sent_again, _ = self.run_watcher(box)
        self.assertEqual(sent_again, [])

    def test_user_decision_replaces_the_estimate(self):
        box = mailbox()
        # Marked as a real answer by the user, although the mail looks automatic.
        box.messages[12][0].labels.add(LABEL_VALIDATED)
        # Ignored by the user, although the estimate says this is a positive answer.
        box.messages[10][0].labels.update({LABEL_VALIDATED, LABEL_AUTO_REPLY})
        sent, _ = self.run_watcher(box)
        titles = [e["title"] for e in sent]
        self.assertNotIn("📬 Réponse de Cyberscope", titles)  # already seen by the user
        self.assertNotIn("⏰ À relancer : Evolusite", titles)  # really answered
        self.assertIn("⏰ À relancer : Cyberscope", titles)  # still waiting for an answer
        self.assertNotIn(LABEL_NOTIFIED, box.messages[12][0].labels)

    def test_dry_run_on_mailbox_without_labels(self):
        from watcher import watch

        box = FakeGmail(ME, [])
        env = {"GMAIL_ADDRESS": ME, "GMAIL_APP_PASSWORD": "x", "DRY_RUN": "1"}
        with mock.patch.dict(os.environ, env), mock.patch.object(watch, "GmailClient", lambda *a: box), \
                mock.patch.object(box, "select", side_effect=gmail.FolderNotFound("missing")), \
                redirect_stdout(io.StringIO()) as out:
            self.assertEqual(watch.run(), 0)
        self.assertIn("Tracked recipients: 0", out.getvalue())


if __name__ == "__main__":
    unittest.main()
