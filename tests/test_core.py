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
from common.gmail import LABEL_NOTIFIED, LABEL_REMINDED, LABEL_SENT, MessageInfo
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


if __name__ == "__main__":
    unittest.main()
