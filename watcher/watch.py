"""Check Gmail for answers to applications and send Discord DMs.

Run every few minutes by GitHub Actions (.github/workflows/watch.yml).
State lives in Gmail labels, so each run is stateless:
  - a reply already notified carries LABEL_NOTIFIED,
  - an application already reminded carries LABEL_REMINDED.

IMPORTANT: the repository is public, so workflow logs are public too.
This script only logs counts, never addresses, names or subjects.

Configuration (environment variables):
  GMAIL_ADDRESS, GMAIL_APP_PASSWORD   Gmail account and app password
  DISCORD_BOT_TOKEN, DISCORD_USER_ID  bot token and your Discord user id
  REMINDER_DAYS        days without answer before a reminder (default 7)
  ALERT_MAX_AGE_HOURS  older replies are marked silently (default 72)
  DRY_RUN=1            print messages locally instead of sending/labelling
"""

from __future__ import annotations

import os
import sys
from datetime import datetime, timedelta, timezone

from common.classify import KIND_AUTO, KIND_BOUNCE, KIND_POSITIVE, KIND_REFUSAL, classify
from common.discord_dm import COLOR_BLUE, COLOR_GREEN, COLOR_GREY, COLOR_ORANGE, COLOR_RED, DiscordDM, DiscordError
from common.gmail import (
    LABEL_AUTO_REPLY,
    LABEL_NOTIFIED,
    LABEL_REMINDED,
    LABEL_SENT,
    FolderNotFound,
    GmailClient,
    GmailError,
    gmail_thread_url,
)
from common.tracking import Reply, SentApplication, find_replies, load_applications

try:
    from zoneinfo import ZoneInfo

    PARIS = ZoneInfo("Europe/Paris")
except Exception:  # tz database missing: fall back to UTC
    PARIS = timezone.utc

KIND_STYLE = {
    KIND_REFUSAL: ("❌", "Refus probable", COLOR_RED),
    KIND_POSITIVE: ("✅", "Plutôt positif", COLOR_GREEN),
    KIND_AUTO: ("🤖", "Réponse automatique / accusé de réception", COLOR_GREY),
    KIND_BOUNCE: ("⚠️", "Mail non distribué : vérifie l'adresse", COLOR_ORANGE),
}
DEFAULT_STYLE = ("❔", "À lire, pas d'estimation possible", COLOR_BLUE)


def env(name: str, default: str | None = None) -> str:
    value = os.environ.get(name, default)
    if value is None or value == "":
        sys.exit(
            f"Missing environment variable {name}. "
            "On GitHub: add it under Settings > Secrets and variables > Actions. "
            "On your PC: see the README, the syntax depends on the terminal."
        )
    return value


def local(dt: datetime | None) -> str:
    return dt.astimezone(PARIS).strftime("%d/%m/%Y à %H:%M") if dt else "?"


def reply_embed(reply: Reply, kind: str, account: str) -> dict:
    icon, verdict, color = KIND_STYLE.get(kind, DEFAULT_STYLE)
    msg = reply.message
    sender = f"{msg.from_name} <{msg.from_addr}>" if msg.from_name else msg.from_addr
    title = (f"⚠️ Mail non distribué : {reply.application.name}" if kind == KIND_BOUNCE
             else f"📬 Réponse de {reply.application.name}")
    return {
        "title": title[:256],
        "url": gmail_thread_url(account, msg.thread_id) if msg.thread_id else None,
        "color": color,
        "fields": [
            {"name": "Objet", "value": (msg.subject or "(sans objet)")[:1024], "inline": False},
            {"name": "Estimation", "value": f"{icon} {verdict}", "inline": True},
            {"name": "Reçu le", "value": local(msg.date), "inline": True},
            {"name": "De", "value": sender[:1024], "inline": False},
        ],
        "footer": {"text": "Estimation automatique : lis le mail pour en être sûr. Le titre ouvre Gmail."},
    }


def reminder_embed(app: SentApplication, days: int, account: str) -> dict:
    last = app.last
    count = len(app.messages)
    history = "Candidature envoyée" if count == 1 else f"Relance n°{count - 1} envoyée"
    return {
        "title": f"⏰ À relancer : {app.name}"[:256],
        "url": gmail_thread_url(account, last.thread_id) if last.thread_id else None,
        "color": COLOR_ORANGE,
        "description": (
            f"{history} le {local(last.date)}, sans réponse depuis {days} jours.\n"
            "La relance est prête dans l'onglet « À relancer » de l'application : "
            "relis-la, modifie-la si besoin, puis envoie-la."
        ),
        "fields": [{"name": "Adresse", "value": app.address, "inline": False}],
    }


def run() -> int:
    account = env("GMAIL_ADDRESS").strip().lower()
    password = env("GMAIL_APP_PASSWORD")
    dry_run = os.environ.get("DRY_RUN", "") not in ("", "0", "false")
    reminder_days = int(os.environ.get("REMINDER_DAYS") or 7)
    max_age = timedelta(hours=int(os.environ.get("ALERT_MAX_AGE_HOURS") or 72))
    discord = None if dry_run else DiscordDM(env("DISCORD_BOT_TOKEN"), env("DISCORD_USER_ID"))
    now = datetime.now(timezone.utc)

    def notify(embed: dict) -> None:
        embed = {k: v for k, v in embed.items() if v is not None}
        if dry_run:
            print("[dry-run] DM:", embed.get("title"), "|", embed.get("description") or embed.get("fields"))
        else:
            discord.send(embed=embed)

    def label(client: GmailClient, uid: int, name: str) -> None:
        if not dry_run:
            client.add_label([uid], name)

    with GmailClient(account, password) as client:
        if not dry_run:
            client.ensure_labels()
        try:
            applications = load_applications(client)
        except FolderNotFound:  # dry run on a mailbox where the labels were never created
            applications = {}
        print(f"Tracked recipients: {len(applications)}")
        if not applications:
            print(f"Nothing to watch yet: no sent e-mail carries the label {LABEL_SENT}.")
            return 0

        replies = find_replies(client, applications)  # leaves All Mail selected
        if not dry_run:
            client.select(client.special_folder("\\All"), readonly=False)

        notified = silent = 0
        answered: set[str] = set()
        for reply in replies:
            msg = reply.message
            if LABEL_NOTIFIED in msg.labels:
                if LABEL_AUTO_REPLY not in msg.labels and not msg.auto_submitted:
                    answered.add(reply.application.address)
                continue
            text = client.fetch_text(msg.uid)
            kind = classify(text, msg.from_addr, msg.subject, msg.auto_submitted).kind
            if kind == KIND_AUTO:
                label(client, msg.uid, LABEL_AUTO_REPLY)
            else:
                answered.add(reply.application.address)
            if msg.date and now - msg.date > max_age:
                silent += 1  # first run: do not flood with old answers
            else:
                notify(reply_embed(reply, kind, account))
                notified += 1
            label(client, msg.uid, LABEL_NOTIFIED)

        # Reminders: once per sent message, after REMINDER_DAYS without a real answer.
        if not dry_run:
            client.select(LABEL_SENT, readonly=False)
        reminded = 0
        for app in applications.values():
            if app.address in answered:
                continue
            last = app.last
            if not last.date or LABEL_REMINDED in last.labels:
                continue
            if now - last.date < timedelta(days=reminder_days):
                continue
            notify(reminder_embed(app, (now - last.date).days, account))
            label(client, last.uid, LABEL_REMINDED)
            reminded += 1

        print(f"Replies found: {len(replies)}, notified: {notified}, marked silently: {silent}, reminders: {reminded}")
    return 0


if __name__ == "__main__":
    try:
        sys.exit(run())
    except (GmailError, DiscordError) as exc:
        sys.exit(f"Error: {exc}")
