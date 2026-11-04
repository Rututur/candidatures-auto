"""Rough guess of what a company's answer means (refusal, positive, automatic...).

This is keyword matching on French (and a bit of English) phrasing, applied
to the new part of the message only (quoted history is stripped first, since
it contains our own application text). It is a hint, not a verdict.
"""

from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass

KIND_REFUSAL = "refusal"
KIND_POSITIVE = "positive"
KIND_UNKNOWN = "unknown"
KIND_AUTO = "auto"  # acknowledgement, out-of-office, no-reply system
KIND_BOUNCE = "bounce"  # undelivered message

# Labels displayed to the user (French UI).
KIND_LABELS = {
    KIND_REFUSAL: "Refus probable",
    KIND_POSITIVE: "Plutôt positif",
    KIND_UNKNOWN: "À lire (pas d'estimation)",
    KIND_AUTO: "Réponse automatique / accusé de réception",
    KIND_BOUNCE: "Mail non distribué",
}

REFUSAL_PATTERNS = [
    r"malheureusement",
    r"\bregret",
    r"ne (?:pouvons|pourrons|pourrions|peut|pourra) (?:pas|malheureusement pas)",
    r"(?:ne sommes|ne serons|ne sera) pas en mesure",
    r"pas en mesure de",
    r"pas en capacite",
    r"pas la possibilite",
    r"pas (?:de |pu |pouvoir )?donner (?:une )?suite",
    r"ne (?:pas |pourrons pas |pouvons pas )?donner suite",
    r"pas de suite favorable",
    r"(?:pas|ne) (?:ete )?retenue?",
    r"pas (?:de|d'offre de) stages?",
    r"pas de (?:place|poste|stagiaire|developpeur|developpement)",
    r"ne (?:prenons|recrutons|recherchons|accueillons) pas",
    r"n'accueill\w+ pas",
    r"deja pourvue?s?",
    r"\bcomplets?\b",
    r"declin\w+",
    r"ne correspond pas",
    r"(?:n'est|ne sera|serait) pas possible",
    r"\bimpossible\b",
    r"reponse negative",
    r"unfortunately",
    r"not (?:able|in a position) to",
]

POSITIVE_PATTERNS = [
    r"\bentretien",
    r"\brencontr\w+",
    r"disponibilites",
    r"rendez-vous|\brdv\b",
    r"convention de stage|\bconvention\b",
    r"avec plaisir",
    r"\bfavorable",
    r"interess\w+",
    r"\bechanger\b|un echange",
    r"(?:vous |nous )?(?:appeler|telephoner|rappeler)",
    r"\bvisio\b|\bteams\b|google meet",
    r"(?:passer|venir) nous voir",
    r"(?:heureux|ravis?|plaisir) de vous accueillir",
    r"\bbienvenue\b",
    r"\binterview\b",
]

ACKNOWLEDGEMENT_PATTERNS = [
    r"accuse de reception",
    r"(?:avons |a )?bien (?:ete )?recue?",
    r"reviendr\w+ vers vous",
    r"en cours d'(?:etude|analyse)",
    r"etudi\w+ (?:votre|attentivement)",
    r"message (?:automatique|genere automatiquement)",
    r"ne pas repondre",
    r"\babsente?\b|\babsence\b|out of office|conges",
]

BOUNCE_SENDERS = ("mailer-daemon", "postmaster", "mail-daemon")
BOUNCE_SUBJECTS = (
    "delivery status notification",
    "undeliverable",
    "undelivered",
    "non remis",
    "echec de la remise",
    "returned mail",
    "mail delivery",
)
NO_REPLY_SENDERS = ("noreply", "no-reply", "ne-pas-repondre", "nepasrepondre", "donotreply")

QUOTE_MARKERS = [
    r"\n(?:Le|On)\s[^\n]{0,200}(?:\n[^\n]{0,200})?(?:a\s+écrit|a\s+ecrit|wrote)\s*:",
    r"\n-{2,}\s*(?:Original Message|Message d'origine|Message original)",
    r"\n_{8,}",
    r"\n(?:De|From)\s*:\s*[^\n]+\n(?:Envoyé|Sent|Date)\s*:",
]


@dataclass
class Verdict:
    kind: str
    refusal_hits: int = 0
    positive_hits: int = 0

    @property
    def label(self) -> str:
        return KIND_LABELS[self.kind]


def normalize(text: str) -> str:
    text = text.replace("’", "'").replace(" ", " ")
    text = unicodedata.normalize("NFKD", text)
    text = "".join(c for c in text if not unicodedata.combining(c))
    return re.sub(r"[ \t]+", " ", text.lower())


def strip_quoted(text: str) -> str:
    """Keep only what the sender wrote, without the quoted conversation."""
    text = "\n" + text.replace("\r\n", "\n")
    cut = len(text)
    for marker in QUOTE_MARKERS:
        match = re.search(marker, text, flags=re.IGNORECASE)
        if match:
            cut = min(cut, match.start())
    lines = [line for line in text[:cut].split("\n") if not line.lstrip().startswith(">")]
    return "\n".join(lines).strip()


def _count(patterns: list[str], text: str) -> int:
    return sum(1 for pattern in patterns if re.search(pattern, text))


def classify(body: str, from_addr: str = "", subject: str = "", auto_submitted: bool = False) -> Verdict:
    sender = from_addr.lower()
    norm_subject = normalize(subject)
    if sender.split("@")[0] in BOUNCE_SENDERS or any(s in norm_subject for s in BOUNCE_SUBJECTS):
        return Verdict(KIND_BOUNCE)

    text = normalize(strip_quoted(body) + "\n" + subject)
    refusal = _count(REFUSAL_PATTERNS, text)
    positive = _count(POSITIVE_PATTERNS, text)

    if refusal and refusal >= positive:
        return Verdict(KIND_REFUSAL, refusal, positive)
    if positive and not auto_submitted:
        return Verdict(KIND_POSITIVE, refusal, positive)
    is_no_reply = any(token in sender for token in NO_REPLY_SENDERS)
    if auto_submitted or is_no_reply or _count(ACKNOWLEDGEMENT_PATTERNS, text):
        return Verdict(KIND_AUTO, refusal, positive)
    return Verdict(KIND_UNKNOWN, refusal, positive)
