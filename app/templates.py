"""E-mail templates and placeholder rendering.

Placeholders are written {like_this} in French because the user edits the
templates in the settings window; PLACEHOLDER_HELP lists them.
"""

from __future__ import annotations

import re

DEFAULT_SUBJECT = "Candidature spontanée – Stage BTS SIO SLAM {periode}"

DEFAULT_BODY = """{salutation}

Étudiant en deuxième année de BTS SIO, option SLAM (développement d'applications), je recherche un stage {periode}.

Je souhaiterais réaliser ce stage au sein de {entreprise}. Vous trouverez ci-joint mon CV et ma lettre de motivation.

Je reste à votre disposition pour toute information complémentaire ou pour un échange, par mail ou par téléphone.

Bien cordialement,

{nom}
{telephone}"""

DEFAULT_FOLLOWUP_SUBJECT = "Re: {objet_initial}"

DEFAULT_FOLLOWUP_BODY = """{salutation}

Je me permets de revenir vers vous au sujet de ma candidature pour un stage {periode}, que je vous ai adressée le {date_envoi}.

Ce stage s'inscrit dans ma deuxième année de BTS SIO, option SLAM, et je serais vraiment intéressé de le réaliser au sein de {entreprise}. Je vous joins à nouveau mon CV et ma lettre de motivation.

Je reste disponible pour un échange à votre convenance.

Bien cordialement,

{nom}
{telephone}"""

PLACEHOLDER_HELP = (
    "{entreprise}  {service}  {interlocuteur}  {salutation}  {periode}  {nom}  {telephone}\n"
    "Relances uniquement : {date_envoi}  {objet_initial}"
)

_PLACEHOLDER = re.compile(r"\{(\w+)\}")
_ROLE_WORDS = re.compile(
    r"\b(service|rh|ressources|recrutement|dsi|direct\w*|responsable|resp|equipe|équipe|"
    r"président|president|pôle|pole|accueil|contact|secrétariat|secretariat|drh)\b",
    re.IGNORECASE,
)


def salutation(contact: str | None) -> str:
    """Opening line guessed from the 'Interlocuteur' column (editable before sending)."""
    contact = (contact or "").strip()
    lowered = contact.lower()
    for short, full in (("m. le ", "Monsieur le "), ("mme la ", "Madame la "), ("mme le ", "Madame le ")):
        if lowered.startswith(short):
            return full + contact[len(short):].strip() + ","
    if lowered.startswith(("monsieur le ", "madame la ", "madame le ")):
        return contact + ","
    name = re.sub(r"\s*\(.*?\)\s*", " ", contact).strip()
    words = name.split()
    if 2 <= len(words) <= 3 and not _ROLE_WORDS.search(name) and all(w[:1].isupper() for w in words):
        return f"Bonjour {name},"
    return "Madame, Monsieur,"


def render(template: str, context: dict[str, str]) -> str:
    """Replace known {placeholders}; unknown ones are left untouched."""
    return _PLACEHOLDER.sub(lambda m: str(context.get(m.group(1), m.group(0))), template)


def unresolved(text: str) -> list[str]:
    """Placeholders still present after rendering (typos in a template)."""
    return sorted({m.group(1) for m in _PLACEHOLDER.finditer(text)})


def tidy(text: str) -> str:
    """Remove blank trailing lines left by empty placeholders (e.g. no phone number)."""
    lines = [line.rstrip() for line in text.replace("\r\n", "\n").split("\n")]
    while lines and not lines[-1]:
        lines.pop()
    return "\n".join(lines)
