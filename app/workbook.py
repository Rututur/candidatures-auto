"""Read and update the application-tracking spreadsheet.

The tool adapts to an existing sheet: it finds the header row, maps columns
by name (accents and case ignored) and adds the few columns it needs.
Every write reloads the file, applies one change and saves atomically, so
edits made in Excel between two writes are not lost (as long as Excel has
closed the file).
"""

from __future__ import annotations

import os
import re
import shutil
import tempfile
import unicodedata
from copy import copy
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path

import openpyxl
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter

# Status values: must match the drop-down list of the user's sheet.
STATUS_TO_CONTACT = "À contacter"
STATUS_LETTER_READY = "Lettre prête"
STATUS_SENT = "Envoyée"
STATUS_FOLLOWED_UP = "Relancée"
STATUS_REPLIED = "Réponse reçue"  # added by the tool
STATUS_INTERVIEW = "Entretien prévu"
STATUS_ACCEPTED = "Acceptée"
STATUS_REFUSED = "Refusée"
STATUS_NO_ANSWER = "Sans réponse"

SENDABLE_STATUSES = {"", STATUS_TO_CONTACT, STATUS_LETTER_READY}
WAITING_STATUSES = {STATUS_SENT, STATUS_FOLLOWED_UP}
REPLY_FILL = "FFD0F0EC"

# What the user can decide about a reply: one of these, or a status to apply.
DECISION_BOUNCE = "bounce"
DECISION_IGNORE = "ignore"

FIELD_ALIASES = {
    "company": ["entreprise", "societe", "structure", "organisme", "employeur"],
    "email": ["email", "e-mail", "mail", "courriel", "adresse mail", "adresse email", "adresse e-mail"],
    "contact": ["interlocuteur", "contact", "destinataire", "nom du contact"],
    "service": ["service vise", "service", "poste", "poste vise"],
    "contact_mode": ["mode de contact", "canal", "moyen de contact"],
    "contact_date": ["date de contact", "date d'envoi", "date envoi", "envoyee le"],
    "followup_date": ["date de relance"],
    "status": ["statut", "etat"],
    "reply_date": ["date de reponse", "reponse le"],
    "notes": ["notes", "commentaires", "remarques", "note"],
    "letter_file": ["fichier lettre (pdf)", "fichier lettre", "lettre (pdf)"],
    "followup_sent": ["relance envoyee le", "date de la relance envoyee"],
    "action": ["action"],
}
REQUIRED_FIELDS = ("company", "email")
DATE_FIELDS = {"contact_date", "reply_date", "followup_sent"}
# Columns created when missing (header text, is a date column).
CREATABLE_FIELDS = {
    "contact_mode": ("Mode de contact", False),
    "contact_date": ("Date de contact", True),
    "status": ("Statut", False),
    "reply_date": ("Date de réponse", True),
    "notes": ("Notes", False),
    "letter_file": ("Fichier lettre (PDF)", False),
    "followup_sent": ("Relance envoyée le", True),
}
DATE_FORMAT = "dd/mm/yyyy"
EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[a-z]{2,}$", re.IGNORECASE)


def _covers(sqref, letter: str) -> bool:
    """True when a range list such as 'N5:N45 P5:P45' starts in column `letter`."""
    return any(part.split(":")[0].rstrip("0123456789") == letter for part in str(sqref).split())


def _list_choices(ws, letter: str) -> list[str]:
    """Values of the drop-down list applied to a column, if any."""
    for dv in ws.data_validations.dataValidation:
        if dv.type == "list" and dv.formula1 and dv.formula1.startswith('"') and _covers(dv.sqref, letter):
            return dv.formula1.strip('"').split(",")
    return []


class WorkbookError(Exception):
    """Errors shown as-is to the user (French)."""


class WorkbookChanged(WorkbookError):
    """The file changed on disk since it was loaded, so row numbers may be out of date."""


def normalize(text) -> str:
    text = unicodedata.normalize("NFKD", str(text or "").replace("’", "'"))
    text = "".join(c for c in text if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", text).strip().lower()


def _as_date(value) -> date | None:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, str):
        for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d/%m/%y"):
            try:
                return datetime.strptime(value.strip(), fmt).date()
            except ValueError:
                pass
    return None


def _as_text(value) -> str:
    return "" if value is None else str(value).strip()


@dataclass
class Row:
    row: int
    company: str
    email: str = ""
    contact: str = ""
    service: str = ""
    status: str = ""
    contact_mode: str = ""
    contact_date: date | None = None
    reply_date: date | None = None
    followup_sent: date | None = None
    notes: str = ""
    letter_file: str = ""

    @property
    def has_valid_email(self) -> bool:
        return bool(EMAIL_RE.match(self.email))

    @property
    def last_contact(self) -> date | None:
        return self.followup_sent or self.contact_date

    def followup_due(self, days: int) -> date | None:
        last = self.last_contact
        return last + timedelta(days=days) if last else None

    def needs_followup(self, days: int, today: date | None = None) -> bool:
        due = self.followup_due(days)
        return (
            self.status in WAITING_STATUSES
            and self.has_valid_email
            and self.reply_date is None
            and due is not None
            and (today or date.today()) >= due
        )

    @property
    def can_be_sent(self) -> bool:
        return self.has_valid_email and self.status in SENDABLE_STATUSES


@dataclass
class Column:
    """A column of the sheet, as shown in the application."""

    index: int
    title: str
    field: str | None
    read_only: bool  # holds formulas
    is_date: bool
    choices: list[str]
    width: float


def _find_header(ws) -> tuple[int, dict[str, int]] | None:
    for row_idx in range(1, min(ws.max_row, 20) + 1):
        columns: dict[str, int] = {}
        for cell in ws[row_idx]:
            label = normalize(cell.value)
            if not label:
                continue
            for field, aliases in FIELD_ALIASES.items():
                if field not in columns and label in aliases:
                    columns[field] = cell.column
        if all(f in columns for f in REQUIRED_FIELDS):
            return row_idx, columns
    return None


class ApplicationWorkbook:
    def __init__(self, path: str | os.PathLike, reminder_days: int = 7):
        self.path = Path(path)
        self.reminder_days = reminder_days
        self._backed_up = False
        self.rows: list[Row] = []
        self.load()

    # -- reading --------------------------------------------------------
    def _locate(self, wb):
        for ws in wb.worksheets:
            found = _find_header(ws)
            if found:
                return ws, found[0], found[1]
        raise WorkbookError(
            "Je ne trouve pas la ligne d'en-têtes : il faut au moins une colonne « Entreprise » "
            "et une colonne « Email » dans le tableur."
        )

    def load(self) -> None:
        if not self.path.exists():
            raise WorkbookError(f"Fichier introuvable : {self.path}")
        self._loaded_mtime = self.path.stat().st_mtime_ns
        try:
            wb = openpyxl.load_workbook(self.path)
        except Exception as exc:
            raise WorkbookError(f"Impossible d'ouvrir le tableur : {exc}") from exc
        ws, self.header_row, self.columns = self._locate(wb)
        self.sheet_name = ws.title
        self.rows = []
        for r in range(self.header_row + 1, ws.max_row + 1):
            get = lambda field: ws.cell(r, self.columns[field]).value if field in self.columns else None  # noqa: E731
            company = _as_text(get("company"))
            if not company:
                continue
            self.rows.append(Row(
                row=r,
                company=company,
                email=_as_text(get("email")).lower(),
                contact=_as_text(get("contact")),
                service=_as_text(get("service")),
                status=_as_text(get("status")),
                contact_mode=_as_text(get("contact_mode")),
                contact_date=_as_date(get("contact_date")),
                reply_date=_as_date(get("reply_date")),
                followup_sent=_as_date(get("followup_sent")),
                notes=_as_text(get("notes")),
                letter_file=_as_text(get("letter_file")),
            ))
        self.table_columns = self._read_columns(ws)
        self.cells = {
            item.row: {c.index: ws.cell(item.row, c.index).value for c in self.table_columns}
            for item in self.rows
        }

    def _read_columns(self, ws) -> list[Column]:
        fields = {col: field for field, col in self.columns.items()}
        result = []
        for cell in ws[self.header_row]:
            title = " ".join(_as_text(cell.value).split())
            if not title:
                continue
            values = (ws.cell(r, cell.column).value for r in range(self.header_row + 1, ws.max_row + 1))
            letter = get_column_letter(cell.column)
            field = fields.get(cell.column)
            result.append(Column(
                index=cell.column,
                title=title,
                field=field,
                read_only=any(isinstance(v, str) and v.startswith("=") for v in values),
                is_date=field in DATE_FIELDS,
                choices=_list_choices(ws, letter),
                width=ws.column_dimensions[letter].width or 15,
            ))
        return result

    def display_value(self, row: int, column: Column) -> str:
        """Text shown for a cell. The file keeps no formula results, so known ones are computed."""
        item = self.by_row(row)
        if column.read_only:
            if column.field == "followup_date":
                due = item.followup_due(self.reminder_days)
                return due.strftime("%d/%m/%Y") if due else ""
            if column.field == "action":
                return "À relancer" if item.needs_followup(self.reminder_days) else ""
            return "(formule)"
        value = self.cells[row][column.index]
        if isinstance(value, (date, datetime)):
            return value.strftime("%d/%m/%Y")
        return _as_text(value)

    def by_row(self, row: int) -> Row:
        for item in self.rows:
            if item.row == row:
                return item
        raise KeyError(row)

    def by_email(self, address: str) -> list[Row]:
        address = address.lower()
        return [r for r in self.rows if r.email == address]

    # -- writing --------------------------------------------------------
    def check_writable(self) -> None:
        try:
            with open(self.path, "r+b"):
                pass
        except PermissionError:
            raise WorkbookError(
                "Le tableur est ouvert dans Excel (ou verrouillé). Ferme-le puis réessaie : "
                "l'outil doit pouvoir l'enregistrer."
            ) from None

    def _backup(self) -> None:
        if self._backed_up:
            return
        target = self.path.with_name(f"{self.path.stem} (sauvegarde auto){self.path.suffix}")
        shutil.copy2(self.path, target)
        self._backed_up = True

    def _save(self, wb) -> None:
        self._backup()
        wb.calculation.fullCalcOnLoad = True  # formulas are recomputed when Excel opens it
        fd, tmp = tempfile.mkstemp(suffix=".xlsx", dir=self.path.parent)
        os.close(fd)
        try:
            wb.save(tmp)
            os.replace(tmp, self.path)
        except PermissionError:
            raise WorkbookError("Impossible d'enregistrer : le tableur est ouvert dans Excel. Ferme-le puis réessaie.") from None
        finally:
            if os.path.exists(tmp):
                os.remove(tmp)

    def update(self, row: int, changes: dict) -> None:
        """Apply {field: value} to one row and save. Notes are appended, not replaced."""
        wb = openpyxl.load_workbook(self.path)
        ws, header_row, columns = self._locate(wb)
        columns = self._upgrade(wb, ws, header_row, columns)
        for field, value in changes.items():
            cell = ws.cell(row, columns[field])
            if field == "notes":
                previous = _as_text(cell.value)
                value = f"{previous}\n{value}" if previous else value
            cell.value = value
        self._save(wb)
        self.load()

    def _parse_input(self, row: int, column: Column, text: str):
        text = text.strip()
        if column.read_only:
            raise WorkbookError(f"La colonne « {column.title} » est calculée par une formule : elle ne se modifie pas.")
        if not text:
            if column.field == "company":
                raise WorkbookError("Le nom de l'entreprise ne peut pas être vide.")
            return None
        if column.is_date:
            parsed = _as_date(text)
            if parsed is None:
                raise WorkbookError(f"« {text} » n'est pas une date valide : écris-la sous la forme jj/mm/aaaa.")
            return parsed
        previous = self.cells[row][column.index]
        if isinstance(previous, (int, float)) and not isinstance(previous, bool):
            try:
                return int(text)
            except ValueError:
                try:
                    return float(text.replace(",", "."))
                except ValueError:
                    return text
        return text

    def edit_cell(self, row: int, column: Column, text: str) -> None:
        """Replace one cell in the file with what the user typed."""
        if self.path.stat().st_mtime_ns != self._loaded_mtime:
            self.load()
            raise WorkbookChanged(
                "Le fichier a été modifié ailleurs, par exemple dans Excel. Le tableau vient d'être rechargé : "
                "vérifie la ligne, puis refais ta modification."
            )
        value = self._parse_input(row, column, text)
        self.check_writable()
        wb = openpyxl.load_workbook(self.path)
        ws, _, _ = self._locate(wb)
        cell = ws.cell(row, column.index)
        cell.value = value
        if isinstance(value, date):
            cell.number_format = DATE_FORMAT
        self._save(wb)
        self.load()

    def ensure_structure(self) -> None:
        """Add the tool's columns and statuses without changing any row."""
        wb = openpyxl.load_workbook(self.path)
        ws, header_row, columns = self._locate(wb)
        self._upgrade(wb, ws, header_row, columns)
        self._save(wb)
        self.load()

    # -- high-level updates -------------------------------------------
    def mark_sent(self, row: int, when: date, letter_path: str) -> None:
        self.update(row, {
            "status": STATUS_SENT,
            "contact_mode": "Email",
            "contact_date": when,
            "letter_file": letter_path,
        })

    def mark_followed_up(self, row: int, when: date) -> None:
        self.update(row, {"status": STATUS_FOLLOWED_UP, "followup_sent": when})

    def apply_decision(self, row: int, decision: str, when: date, subject: str) -> None:
        """Record what the user decided about the reply received on `when`."""
        item = self.by_row(row)
        stamp = f"[Validé {date.today():%d/%m/%Y}]"
        mail = f"mail du {when:%d/%m/%Y}, objet : {subject or 'sans objet'}"
        if decision == DECISION_IGNORE:
            # Older versions applied their estimate directly: undo it if it came from this mail.
            if item.reply_date == when and f"[Auto {when:%d/%m/%Y}] Réponse reçue" in item.notes:
                self.update(row, {
                    "status": STATUS_FOLLOWED_UP if item.followup_sent else STATUS_SENT,
                    "reply_date": None,
                    "notes": f"{stamp} Estimation annulée ({mail}).",
                })
            return
        if decision == DECISION_BOUNCE:
            self.update(row, {"notes": f"{stamp} Mail non distribué à {item.email} : vérifie l'adresse."})
            return
        changes = {"status": decision, "notes": f"{stamp} {decision} ({mail})."}
        if item.reply_date is None:
            changes["reply_date"] = when
        self.update(row, changes)

    # -- structure upgrades (idempotent) ------------------------------
    def _upgrade(self, wb, ws, header_row: int, columns: dict[str, int]) -> dict[str, int]:
        columns = dict(columns)
        last_row = max(ws.max_row, header_row + 1)
        for field, (title, is_date) in CREATABLE_FIELDS.items():
            if field in columns:
                continue
            col = max(c.column for c in ws[header_row] if c.value is not None) + 1
            self._add_column(ws, header_row, last_row, col, title, is_date, columns)
            columns[field] = col
        if "status" in columns:
            self._add_reply_status(ws, header_row, last_row, columns["status"])
        if "followup_date" in columns and "contact_date" in columns and "followup_sent" in columns:
            self._update_followup_formulas(ws, header_row, last_row, columns)
        for other in wb.worksheets:
            if other is not ws:
                self._update_side_sheets(other)
        return columns

    @staticmethod
    def _add_column(ws, header_row, last_row, col, title, is_date, columns) -> None:
        template_col = columns.get("notes") or (col - 1)
        header = ws.cell(header_row, col, title)
        source = ws.cell(header_row, col - 1)
        if source.has_style:
            header._style = copy(source._style)
        for r in range(header_row + 1, last_row + 1):
            src = ws.cell(r, template_col)
            if src.has_style:
                ws.cell(r, col)._style = copy(src._style)
            if is_date:
                ws.cell(r, col).number_format = DATE_FORMAT
        letter = get_column_letter(col)
        ws.column_dimensions[letter].width = 16 if is_date else 30
        if ws.auto_filter.ref:
            start = ws.auto_filter.ref.split(":")[0]
            ws.auto_filter.ref = f"{start}:{letter}{last_row}"

    @staticmethod
    def _add_reply_status(ws, header_row, last_row, status_col) -> None:
        letter = get_column_letter(status_col)
        for dv in ws.data_validations.dataValidation:
            if dv.type == "list" and dv.formula1 and STATUS_SENT in dv.formula1 and _covers(dv.sqref, letter):
                if STATUS_REPLIED not in dv.formula1:
                    items = dv.formula1.strip('"').split(",")
                    pos = items.index(STATUS_FOLLOWED_UP) + 1 if STATUS_FOLLOWED_UP in items else len(items)
                    items.insert(pos, STATUS_REPLIED)
                    dv.formula1 = '"' + ",".join(items) + '"'
        rng = f"{letter}{header_row + 1}:{letter}{last_row}"
        for cf in ws.conditional_formatting:
            if _covers(cf.sqref, letter):
                rng = str(cf.sqref)
                if any(STATUS_REPLIED in str(rule.formula) for rule in cf.rules):
                    return
        first = rng.split(":")[0]
        first_row = re.sub(r"[A-Z]+", "", first)
        ws.conditional_formatting.add(
            rng,
            FormulaRule(formula=[f'${letter}{first_row}="{STATUS_REPLIED}"'], fill=PatternFill(bgColor=REPLY_FILL)),
        )

    def _update_followup_formulas(self, ws, header_row, last_row, columns) -> None:
        """Make 'Date de relance' = last contact (sending or follow-up) + reminder days."""
        f_col = columns["followup_date"]
        c = get_column_letter(columns["contact_date"])
        s = get_column_letter(columns["followup_sent"])
        for r in range(header_row + 1, last_row + 1):
            cell = ws.cell(r, f_col)
            value = cell.value
            if not isinstance(value, str) or not value.startswith("="):
                continue
            original = re.fullmatch(rf'=IF\({c}{r}="","",{c}{r}\+\d+\)', value)
            ours = re.fullmatch(rf'=IF\({c}{r}="","",IF\({s}{r}="",{c}{r},{s}{r}\)\+\d+\)', value)
            if original or ours:
                cell.value = f'=IF({c}{r}="","",IF({s}{r}="",{c}{r},{s}{r})+{self.reminder_days})'

    def _update_side_sheets(self, ws) -> None:
        """Keep the summary and legend sheets consistent with the tool's changes."""
        sans_reponse_row = total_row = None
        has_replied_row = False
        for row in ws.iter_rows(min_col=1, max_col=2):
            a, b = row[0], row[1]
            if isinstance(a.value, str):
                if isinstance(b.value, str) and b.value.upper().startswith("=COUNTIF("):
                    if a.value == STATUS_NO_ANSWER:
                        sans_reponse_row = a.row
                    if a.value == STATUS_REPLIED:
                        has_replied_row = True
                if a.value.strip().lower() == "total" and isinstance(b.value, str) and b.value.upper().startswith("=SUM("):
                    total_row = a.row
                # Legend texts
                text = a.value
                new = re.sub(r"date de contact \+ \d+ jours", f"dernier envoi (candidature ou relance) + {self.reminder_days} jours", text)
                if new != text:
                    a.value = new
            if isinstance(b.value, str):
                text = b.value
                new = re.sub(r"date de contact \+ \d+ jours", f"date du dernier envoi (candidature ou relance) + {self.reminder_days} jours", text)
                if f"{STATUS_FOLLOWED_UP}, " in new and STATUS_REPLIED not in new and STATUS_SENT in new:
                    new = new.replace(f"{STATUS_FOLLOWED_UP}, ", f"{STATUS_FOLLOWED_UP}, {STATUS_REPLIED}, ", 1)
                if new != text:
                    b.value = new

        if has_replied_row or not sans_reponse_row or not total_row or total_row <= sans_reponse_row:
            return
        insert_at = sans_reponse_row + 1
        template_formula = ws.cell(sans_reponse_row, 2).value
        ws.insert_rows(insert_at)
        for col in (1, 2):
            src = ws.cell(sans_reponse_row, col)
            dst = ws.cell(insert_at, col)
            if src.has_style:
                dst._style = copy(src._style)
        ws.cell(insert_at, 1, STATUS_REPLIED)
        ws.cell(insert_at, 2, template_formula.replace(f"A{sans_reponse_row})", f"A{insert_at})"))
        # Rows below moved down by one: fix the total and keep row heights simple.
        total_cell = ws.cell(total_row + 1, 2)
        total_cell.value = re.sub(
            r"=SUM\(B(\d+):B(\d+)\)",
            lambda m: f"=SUM(B{m.group(1)}:B{int(m.group(2)) + 1})",
            total_cell.value,
            flags=re.IGNORECASE,
        )
