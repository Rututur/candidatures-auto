"""Tkinter desktop window: batch sending, follow-ups, settings and Gmail sync.

All user-facing text is French; code, comments and logs-for-developers are English.
Network work runs in a background thread; it talks to the UI through a queue.
"""

from __future__ import annotations

import os
import queue
import re
import shutil
import sys
import threading
import traceback
import tkinter as tk
import webbrowser
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from common.classify import KIND_LABELS, strip_quoted
from common.gmail import GmailClient, gmail_thread_url

from . import templates
from .mailer import GmailSender, OutgoingMail, test_connection
from .settings import Settings, get_password, set_password
from .sync import PendingReply, SentLabeler, record_decision, synchronize
from .workbook import (
    DECISION_BOUNCE,
    DECISION_IGNORE,
    EMAIL_RE,
    STATUS_ACCEPTED,
    STATUS_INTERVIEW,
    STATUS_REFUSED,
    STATUS_REPLIED,
    ApplicationWorkbook,
    Row,
    WorkbookError,
)

MODE_SEND = "send"
MODE_FOLLOWUP = "followup"
FILTER_READY = "Prêtes à envoyer (statut « À contacter » ou « Lettre prête »)"
FILTER_ALL = "Toutes les entreprises avec un email"
CHECKED, UNCHECKED = "☑", "☐"
DECISION_LABELS = {
    STATUS_REFUSED: "Refus",
    STATUS_REPLIED: "Réponse à traiter",
    STATUS_INTERVIEW: "Entretien prévu",
    STATUS_ACCEPTED: "Acceptée",
    DECISION_BOUNCE: "Mail non distribué",
    DECISION_IGNORE: "Ignorer (automatique ou sans rapport)",
}


@dataclass
class Draft:
    to: str
    subject: str
    body: str
    letter: str = ""
    checked: bool = False
    attach_again: bool = True


def fmt_date(value: date | None) -> str:
    return value.strftime("%d/%m/%Y") if value else ""


def strip_reply_prefix(subject: str) -> str:
    return re.sub(r"^(?:\s*(?:re|tr|fwd?)\s*:\s*)+", "", subject or "", flags=re.IGNORECASE).strip()


class MailTab(ttk.Frame):
    """List of companies on the left, editable e-mail on the right."""

    def __init__(self, master, app: "App", mode: str):
        super().__init__(master, padding=8)
        self.app = app
        self.mode = mode
        self.drafts: dict[int, Draft] = {}
        self.rows: dict[int, Row] = {}
        self.current: int | None = None

        paned = ttk.Panedwindow(self, orient="horizontal")
        paned.pack(fill="both", expand=True)

        # -- left: list ---------------------------------------------------
        left = ttk.Frame(paned)
        paned.add(left, weight=1)
        if mode == MODE_SEND:
            self.filter_var = tk.StringVar(value=FILTER_READY)
            combo = ttk.Combobox(left, textvariable=self.filter_var, state="readonly",
                                 values=[FILTER_READY, FILTER_ALL])
            combo.pack(fill="x", pady=(0, 6))
            combo.bind("<<ComboboxSelected>>", lambda e: self.app.refresh_tabs())
            columns = [("check", "", 30), ("company", "Entreprise", 170), ("email", "Email", 180),
                       ("status", "Statut", 85), ("letter", "Lettre", 100)]
        else:
            ttk.Label(left, text="Candidatures sans réponse dont le délai de relance est passé.",
                      foreground="#555").pack(anchor="w", pady=(0, 6))
            columns = [("check", "", 30), ("company", "Entreprise", 170), ("email", "Email", 180),
                       ("last", "Dernier envoi", 90), ("days", "Jours", 45)]

        list_frame = ttk.Frame(left)
        list_frame.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(list_frame, columns=[c[0] for c in columns], show="headings", selectmode="browse")
        for key, title, width in columns:
            self.tree.heading(key, text=title)
            self.tree.column(key, width=width, minwidth=30, stretch=key not in ("check", "days"),
                             anchor="center" if key in ("check", "days", "last") else "w")
        scroll = ttk.Scrollbar(list_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll.set)
        self.tree.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")
        self.tree.bind("<<TreeviewSelect>>", self._on_select)
        self.tree.bind("<Button-1>", self._on_click, add=True)
        self.tree.bind("<space>", lambda e: self._toggle(self.current))
        self.tree.tag_configure("checked", background="#E8F3EC")

        buttons = ttk.Frame(left)
        buttons.pack(fill="x", pady=(6, 0))
        ttk.Button(buttons, text="Tout cocher", command=lambda: self._check_all(True)).pack(side="left")
        ttk.Button(buttons, text="Tout décocher", command=lambda: self._check_all(False)).pack(side="left", padx=4)
        self.count_label = ttk.Label(buttons, text="")
        self.count_label.pack(side="right")

        # -- right: editor ---------------------------------------------------
        right = ttk.Frame(paned, padding=(10, 0, 0, 0))
        paned.add(right, weight=1)
        self.title_label = ttk.Label(right, text="Sélectionne une entreprise dans la liste", style="Title.TLabel")
        self.title_label.grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 8))

        ttk.Label(right, text="À :").grid(row=1, column=0, sticky="w")
        self.to_var = tk.StringVar()
        ttk.Entry(right, textvariable=self.to_var).grid(row=1, column=1, columnspan=2, sticky="ew", pady=2)
        ttk.Label(right, text="Objet :").grid(row=2, column=0, sticky="w")
        self.subject_var = tk.StringVar()
        ttk.Entry(right, textvariable=self.subject_var).grid(row=2, column=1, columnspan=2, sticky="ew", pady=2)

        body_frame = ttk.Frame(right)
        body_frame.grid(row=3, column=0, columnspan=3, sticky="nsew", pady=6)
        self.body = tk.Text(body_frame, wrap="word", undo=True, height=16, font=app.text_font,
                            relief="solid", borderwidth=1, padx=8, pady=6)
        body_scroll = ttk.Scrollbar(body_frame, orient="vertical", command=self.body.yview)
        self.body.configure(yscrollcommand=body_scroll.set)
        self.body.pack(side="left", fill="both", expand=True)
        body_scroll.pack(side="right", fill="y")

        ttk.Label(right, text="Lettre (PDF) :").grid(row=4, column=0, sticky="w")
        self.letter_var = tk.StringVar()
        ttk.Entry(right, textvariable=self.letter_var).grid(row=4, column=1, sticky="ew", pady=2)
        ttk.Button(right, text="Choisir…", command=self._choose_letter).grid(row=4, column=2, padx=(4, 0))

        self.attach_var = tk.BooleanVar(value=True)
        if mode == MODE_FOLLOWUP:
            ttk.Checkbutton(right, text="Joindre à nouveau le CV et la lettre", variable=self.attach_var).grid(
                row=5, column=1, sticky="w", pady=2)
        self.cv_label = ttk.Label(right, text="", foreground="#555")
        self.cv_label.grid(row=6, column=0, columnspan=2, sticky="w", pady=(2, 0))
        ttk.Button(right, text="Rétablir le modèle", command=self._reset_draft).grid(row=6, column=2, sticky="e")

        right.columnconfigure(1, weight=1)
        right.rowconfigure(3, weight=1)
        self._set_editor_state("disabled")

        # -- bottom: action ----------------------------------------------------
        self.action = ttk.Button(self, style="Accent.TButton", command=lambda: self.app.start_batch(self.mode))
        self.action.pack(fill="x", pady=(8, 0), ipady=4)
        self.update_counts()

    # -- list --------------------------------------------------------------
    def wanted_rows(self, workbook: ApplicationWorkbook) -> list[Row]:
        days = self.app.settings.reminder_days
        if self.mode == MODE_FOLLOWUP:
            # A company whose reply is still to be validated has answered: do not follow up.
            waiting = {reply.address for reply in self.app.pending}
            return [r for r in workbook.rows if r.needs_followup(days) and r.email not in waiting]
        if self.filter_var.get() == FILTER_ALL:
            return [r for r in workbook.rows if r.has_valid_email]
        return [r for r in workbook.rows if r.can_be_sent]

    def refresh(self, rows: list[Row]) -> None:
        self._save_editor()
        self.rows = {r.row: r for r in rows}
        self.drafts = {k: v for k, v in self.drafts.items() if k in self.rows}
        self.tree.delete(*self.tree.get_children())
        for row in rows:
            self.tree.insert("", "end", iid=str(row.row), values=self._values(row))
            self._paint(row.row)
        if self.current in self.rows:
            self.tree.selection_set(str(self.current))
        else:
            self.current = None
            self._clear_editor()
        self.update_counts()

    def _values(self, row: Row) -> list[str]:
        draft = self.drafts.get(row.row)
        mark = CHECKED if draft and draft.checked else UNCHECKED
        if self.mode == MODE_FOLLOWUP:
            last = row.last_contact
            days = (date.today() - last).days if last else ""
            return [mark, row.company, row.email, fmt_date(last), days]
        letter = Path(draft.letter).name if draft and draft.letter else (Path(row.letter_file).name if row.letter_file else "—")
        return [mark, row.company, row.email, row.status or "—", letter]

    def _paint(self, row_id: int) -> None:
        draft = self.drafts.get(row_id)
        self.tree.item(str(row_id), values=self._values(self.rows[row_id]),
                       tags=("checked",) if draft and draft.checked else ())

    def _on_click(self, event) -> None:
        if self.tree.identify_region(event.x, event.y) == "cell" and self.tree.identify_column(event.x) == "#1":
            item = self.tree.identify_row(event.y)
            if item:
                self.tree.selection_set(item)
                self._on_select()
                self._toggle(int(item))
                return "break"

    def _toggle(self, row_id: int | None) -> None:
        if row_id is None:
            return
        draft = self._draft(row_id)
        draft.checked = not draft.checked
        self._paint(row_id)
        self.update_counts()

    def _check_all(self, value: bool) -> None:
        for row_id in self.rows:
            self._draft(row_id).checked = value
            self._paint(row_id)
        self.update_counts()

    def update_counts(self) -> None:
        n = sum(1 for d in self.drafts.values() if d.checked)
        test = " — MODE TEST" if self.app.settings.test_mode else ""
        noun = "candidature" if self.mode == MODE_SEND else "relance"
        plural = "s" if n > 1 else ""
        self.action.configure(text=f"Envoyer {n} {noun}{plural} cochée{plural}{test}")
        self.count_label.configure(text=f"{len(self.rows)} entreprise{'s' if len(self.rows) > 1 else ''}")
        self.cv_label.configure(text=f"CV joint : {Path(self.app.settings.cv_path).name or 'non défini (Réglages)'}")

    # -- editor -------------------------------------------------------------
    def _draft(self, row_id: int) -> Draft:
        if row_id not in self.drafts:
            self.drafts[row_id] = self.app.make_draft(self.rows[row_id], self.mode)
        return self.drafts[row_id]

    def _on_select(self, _event=None) -> None:
        selection = self.tree.selection()
        if not selection:
            return
        row_id = int(selection[0])
        if row_id == self.current:
            return
        self._save_editor()
        self.current = row_id
        self._load_editor(self._draft(row_id))

    def _set_editor_state(self, state: str) -> None:
        self.body.configure(state="normal" if state == "normal" else "disabled",
                            background="white" if state == "normal" else "#F2F2F2")

    def _load_editor(self, draft: Draft) -> None:
        row = self.rows[self.current]
        extra = f" — {row.contact}" if row.contact else ""
        self.title_label.configure(text=f"{row.company}{extra}")
        self.to_var.set(draft.to)
        self.subject_var.set(draft.subject)
        self._set_editor_state("normal")
        self.body.delete("1.0", "end")
        self.body.insert("1.0", draft.body)
        self.body.edit_reset()
        self.letter_var.set(draft.letter)
        self.attach_var.set(draft.attach_again)

    def _clear_editor(self) -> None:
        self.title_label.configure(text="Sélectionne une entreprise dans la liste")
        for var in (self.to_var, self.subject_var, self.letter_var):
            var.set("")
        self._set_editor_state("normal")
        self.body.delete("1.0", "end")
        self._set_editor_state("disabled")

    def _save_editor(self) -> None:
        if self.current is None or self.current not in self.drafts:
            return
        draft = self.drafts[self.current]
        draft.to = self.to_var.get().strip()
        draft.subject = self.subject_var.get().strip()
        draft.body = self.body.get("1.0", "end-1c")
        draft.letter = self.letter_var.get().strip()
        draft.attach_again = self.attach_var.get()
        if self.current in self.rows:
            self._paint(self.current)

    def _choose_letter(self) -> None:
        if self.current is None:
            return
        settings = self.app.settings
        path = filedialog.askopenfilename(
            parent=self, title=f"Lettre de motivation pour {self.rows[self.current].company}",
            initialdir=settings.last_letter_dir or str(Path.home()),
            filetypes=[("PDF", "*.pdf"), ("Tous les fichiers", "*.*")],
        )
        if path:
            self.letter_var.set(os.path.normpath(path))
            settings.last_letter_dir = os.path.dirname(path)
            settings.save()
            self._save_editor()

    def _reset_draft(self) -> None:
        if self.current is None:
            return
        old = self.drafts[self.current]
        fresh = self.app.make_draft(self.rows[self.current], self.mode)
        fresh.checked, fresh.letter = old.checked, self.letter_var.get().strip() or fresh.letter
        self.drafts[self.current] = fresh
        self._load_editor(fresh)

    def checked_items(self) -> list[tuple[Row, Draft]]:
        self._save_editor()
        return [(self.rows[k], d) for k, d in self.drafts.items() if d.checked and k in self.rows]

    def uncheck(self, row_id: int) -> None:
        if row_id in self.drafts:
            self.drafts[row_id].checked = False


class ReviewTab(ttk.Frame):
    """Replies waiting for a decision: the estimate only applies once the user approves it."""

    def __init__(self, master, app: "App"):
        super().__init__(master, padding=8)
        self.app = app
        self.replies: dict[int, PendingReply] = {}
        self.current: int | None = None

        paned = ttk.Panedwindow(self, orient="horizontal")
        paned.pack(fill="both", expand=True)

        # -- left: list ---------------------------------------------------
        left = ttk.Frame(paned)
        paned.add(left, weight=1)
        ttk.Label(left, text="Le tableur ne change que lorsque tu valides une réponse.",
                  foreground="#555").pack(anchor="w", pady=(0, 6))
        list_frame = ttk.Frame(left)
        list_frame.pack(fill="both", expand=True)
        columns = [("company", "Entreprise", 170), ("received", "Reçu le", 80), ("estimate", "Estimation", 210)]
        self.tree = ttk.Treeview(list_frame, columns=[c[0] for c in columns], show="headings", selectmode="browse")
        for key, title, width in columns:
            self.tree.heading(key, text=title)
            self.tree.column(key, width=width, minwidth=30, anchor="center" if key == "received" else "w")
        scroll = ttk.Scrollbar(list_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll.set)
        self.tree.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")
        self.tree.bind("<<TreeviewSelect>>", self._on_select)

        # -- right: the mail and the decision ---------------------------------
        right = ttk.Frame(paned, padding=(10, 0, 0, 0))
        paned.add(right, weight=1)
        self.title_label = ttk.Label(right, style="Title.TLabel")
        self.title_label.pack(anchor="w")
        self.info_label = ttk.Label(right, foreground="#555", justify="left", wraplength=520)
        self.info_label.pack(anchor="w", pady=(2, 6))

        text_frame = ttk.Frame(right)
        text_frame.pack(fill="both", expand=True)
        self.body = tk.Text(text_frame, wrap="word", state="disabled", height=14, font=app.text_font,
                            relief="solid", borderwidth=1, padx=8, pady=6)
        body_scroll = ttk.Scrollbar(text_frame, orient="vertical", command=self.body.yview)
        self.body.configure(yscrollcommand=body_scroll.set)
        self.body.pack(side="left", fill="both", expand=True)
        body_scroll.pack(side="right", fill="y")

        self.gmail_button = ttk.Button(right, text="Ouvrir dans Gmail", command=self._open_in_gmail)
        self.gmail_button.pack(anchor="e", pady=6)
        self.approve_button = ttk.Button(right, style="Accent.TButton", command=self._approve)
        self.approve_button.pack(fill="x", ipady=4)
        ttk.Label(right, text="Ou corrige l'estimation :").pack(anchor="w", pady=(8, 2))
        self.correction_buttons: dict[str, ttk.Button] = {}
        grid = ttk.Frame(right)
        grid.pack(fill="x")
        for index, (decision, label) in enumerate(DECISION_LABELS.items()):
            button = ttk.Button(grid, text=label, command=lambda d=decision: self._decide(d))
            button.grid(row=index // 2, column=index % 2, sticky="ew", padx=2, pady=2)
            self.correction_buttons[decision] = button
        grid.columnconfigure((0, 1), weight=1, uniform="decisions")
        self._show(None)

    def refresh(self, replies: list[PendingReply]) -> None:
        self.replies = {r.uid: r for r in replies}
        self.tree.delete(*self.tree.get_children())
        for reply in replies:
            self.tree.insert("", "end", iid=str(reply.uid),
                             values=[reply.company, fmt_date(reply.received), KIND_LABELS[reply.kind]])
        # Keep the current mail selected, otherwise move on to the next one.
        if self.current not in self.replies:
            self.current = replies[0].uid if replies else None
        if self.current is None:
            self._show(None)
        else:
            self.tree.selection_set(str(self.current))
            self._show(self.replies[self.current])

    def _on_select(self, _event=None) -> None:
        selection = self.tree.selection()
        if selection and int(selection[0]) != self.current:
            self.current = int(selection[0])
            self._show(self.replies[self.current])

    def _show(self, reply: PendingReply | None) -> None:
        state = ["!disabled"] if reply else ["disabled"]
        for button in (self.gmail_button, self.approve_button, *self.correction_buttons.values()):
            button.state(state)
        self.title_label.configure(text=reply.company if reply else "Aucune réponse à valider")
        self.info_label.configure(text=(
            f"De : {reply.sender}\nObjet : {reply.subject or '(sans objet)'}\nReçu le : {fmt_date(reply.received)}"
            if reply else ""
        ))
        self.approve_button.configure(
            text=f"Approuver : {DECISION_LABELS[reply.suggestion]}" if reply else "Approuver")
        self.body.configure(state="normal")
        self.body.delete("1.0", "end")
        if reply:
            self.body.insert("1.0", strip_quoted(reply.text) or reply.text)
        self.body.configure(state="disabled")

    def _open_in_gmail(self) -> None:
        if self.current is not None:
            webbrowser.open(gmail_thread_url(self.app.settings.gmail_address, self.replies[self.current].thread_id))

    def _approve(self) -> None:
        if self.current is not None:
            self._decide(self.replies[self.current].suggestion)

    def _decide(self, decision: str) -> None:
        if self.current is not None:
            self.app.decide(self.replies[self.current], decision)


class SettingsDialog(tk.Toplevel):
    def __init__(self, app: "App", welcome: bool = False):
        super().__init__(app)
        self.app = app
        self.title("Réglages")
        self.transient(app)
        self.resizable(True, True)
        self.minsize(640, 520)
        s = app.settings

        frame = ttk.Frame(self, padding=12)
        frame.pack(fill="both", expand=True)
        if welcome:
            ttk.Label(frame, text="Bienvenue ! Renseigne ton compte Gmail et ton CV pour commencer.",
                      style="Title.TLabel").pack(anchor="w", pady=(0, 8))
        notebook = ttk.Notebook(frame)
        notebook.pack(fill="both", expand=True)

        self.vars: dict[str, tk.Variable] = {}

        def entry(parent, row, label, key, value, show=None, browse=None, hint=None):
            ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", pady=3, padx=(0, 8))
            var = tk.StringVar(value=str(value))
            self.vars[key] = var
            ttk.Entry(parent, textvariable=var, show=show).grid(row=row, column=1, sticky="ew", pady=3)
            if browse:
                ttk.Button(parent, text="Choisir…", command=browse).grid(row=row, column=2, padx=(4, 0))
            if hint:
                ttk.Label(parent, text=hint, foreground="#666", wraplength=380, justify="left").grid(
                    row=row + 1, column=1, sticky="w", pady=(0, 4))

        account = ttk.Frame(notebook, padding=10)
        notebook.add(account, text="Compte et CV")
        entry(account, 0, "Adresse Gmail", "gmail_address", s.gmail_address)
        entry(account, 1, "Mot de passe d'application", "password", app.password, show="•",
              hint="16 caractères, créé sur myaccount.google.com/apppasswords (validation en 2 étapes requise). "
                   "Il est rangé dans le gestionnaire d'identifiants de Windows, pas dans un fichier.")
        entry(account, 3, "Ton nom (signature)", "sender_name", s.sender_name)
        entry(account, 4, "Téléphone (signature)", "phone", s.phone)
        entry(account, 5, "CV (PDF)", "cv_path", s.cv_path, browse=self._browse_cv)
        entry(account, 6, "Période du stage", "period", s.period,
              hint="Insérée à la place de {periode} dans les mails.")
        ttk.Button(account, text="Tester la connexion Gmail", command=self._test).grid(row=8, column=1, sticky="w", pady=8)
        self.test_label = ttk.Label(account, text="")
        self.test_label.grid(row=9, column=1, sticky="w")
        account.columnconfigure(1, weight=1)

        sending = ttk.Frame(notebook, padding=10)
        notebook.add(sending, text="Envoi et relances")
        entry(sending, 0, "Pause entre deux mails (s)", "delay_seconds", s.delay_seconds,
              hint="Évite d'être classé comme envoi en masse. 20 à 60 s, c'est bien.")
        entry(sending, 2, "Relancer après (jours)", "reminder_days", s.reminder_days,
              hint="Utilisé pour l'onglet « À relancer » et la colonne « Date de relance » du tableur. "
                   "Pense à mettre la même valeur (REMINDER_DAYS) sur GitHub pour les rappels Discord.")
        entry(sending, 4, "Adresse du mode test", "test_address", s.test_address,
              hint="En mode test, tous les mails partent ici (vide = ton adresse Gmail) et le tableur n'est pas modifié.")
        sending.columnconfigure(1, weight=1)

        tpl = ttk.Frame(notebook, padding=10)
        notebook.add(tpl, text="Modèles de mail")
        ttk.Label(tpl, text="Variables : " + templates.PLACEHOLDER_HELP, foreground="#555", justify="left").grid(
            row=0, column=0, columnspan=2, sticky="w", pady=(0, 6))
        self.texts: dict[str, tk.Text] = {}
        rows = [("subject_template", "Objet (candidature)", 1), ("body_template", "Message (candidature)", 9),
                ("followup_subject_template", "Objet (relance)", 1), ("followup_body_template", "Message (relance)", 9)]
        for i, (key, label, height) in enumerate(rows, start=1):
            ttk.Label(tpl, text=label).grid(row=i, column=0, sticky="nw", padx=(0, 8), pady=3)
            text = tk.Text(tpl, height=height, wrap="word", font=app.text_font, relief="solid", borderwidth=1)
            text.insert("1.0", getattr(s, key))
            text.grid(row=i, column=1, sticky="nsew", pady=3)
            self.texts[key] = text
            tpl.rowconfigure(i, weight=height)
        ttk.Button(tpl, text="Remettre les modèles par défaut", command=self._default_templates).grid(
            row=5, column=1, sticky="e", pady=4)
        tpl.columnconfigure(1, weight=1)

        buttons = ttk.Frame(frame)
        buttons.pack(fill="x", pady=(10, 0))
        ttk.Button(buttons, text="Enregistrer", style="Accent.TButton", command=self._save).pack(side="right")
        ttk.Button(buttons, text="Annuler", command=self.destroy).pack(side="right", padx=6)
        self.grab_set()

    def _browse_cv(self) -> None:
        path = filedialog.askopenfilename(parent=self, title="Ton CV", filetypes=[("PDF", "*.pdf")])
        if path:
            self.vars["cv_path"].set(os.path.normpath(path))

    def _default_templates(self) -> None:
        defaults = {
            "subject_template": templates.DEFAULT_SUBJECT,
            "body_template": templates.DEFAULT_BODY,
            "followup_subject_template": templates.DEFAULT_FOLLOWUP_SUBJECT,
            "followup_body_template": templates.DEFAULT_FOLLOWUP_BODY,
        }
        for key, value in defaults.items():
            self.texts[key].delete("1.0", "end")
            self.texts[key].insert("1.0", value)

    def _test(self) -> None:
        address = self.vars["gmail_address"].get().strip()
        password = self.vars["password"].get().strip()
        self.test_label.configure(text="Connexion en cours…", foreground="#555")

        def work():
            try:
                test_connection(address, password)
                result = ("Connexion réussie (envoi et lecture).", "#1E7B34")
            except Exception as exc:
                result = (f"Échec : {exc}", "#B42318")
            self.after(0, lambda: self.test_label.configure(text=result[0], foreground=result[1]))

        threading.Thread(target=work, daemon=True).start()

    def _save(self) -> None:
        s = self.app.settings
        try:
            delay = int(self.vars["delay_seconds"].get())
            days = int(self.vars["reminder_days"].get())
        except ValueError:
            messagebox.showerror("Réglages", "La pause et le délai de relance doivent être des nombres entiers.", parent=self)
            return
        address = self.vars["gmail_address"].get().strip().lower()
        if address and not EMAIL_RE.match(address):
            messagebox.showerror("Réglages", "L'adresse Gmail n'a pas l'air valide.", parent=self)
            return
        s.gmail_address = address
        for key in ("sender_name", "phone", "cv_path", "period", "test_address"):
            setattr(s, key, self.vars[key].get().strip())
        s.delay_seconds, s.reminder_days = max(0, delay), max(1, days)
        for key, text in self.texts.items():
            setattr(s, key, text.get("1.0", "end-1c").strip())
        password = self.vars["password"].get().strip()
        if address and password and password != self.app.password:
            try:
                set_password(address, password)
            except Exception as exc:
                messagebox.showerror("Réglages", f"Impossible d'enregistrer le mot de passe : {exc}", parent=self)
                return
        self.app.password = password
        s.save()
        self.destroy()
        self.app.on_settings_changed()


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Candidatures — envoi, relances et suivi")
        self.geometry("1200x740")
        self.minsize(980, 600)
        self.settings = Settings.load()
        self.password = get_password(self.settings.gmail_address)
        self.workbook: ApplicationWorkbook | None = None
        self.applications: dict = {}
        self.pending: list[PendingReply] = []
        self.events: queue.Queue = queue.Queue()
        self.busy = False
        self.cancel_event = threading.Event()
        self.pending_marks: list[tuple] = []
        self._setup_style()
        self._build()
        self.after(100, self._poll)
        self.after(200, self._startup)

    def report_callback_exception(self, exc, value, tb) -> None:
        """Show unexpected errors instead of failing silently (pythonw has no console)."""
        details = "".join(traceback.format_exception(exc, value, tb))
        self._append_log(f"Erreur inattendue : {value}")
        messagebox.showerror("Erreur inattendue", f"{value}\n\nDétails dans l'onglet Journal.")
        self._append_log(details)

    # -- layout ------------------------------------------------------------
    def _setup_style(self) -> None:
        style = ttk.Style(self)
        if sys.platform != "win32" and "clam" in style.theme_names():
            style.theme_use("clam")
        base = "Segoe UI" if sys.platform == "win32" else "TkDefaultFont"
        self.text_font = ("Segoe UI", 10) if sys.platform == "win32" else ("DejaVu Sans", 10)
        style.configure("Title.TLabel", font=(base, 12, "bold"))
        style.configure("Accent.TButton", font=(base, 10, "bold"))
        style.configure("Test.TCheckbutton", foreground="#B54708", font=(base, 10, "bold"))
        style.configure("Real.TCheckbutton", foreground="#1E7B34", font=(base, 10, "bold"))

    def _build(self) -> None:
        top = ttk.Frame(self, padding=(10, 8))
        top.pack(fill="x")
        ttk.Label(top, text="Tableur :").pack(side="left")
        self.path_var = tk.StringVar(value="aucun")
        ttk.Label(top, textvariable=self.path_var, foreground="#333").pack(side="left", padx=(4, 8))
        ttk.Button(top, text="Ouvrir…", command=self.choose_workbook).pack(side="left")
        ttk.Button(top, text="Recharger", command=self.reload_workbook).pack(side="left", padx=4)
        ttk.Button(top, text="Modifier dans Excel", command=self.open_in_excel).pack(side="left")
        ttk.Button(top, text="Enregistrer une copie…", command=self.save_workbook_copy).pack(side="left", padx=4)
        self.sync_button = ttk.Button(top, text="Synchroniser avec Gmail", command=lambda: self.run_sync(False))
        self.sync_button.pack(side="left", padx=4)
        ttk.Button(top, text="Réglages", command=lambda: SettingsDialog(self)).pack(side="right")
        self.test_var = tk.BooleanVar(value=self.settings.test_mode)
        self.test_check = ttk.Checkbutton(top, variable=self.test_var, command=self._toggle_test)
        self.test_check.pack(side="right", padx=12)

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=10)
        self.send_tab = MailTab(self.notebook, self, MODE_SEND)
        self.followup_tab = MailTab(self.notebook, self, MODE_FOLLOWUP)
        self.review_tab = ReviewTab(self.notebook, self)
        log_frame = ttk.Frame(self.notebook, padding=8)
        self.log_text = tk.Text(log_frame, wrap="word", state="disabled", font=self.text_font,
                                relief="solid", borderwidth=1, padx=8, pady=6)
        self.log_text.pack(fill="both", expand=True)
        self.notebook.add(self.send_tab, text="  Envoyer  ")
        self.notebook.add(self.followup_tab, text="  À relancer  ")
        self.notebook.add(self.review_tab, text="  À valider  ")
        self.notebook.add(log_frame, text="  Journal  ")

        bottom = ttk.Frame(self, padding=(10, 6))
        bottom.pack(fill="x")
        self.status_var = tk.StringVar(value="Prêt.")
        ttk.Label(bottom, textvariable=self.status_var).pack(side="left")
        self.cancel_button = ttk.Button(bottom, text="Arrêter après ce mail", command=self.cancel_event.set)
        self.progress = ttk.Progressbar(bottom, length=240, mode="determinate")
        self._update_test_label()

    def _update_test_label(self) -> None:
        if self.settings.test_mode:
            self.test_check.configure(style="Test.TCheckbutton",
                                      text=f"Mode test : tout part vers {self.settings.effective_test_address or 'ton adresse'}")
        else:
            self.test_check.configure(style="Real.TCheckbutton", text="Mode test désactivé : envois réels")
        for tab in (self.send_tab, self.followup_tab):
            tab.update_counts()

    def _toggle_test(self) -> None:
        if not self.test_var.get() and not messagebox.askyesno(
            "Envois réels", "Désactiver le mode test ? Les prochains mails partiront vraiment aux entreprises."
        ):
            self.test_var.set(True)
            return
        self.settings.test_mode = self.test_var.get()
        self.settings.save()
        self._update_test_label()

    # -- logging / events -------------------------------------------------
    def log(self, message: str) -> None:
        """Thread-safe: may be called from the worker thread."""
        self.events.put(("log", message))

    def _append_log(self, message: str) -> None:
        self.log_text.configure(state="normal")
        self.log_text.insert("end", f"{datetime.now():%H:%M:%S}  {message}\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _poll(self) -> None:
        try:
            while True:
                kind, *payload = self.events.get_nowait()
                handler = getattr(self, f"_on_{kind}")
                handler(*payload)
        except queue.Empty:
            pass
        self.after(100, self._poll)

    def _on_log(self, message: str) -> None:
        self._append_log(message)
        self.status_var.set(message)

    def _on_progress(self, done: int, total: int) -> None:
        self.progress.configure(maximum=max(total, 1), value=done)

    def _on_call(self, func, *args) -> None:
        func(*args)

    # -- startup / workbook -----------------------------------------------
    def _startup(self) -> None:
        if not self.settings.gmail_address or not self.password or not self.settings.cv_path:
            SettingsDialog(self, welcome=True)
        if self.settings.workbook_path and Path(self.settings.workbook_path).exists():
            self.open_workbook(self.settings.workbook_path, sync=True)
        else:
            self.after(300, self.choose_workbook)

    def choose_workbook(self) -> None:
        path = filedialog.askopenfilename(
            parent=self, title="Ton tableau de suivi des candidatures",
            filetypes=[("Excel", "*.xlsx *.xlsm"), ("Tous les fichiers", "*.*")],
        )
        if path:
            self.open_workbook(path, sync=True)

    def open_workbook(self, path: str, sync: bool = False) -> None:
        try:
            self.workbook = ApplicationWorkbook(path, self.settings.reminder_days)
        except WorkbookError as exc:
            messagebox.showerror("Tableur", str(exc))
            return
        self.settings.workbook_path = os.path.normpath(path)
        self.settings.save()
        self.path_var.set(Path(path).name)
        self._append_log(f"Tableur ouvert : {Path(path).name} — {len(self.workbook.rows)} entreprises.")
        self.refresh_tabs()
        if sync and self.settings.gmail_address and self.password:
            self.run_sync(silent=True)

    def reload_workbook(self) -> None:
        if self.workbook and not self.busy:
            try:
                self.workbook.load()
            except WorkbookError as exc:
                messagebox.showerror("Tableur", str(exc))
                return
            self.refresh_tabs()
            self._append_log("Tableur rechargé.")

    def open_in_excel(self) -> None:
        if not self.workbook:
            return
        try:
            os.startfile(self.workbook.path)
        except OSError as exc:
            messagebox.showerror("Tableur", f"Impossible d'ouvrir le fichier : {exc}")
            return
        self.status_var.set("Modifie le tableur, enregistre et ferme Excel, puis clique sur « Recharger ».")

    def save_workbook_copy(self) -> None:
        if not self.workbook:
            return
        source = self.workbook.path
        target = filedialog.asksaveasfilename(
            parent=self, title="Enregistrer une copie du tableur", initialfile=f"{source.stem} (copie){source.suffix}",
            defaultextension=source.suffix, filetypes=[("Excel", f"*{source.suffix}")],
        )
        if not target:
            return
        try:
            shutil.copy2(source, target)
        except OSError as exc:
            messagebox.showerror("Tableur", f"Impossible d'enregistrer la copie : {exc}")
            return
        self.status_var.set(f"Copie enregistrée : {target}")

    def refresh_tabs(self) -> None:
        if not self.workbook:
            return
        for tab in (self.send_tab, self.followup_tab):
            tab.refresh(tab.wanted_rows(self.workbook))
        n = len(self.followup_tab.rows)
        self.notebook.tab(self.followup_tab, text=f"  À relancer ({n})  " if n else "  À relancer  ")
        self.review_tab.refresh(self.pending)
        n = len(self.pending)
        self.notebook.tab(self.review_tab, text=f"  À valider ({n})  " if n else "  À valider  ")

    def on_settings_changed(self) -> None:
        self.test_var.set(self.settings.test_mode)
        self._update_test_label()
        if self.workbook:
            self.workbook.reminder_days = self.settings.reminder_days
            self.refresh_tabs()

    # -- drafts --------------------------------------------------------------
    def make_draft(self, row: Row, mode: str) -> Draft:
        s = self.settings
        context = {
            "entreprise": row.company,
            "service": row.service,
            "interlocuteur": row.contact,
            "salutation": templates.salutation(row.contact),
            "periode": s.period,
            "nom": s.sender_name,
            "telephone": s.phone,
            "date_envoi": fmt_date(row.contact_date),
        }
        letter = row.letter_file if row.letter_file and Path(row.letter_file).exists() else ""
        if mode == MODE_SEND:
            subject = templates.render(s.subject_template, context)
            body = templates.render(s.body_template, context)
        else:
            app = self.applications.get(row.email)
            original = strip_reply_prefix(app.first.subject) if app else ""
            context["objet_initial"] = original or templates.render(s.subject_template, context)
            subject = templates.render(s.followup_subject_template, context)
            body = templates.render(s.followup_body_template, context)
        return Draft(to=row.email, subject=subject.strip(), body=templates.tidy(body), letter=letter)

    # -- Gmail sync --------------------------------------------------------
    def _set_busy(self, busy: bool, label: str = "") -> None:
        self.busy = busy
        state = "disabled" if busy else "normal"
        self.sync_button.configure(state=state)
        for tab in (self.send_tab, self.followup_tab):
            tab.action.configure(state=state)
        if busy:
            self.progress.pack(side="right")
            if label:
                self.status_var.set(label)
        else:
            self.progress.pack_forget()
            self.cancel_button.pack_forget()

    def run_sync(self, silent: bool) -> None:
        if self.busy or not self.workbook:
            return
        if not self.settings.gmail_address or not self.password:
            if not silent:
                messagebox.showinfo("Gmail", "Renseigne d'abord ton adresse Gmail et ton mot de passe d'application dans les Réglages.")
            return
        self._set_busy(True, "Synchronisation avec Gmail…")
        self.progress.configure(mode="indeterminate")
        self.progress.start(12)

        def work():
            try:
                report = synchronize(self.settings.gmail_address, self.password, self.workbook, self.log)
                self.events.put(("call", self._sync_done, report, None, silent))
            except Exception as exc:
                self.events.put(("call", self._sync_done, None, exc, silent))

        threading.Thread(target=work, daemon=True).start()

    def _sync_done(self, report, error, silent: bool) -> None:
        self.progress.stop()
        self.progress.configure(mode="determinate")
        self._set_busy(False)
        if error:
            self._append_log(f"Synchronisation impossible : {error}")
            self.status_var.set("Synchronisation impossible (voir Journal).")
            if not silent or isinstance(error, WorkbookError):
                messagebox.showerror("Synchronisation", str(error))
            return
        self.applications = report.applications
        self.pending = report.pending
        summary = (f"Synchronisé : {len(report.applications)} candidature(s) suivie(s), "
                   f"{report.replies_found} réponse(s) trouvée(s), {len(report.pending)} à valider.")
        self._append_log(summary)
        self.status_var.set(summary)
        self.refresh_tabs()
        if report.pending:
            self.notebook.select(self.review_tab)
        elif not silent:
            messagebox.showinfo("Synchronisation", summary)

    # -- reply validation ---------------------------------------------------
    def decide(self, reply: PendingReply, decision: str) -> None:
        if self.busy or not self.workbook:
            return
        self._set_busy(True, "Enregistrement de ta décision…")
        self.progress.configure(mode="indeterminate")
        self.progress.start(12)

        def work():
            try:
                record_decision(self.settings.gmail_address, self.password, self.workbook, reply, decision)
                self.events.put(("call", self._decision_done, reply, decision, None))
            except Exception as exc:
                self.events.put(("call", self._decision_done, reply, decision, exc))

        threading.Thread(target=work, daemon=True).start()

    def _decision_done(self, reply: PendingReply, decision: str, error) -> None:
        self.progress.stop()
        self.progress.configure(mode="determinate")
        self._set_busy(False)
        if error:
            self._append_log(f"Décision non enregistrée pour {reply.company} : {error}")
            self.status_var.set("Décision non enregistrée (voir Journal).")
            messagebox.showerror("Validation", str(error))
            return
        self.pending = [r for r in self.pending if r.uid != reply.uid]
        summary = f"{reply.company} : {DECISION_LABELS[decision]} enregistré."
        self._append_log(summary)
        self.status_var.set(summary)
        self.refresh_tabs()

    # -- batch sending ------------------------------------------------------
    def _validate(self, mode: str, items) -> list[str]:
        s = self.settings
        errors = []
        if not s.gmail_address or not self.password:
            errors.append("Adresse Gmail ou mot de passe d'application manquant (Réglages).")
        cv = Path(s.cv_path) if s.cv_path else None
        needs_cv = mode == MODE_SEND or any(d.attach_again for _, d in items)
        if needs_cv and (not cv or not cv.is_file()):
            errors.append("CV introuvable : choisis-le dans les Réglages.")
        if s.test_mode and not EMAIL_RE.match(s.effective_test_address):
            errors.append("Adresse du mode test invalide (Réglages).")
        for row, draft in items:
            name = row.company
            if not EMAIL_RE.match(draft.to):
                errors.append(f"{name} : adresse « {draft.to} » invalide.")
            if not draft.subject or not draft.body.strip():
                errors.append(f"{name} : objet ou message vide.")
            leftover = templates.unresolved(draft.subject + "\n" + draft.body)
            if leftover:
                errors.append(f"{name} : variable(s) non remplacée(s) : " + ", ".join("{" + x + "}" for x in leftover))
            if mode == MODE_SEND or draft.attach_again:
                letter = Path(draft.letter) if draft.letter else None
                if not letter or not letter.is_file():
                    errors.append(f"{name} : lettre de motivation non choisie ou introuvable.")
                elif letter.suffix.lower() != ".pdf":
                    errors.append(f"{name} : la lettre doit être un PDF.")
        return errors

    def start_batch(self, mode: str) -> None:
        if self.busy or not self.workbook:
            return
        tab = self.send_tab if mode == MODE_SEND else self.followup_tab
        items = tab.checked_items()
        if not items:
            messagebox.showinfo("Envoi", "Coche au moins une entreprise dans la liste (case à gauche).")
            return
        errors = self._validate(mode, items)
        if errors:
            messagebox.showerror("À corriger avant l'envoi", "\n".join(errors[:20]) + ("\n…" if len(errors) > 20 else ""))
            return
        test = self.settings.test_mode
        if not test:
            try:
                self.workbook.check_writable()
            except WorkbookError as exc:
                messagebox.showerror("Tableur", str(exc))
                return
        noun = "candidature(s)" if mode == MODE_SEND else "relance(s)"
        names = "\n".join(f"• {row.company} ({draft.to})" for row, draft in items[:15])
        if len(items) > 15:
            names += f"\n… et {len(items) - 15} autre(s)"
        where = (f"MODE TEST : tout part vers {self.settings.effective_test_address}, le tableur ne sera pas modifié."
                 if test else "ENVOI RÉEL aux entreprises. Le tableur sera mis à jour après chaque mail.")
        minutes = (len(items) - 1) * self.settings.delay_seconds / 60
        duration = f"\nDurée estimée : environ {minutes:.0f} min." if minutes >= 1 else ""
        if not messagebox.askyesno("Confirmer l'envoi", f"Envoyer {len(items)} {noun} ?\n\n{names}\n\n{where}{duration}"):
            return
        self.cancel_event.clear()
        self._set_busy(True, "Envoi en cours…")
        self.cancel_button.pack(side="right", padx=6)
        self.progress.configure(maximum=len(items), value=0)
        threading.Thread(target=self._send_worker, args=(mode, items, test), daemon=True).start()

    def _send_worker(self, mode: str, items, test: bool) -> None:
        s = self.settings
        sent = failed = 0
        client = None
        try:
            with GmailSender(s.gmail_address, self.password, s.sender_name) as sender:
                labeler = None
                if not test:
                    client = GmailClient(s.gmail_address, self.password)
                    client.connect()
                    labeler = SentLabeler(client)
                for index, (row, draft) in enumerate(items):
                    if self.cancel_event.is_set():
                        self.log("Envoi arrêté à ta demande.")
                        break
                    attachments = []
                    if mode == MODE_SEND or draft.attach_again:
                        attachments = [Path(s.cv_path), Path(draft.letter)]
                    in_reply_to, references = "", []
                    app = self.applications.get(row.email)
                    if mode == MODE_FOLLOWUP and app and app.message_ids:
                        references = app.message_ids
                        in_reply_to = references[-1]
                    mail = OutgoingMail(
                        to_address=s.effective_test_address if test else draft.to,
                        to_name=row.company,
                        subject=("[TEST] " if test else "") + draft.subject,
                        body=draft.body,
                        attachments=attachments,
                        in_reply_to="" if test else in_reply_to,
                        references=[] if test else references,
                    )
                    try:
                        msg = sender.send(mail)
                    except Exception as exc:
                        failed += 1
                        self.log(f"Échec de l'envoi à {row.company} : {exc}")
                        continue
                    sent += 1
                    self.log(f"Envoyé{' (test)' if test else ''} : {row.company}")
                    if not test:
                        try:
                            if not labeler.label(msg["Message-ID"]):
                                self.log(f"Mail à {row.company} envoyé mais pas encore visible dans Gmail : "
                                         "il sera retrouvé à la prochaine synchronisation.")
                        except Exception as exc:
                            self.log(f"Libellé Gmail non posé pour {row.company} ({exc}) : la synchronisation le rattrapera.")
                        self.events.put(("call", self._record_sent, mode, row.row, draft.letter))
                    self.events.put(("progress", index + 1, len(items)))
                    if index < len(items) - 1 and self.cancel_event.wait(s.delay_seconds):
                        self.log("Envoi arrêté à ta demande.")
                        break
        except Exception as exc:
            self.log(f"Connexion à Gmail impossible : {exc}")
        finally:
            if client:
                client.close()
            self.events.put(("call", self._batch_done, mode, sent, failed, test))

    def _record_sent(self, mode: str, row_id: int, letter: str) -> bool:
        tab = self.send_tab if mode == MODE_SEND else self.followup_tab
        tab.uncheck(row_id)
        try:
            if mode == MODE_SEND:
                self.workbook.mark_sent(row_id, date.today(), letter)
            else:
                self.workbook.mark_followed_up(row_id, date.today())
            return True
        except WorkbookError as exc:
            self.pending_marks.append((mode, row_id, letter))
            self._append_log(f"Tableur non mis à jour : {exc}")
            return False

    def _flush_pending_marks(self) -> None:
        """Retry spreadsheet updates that failed because Excel had the file open."""
        while self.pending_marks:
            pending, self.pending_marks = self.pending_marks, []
            for item in pending:
                self._record_sent(*item)
            if not self.pending_marks:
                return
            names = ", ".join(self.workbook.by_row(r).company for _, r, _ in self.pending_marks)
            if not messagebox.askretrycancel(
                "Tableur à mettre à jour",
                f"Ces envois sont partis mais n'ont pas pu être notés dans le tableur : {names}.\n\n"
                "Ferme le fichier dans Excel puis clique sur « Réessayer ».",
            ):
                messagebox.showwarning(
                    "Tableur à compléter",
                    f"Pense à mettre à jour à la main (statut et date du jour) : {names}.",
                )
                self.pending_marks = []
                return

    def _batch_done(self, mode: str, sent: int, failed: int, test: bool) -> None:
        self._set_busy(False)
        self._flush_pending_marks()
        self.refresh_tabs()
        summary = f"{sent} mail(s) envoyé(s)" + (f", {failed} échec(s)" if failed else "") + (" en mode test." if test else ".")
        self.status_var.set(summary)
        self._append_log(summary)
        hint = "\n\nVa voir ta boîte mail pour vérifier le rendu." if test and sent else ""
        if failed:
            hint += "\n\nLes échecs sont détaillés dans l'onglet Journal ; ces entreprises sont restées cochées."
        (messagebox.showwarning if failed else messagebox.showinfo)("Envoi terminé", summary + hint)
        # Label any mail Gmail had not indexed in time: the watcher only tracks labelled mails.
        if sent and not test:
            self.run_sync(silent=True)


def main() -> None:
    if sys.platform == "win32":
        try:
            import ctypes

            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass
    try:
        App().mainloop()
    except Exception as exc:  # startup failure: pythonw has no console to show it
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("Impossible de démarrer", f"{exc}\n\n{traceback.format_exc()[-1500:]}")
        root.destroy()
