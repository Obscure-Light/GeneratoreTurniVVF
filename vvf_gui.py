from __future__ import annotations

import argparse
import tkinter as tk
import queue
import subprocess
import sys
import threading
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Dict, List, Optional, Set

from database import (
    Database,
    DEFAULT_ACTIVE_WEEKDAYS,
    DEFAULT_AUTISTA_POGLIANI,
    DEFAULT_AUTISTA_VARCHI,
    DEFAULT_MIN_ESPERTI,
    DEFAULT_VIGILE_ESCLUSO_ESTATE,
    DEFAULT_WEEKLY_CAP,
    ROLE_AUTISTA,
    ROLE_AUTISTA_VIGILE,
    ROLE_VIGILE,
)

ROLE_OPTIONS = [ROLE_AUTISTA, ROLE_VIGILE, ROLE_AUTISTA_VIGILE]
GRADE_OPTIONS = ["JUNIOR", "SENIOR", "ALTRO"]
WEEKDAY_LABELS = {
    0: "Lunedì",
    1: "Martedì",
    2: "Mercoledì",
    3: "Giovedì",
    4: "Venerdì",
    5: "Sabato",
    6: "Domenica",
}

LIV_JUNIOR = "JUNIOR"


class SchedulerManagerApp(tk.Tk):
    """Interfaccia grafica per amministrare il database del VVF Scheduler."""

    def __init__(self, db_path: Path):
        super().__init__()
        self.title("VVF Scheduler – Gestione configurazione")
        self.geometry("1120x680")

        self.db_path = Path(db_path)
        self.db = Database(self.db_path)
        self.people_cache: Dict[int, Dict[str, object]] = {}
        self.name_to_id: Dict[str, int] = {}
        self.autisti_names: Set[str] = set()
        self.vigili_names: Set[str] = set()
        self.selected_person_id: Optional[int] = None
        self.generate_running = False
        self.generate_queue: "queue.Queue[str]" = queue.Queue()

        self._build_ui()
        self.refresh_all()
        self.protocol("WM_DELETE_WINDOW", self.on_close)

    # ---------------------- UI Skeleton ---------------------- #
    def _build_ui(self) -> None:
        notebook = ttk.Notebook(self)
        notebook.pack(fill="both", expand=True, padx=10, pady=10)

        self.people_frame = ttk.Frame(notebook)
        notebook.add(self.people_frame, text="Personale")
        self._build_people_tab(self.people_frame)

        self.pairs_frame = ttk.Frame(notebook)
        notebook.add(self.pairs_frame, text="Coppie & Vincoli")
        self._build_pairs_tab(self.pairs_frame)

        self.vacations_frame = ttk.Frame(notebook)
        notebook.add(self.vacations_frame, text="Ferie")
        self._build_vacations_tab(self.vacations_frame)

        self.settings_frame = ttk.Frame(notebook)
        notebook.add(self.settings_frame, text="Impostazioni")
        self._build_settings_tab(self.settings_frame)

        self.generation_frame = ttk.Frame(notebook)
        notebook.add(self.generation_frame, text="Genera turni")
        self._build_generation_tab(self.generation_frame)

    # ---------------------- Tab: Personale ---------------------- #
    def _build_people_tab(self, container: ttk.Frame) -> None:
        container.columnconfigure(0, weight=3)
        container.columnconfigure(1, weight=2)
        container.rowconfigure(0, weight=1)

        # Treeview elenco persone
        tree_frame = ttk.Frame(container)
        tree_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 12))
        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)

        columns = (
            "nome",
            "ruolo",
            "grado",
            "autista",
            "vigile",
            "weekly_cap",
            "telefono",
            "email",
        )
        self.people_tree = ttk.Treeview(
            tree_frame,
            columns=columns,
            show="headings",
            selectmode="browse",
        )
        headers = [
            ("nome", "Nome completo", 200),
            ("ruolo", "Ruolo", 120),
            ("grado", "Grado", 80),
            ("autista", "Autista", 70),
            ("vigile", "Vigile", 70),
            ("weekly_cap", "Turni/sett.", 90),
            ("telefono", "Telefono", 120),
            ("email", "E-mail", 160),
        ]
        for col, label, width in headers:
            self.people_tree.heading(col, text=label)
            self.people_tree.column(col, width=width, anchor="center")
        self.people_tree.grid(row=0, column=0, sticky="nsew")
        self.people_tree.bind("<<TreeviewSelect>>", self.on_person_select)

        scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=self.people_tree.yview)
        self.people_tree.configure(yscrollcommand=scroll.set)
        scroll.grid(row=0, column=1, sticky="ns")

        # Form di dettaglio
        form = ttk.LabelFrame(container, text="Dettaglio")
        form.grid(row=0, column=1, sticky="nsew")
        for i in range(8):
            form.rowconfigure(i, weight=0)
        form.rowconfigure(8, weight=1)
        form.columnconfigure(1, weight=1)

        ttk.Label(form, text="Nome").grid(row=0, column=0, sticky="w", padx=4, pady=4)
        self.person_first_name = tk.StringVar()
        ttk.Entry(form, textvariable=self.person_first_name).grid(row=0, column=1, sticky="ew", padx=4, pady=4)

        ttk.Label(form, text="Cognome").grid(row=1, column=0, sticky="w", padx=4, pady=4)
        self.person_last_name = tk.StringVar()
        ttk.Entry(form, textvariable=self.person_last_name).grid(row=1, column=1, sticky="ew", padx=4, pady=4)

        ttk.Label(form, text="Telefono").grid(row=2, column=0, sticky="w", padx=4, pady=4)
        self.person_phone = tk.StringVar()
        ttk.Entry(form, textvariable=self.person_phone).grid(row=2, column=1, sticky="ew", padx=4, pady=4)

        ttk.Label(form, text="E-mail").grid(row=3, column=0, sticky="w", padx=4, pady=4)
        self.person_email = tk.StringVar()
        ttk.Entry(form, textvariable=self.person_email).grid(row=3, column=1, sticky="ew", padx=4, pady=4)

        ttk.Label(form, text="Ruolo operativo").grid(row=4, column=0, sticky="w", padx=4, pady=4)
        self.person_role = tk.StringVar(value=ROLE_VIGILE)
        self.person_role_combo = ttk.Combobox(
            form, textvariable=self.person_role, state="readonly", values=ROLE_OPTIONS
        )
        self.person_role_combo.grid(row=4, column=1, sticky="ew", padx=4, pady=4)
        self.person_role_combo.bind("<<ComboboxSelected>>", lambda _: self._toggle_grade_state())

        ttk.Label(form, text="Grado").grid(row=5, column=0, sticky="w", padx=4, pady=4)
        self.person_grade = tk.StringVar(value="JUNIOR")
        self.person_grade_combo = ttk.Combobox(
            form, textvariable=self.person_grade, state="readonly", values=GRADE_OPTIONS
        )
        self.person_grade_combo.grid(row=5, column=1, sticky="ew", padx=4, pady=4)

        ttk.Label(form, text="Turni max a settimana (0 = nessun limite)").grid(row=6, column=0, sticky="w", padx=4, pady=4)
        self.person_weekly_cap = tk.IntVar(value=DEFAULT_WEEKLY_CAP)
        ttk.Spinbox(form, from_=0, to=7, textvariable=self.person_weekly_cap, width=5).grid(
            row=6, column=1, sticky="w", padx=4, pady=4
        )

        btn_frame = ttk.Frame(form)
        btn_frame.grid(row=7, column=0, columnspan=2, pady=12)
        ttk.Button(btn_frame, text="Salva", command=self.save_person).grid(row=0, column=0, padx=4)
        ttk.Button(btn_frame, text="Nuovo", command=self.reset_person_form).grid(row=0, column=1, padx=4)
        ttk.Button(btn_frame, text="Elimina", command=self.delete_person).grid(row=0, column=2, padx=4)
        ttk.Button(btn_frame, text="Aggiorna elenco", command=self.refresh_people_list).grid(row=0, column=3, padx=4)

        help_label = ttk.Label(
            form,
            text=(
                "Nota:\n"
                "• Il campo \"Ruolo\" determina automaticamente se la persona è autista, vigile o entrambi.\n"
                "• Il limite settimanale si applica sia agli autisti sia ai vigili; 0 significa nessun limite."
            ),
            foreground="#555555",
            justify="left",
        )
        help_label.grid(row=8, column=0, columnspan=2, sticky="sw", padx=4, pady=(8, 4))

    def _toggle_grade_state(self) -> None:
        ruolo = self.person_role.get()
        if ROLE_VIGILE in ruolo or ROLE_AUTISTA_VIGILE in ruolo:
            self.person_grade_combo.configure(state="readonly")
        else:
            self.person_grade.set("")
            self.person_grade_combo.configure(state="disabled")

    def reset_person_form(self) -> None:
        self.selected_person_id = None
        self.person_first_name.set("")
        self.person_last_name.set("")
        self.person_phone.set("")
        self.person_email.set("")
        self.person_role.set(ROLE_VIGILE)
        self.person_grade.set("JUNIOR")
        self.person_weekly_cap.set(DEFAULT_WEEKLY_CAP)
        self._toggle_grade_state()
        self.people_tree.selection_remove(self.people_tree.selection())

    def on_person_select(self, _event=None) -> None:
        selection = self.people_tree.selection()
        if not selection:
            return
        person_id = int(selection[0])
        row = self.people_cache.get(person_id)
        if not row:
            return
        self.selected_person_id = person_id
        self.person_first_name.set(row["first_name"] or "")
        self.person_last_name.set(row["last_name"] or "")
        self.person_phone.set(row["phone"] or "")
        self.person_email.set(row["email"] or "")
        ruolo = row["ruolo"] or ROLE_VIGILE
        if ruolo not in ROLE_OPTIONS:
            ruolo = ROLE_VIGILE
        self.person_role.set(ruolo)
        if row["is_vigile"]:
            self.person_grade.set(row["grado"] or "JUNIOR")
        else:
            self.person_grade.set("")
        self.person_weekly_cap.set(int(row["weekly_cap"]) if row["weekly_cap"] is not None else DEFAULT_WEEKLY_CAP)
        self._toggle_grade_state()

    def save_person(self) -> None:
        first = self.person_first_name.get().strip()
        last = self.person_last_name.get().strip()
        if not first and not last:
            messagebox.showerror("Errore", "Inserisci almeno il nome o il cognome.")
            return
        display_name = f"{first} {last}".strip()
        ruolo = self.person_role.get()
        grado = self.person_grade.get().strip()
        phone = self.person_phone.get().strip()
        email = self.person_email.get().strip()
        weekly_cap = max(0, self.person_weekly_cap.get())

        is_autista = ruolo in (ROLE_AUTISTA, ROLE_AUTISTA_VIGILE)
        is_vigile = ruolo in (ROLE_VIGILE, ROLE_AUTISTA_VIGILE)
        if is_vigile:
            livello = grado if grado in ("JUNIOR", "SENIOR") else LIV_JUNIOR
            grado_db = grado if grado else LIV_JUNIOR
        else:
            livello = LIV_JUNIOR
            grado_db = ""

        try:
            if self.selected_person_id is None:
                self.db.upsert_person(
                    display_name,
                    first_name=first,
                    last_name=last,
                    phone=phone,
                    email=email,
                    ruolo=ruolo,
                    grado=grado_db,
                    is_autista=is_autista,
                    is_vigile=is_vigile,
                    livello=livello,
                    weekly_cap=weekly_cap,
                )
            else:
                self.db.update_person(
                    self.selected_person_id,
                    name=display_name,
                    first_name=first,
                    last_name=last,
                    phone=phone,
                    email=email,
                    ruolo=ruolo,
                    grado=grado_db,
                    is_autista=is_autista,
                    is_vigile=is_vigile,
                    livello=livello,
                    weekly_cap=weekly_cap,
                )
            self.refresh_people_list()
            self.refresh_pairs_lists()
            self.refresh_settings_inputs()
        except ValueError as exc:
            messagebox.showerror("Errore", str(exc))

    def delete_person(self) -> None:
        if self.selected_person_id is None:
            messagebox.showwarning("Attenzione", "Seleziona prima una persona da eliminare.")
            return
        row = self.people_cache.get(self.selected_person_id)
        nome = row["name"] if row else "la persona selezionata"
        if not messagebox.askyesno("Conferma", f"Eliminare {nome}? L'operazione è irreversibile."):
            return
        self.db.delete_person(self.selected_person_id)
        self.reset_person_form()
        self.refresh_people_list()
        self.refresh_pairs_lists()
        self.refresh_settings_inputs()

    def refresh_people_list(self) -> None:
        self.people_tree.delete(*self.people_tree.get_children())
        self.people_cache.clear()
        self.name_to_id.clear()
        self.autisti_names.clear()
        self.vigili_names.clear()

        for row in self.db.list_people():
            person_id = int(row["id"])
            self.people_cache[person_id] = dict(row)
            self.name_to_id[row["name"]] = person_id
            if row["is_autista"]:
                self.autisti_names.add(row["name"])
            if row["is_vigile"]:
                self.vigili_names.add(row["name"])

            self.people_tree.insert(
                "",
                "end",
                iid=str(person_id),
                values=(
                    row["name"],
                    row["ruolo"] or "",
                    row["grado"] or "",
                    "Sì" if row["is_autista"] else "No",
                    "Sì" if row["is_vigile"] else "No",
                    row["weekly_cap"] if row["weekly_cap"] is not None else DEFAULT_WEEKLY_CAP,
                    row["phone"] or "",
                    row["email"] or "",
                ),
            )

        if (
            self.selected_person_id is not None
            and str(self.selected_person_id) in self.people_tree.get_children()
        ):
            self.people_tree.selection_set(str(self.selected_person_id))
        else:
            self.reset_person_form()

    # ---------------------- Tab: Coppie & Vincoli ---------------------- #
    def _build_pairs_tab(self, container: ttk.Frame) -> None:
        container.columnconfigure(0, weight=1)
        container.rowconfigure(0, weight=1)
        container.rowconfigure(1, weight=1)

        forbidden_group = ttk.LabelFrame(container, text="Coppie vietate (vigili)")
        forbidden_group.grid(row=0, column=0, sticky="nsew", padx=4, pady=4)
        forbidden_group.columnconfigure(0, weight=1)
        forbidden_group.rowconfigure(0, weight=1)

        self.forbidden_tree = ttk.Treeview(
            forbidden_group,
            columns=("vigile1", "vigile2", "tipo"),
            show="headings",
            selectmode="browse",
        )
        self.forbidden_tree.heading("vigile1", text="Vigile 1")
        self.forbidden_tree.heading("vigile2", text="Vigile 2")
        self.forbidden_tree.heading("tipo", text="Vincolo")
        for col in ("vigile1", "vigile2", "tipo"):
            self.forbidden_tree.column(col, anchor="center", width=150)
        self.forbidden_tree.grid(row=0, column=0, sticky="nsew", padx=(0, 4), pady=(0, 4))
        scroll_f = ttk.Scrollbar(forbidden_group, orient="vertical", command=self.forbidden_tree.yview)
        self.forbidden_tree.configure(yscrollcommand=scroll_f.set)
        scroll_f.grid(row=0, column=1, sticky="ns")

        form = ttk.Frame(forbidden_group)
        form.grid(row=1, column=0, columnspan=2, sticky="ew", pady=4)
        form.columnconfigure(1, weight=1)
        form.columnconfigure(3, weight=1)

        ttk.Label(form, text="Vigile 1").grid(row=0, column=0, padx=4, pady=2, sticky="w")
        self.forbidden_vigile1 = tk.StringVar()
        self.forbidden_vigile1_combo = ttk.Combobox(form, textvariable=self.forbidden_vigile1, state="readonly")
        self.forbidden_vigile1_combo.grid(row=0, column=1, padx=4, pady=2, sticky="ew")

        ttk.Label(form, text="Vigile 2").grid(row=0, column=2, padx=4, pady=2, sticky="w")
        self.forbidden_vigile2 = tk.StringVar()
        self.forbidden_vigile2_combo = ttk.Combobox(form, textvariable=self.forbidden_vigile2, state="readonly")
        self.forbidden_vigile2_combo.grid(row=0, column=3, padx=4, pady=2, sticky="ew")

        self.forbidden_is_hard = tk.BooleanVar(value=True)
        ttk.Checkbutton(form, text="Vincolo duro (irrinunciabile)", variable=self.forbidden_is_hard).grid(
            row=0, column=4, padx=4, pady=2
        )

        ttk.Button(form, text="Aggiungi/Aggiorna", command=self.add_forbidden_pair).grid(row=0, column=5, padx=4, pady=2)
        ttk.Button(form, text="Elimina selezionata", command=self.delete_forbidden_pair).grid(row=0, column=6, padx=4, pady=2)

        preferred_group = ttk.LabelFrame(container, text="Coppie preferenziali (autista + vigile)")
        preferred_group.grid(row=1, column=0, sticky="nsew", padx=4, pady=4)
        preferred_group.columnconfigure(0, weight=1)
        preferred_group.rowconfigure(0, weight=1)

        self.preferred_tree = ttk.Treeview(
            preferred_group,
            columns=("autista", "vigile", "tipo"),
            show="headings",
            selectmode="browse",
        )
        for col, label in (("autista", "Autista"), ("vigile", "Vigile"), ("tipo", "Vincolo")):
            self.preferred_tree.heading(col, text=label)
            self.preferred_tree.column(col, anchor="center", width=150)
        self.preferred_tree.grid(row=0, column=0, sticky="nsew", padx=(0, 4), pady=(0, 4))
        scroll_p = ttk.Scrollbar(preferred_group, orient="vertical", command=self.preferred_tree.yview)
        self.preferred_tree.configure(yscrollcommand=scroll_p.set)
        scroll_p.grid(row=0, column=1, sticky="ns")

        form_p = ttk.Frame(preferred_group)
        form_p.grid(row=1, column=0, columnspan=2, sticky="ew", pady=4)
        form_p.columnconfigure(1, weight=1)
        form_p.columnconfigure(3, weight=1)

        ttk.Label(form_p, text="Autista").grid(row=0, column=0, padx=4, pady=2, sticky="w")
        self.preferred_autista = tk.StringVar()
        self.preferred_autista_combo = ttk.Combobox(form_p, textvariable=self.preferred_autista, state="readonly")
        self.preferred_autista_combo.grid(row=0, column=1, padx=4, pady=2, sticky="ew")

        ttk.Label(form_p, text="Vigile").grid(row=0, column=2, padx=4, pady=2, sticky="w")
        self.preferred_vigile = tk.StringVar()
        self.preferred_vigile_combo = ttk.Combobox(form_p, textvariable=self.preferred_vigile, state="readonly")
        self.preferred_vigile_combo.grid(row=0, column=3, padx=4, pady=2, sticky="ew")

        self.preferred_is_hard = tk.BooleanVar(value=False)
        ttk.Checkbutton(form_p, text="Vincolo duro (obbligatorio)", variable=self.preferred_is_hard).grid(
            row=0, column=4, padx=4, pady=2
        )

        ttk.Button(form_p, text="Aggiungi/Aggiorna", command=self.add_preferred_pair).grid(row=0, column=5, padx=4, pady=2)
        ttk.Button(form_p, text="Elimina selezionata", command=self.delete_preferred_pair).grid(row=0, column=6, padx=4, pady=2)

    def refresh_pairs_lists(self) -> None:
        vigili_sorted = sorted(self.vigili_names)
        autisti_sorted = sorted(self.autisti_names)
        self.forbidden_vigile1_combo["values"] = vigili_sorted
        self.forbidden_vigile2_combo["values"] = vigili_sorted
        self.preferred_autista_combo["values"] = autisti_sorted
        self.preferred_vigile_combo["values"] = vigili_sorted

        self.forbidden_tree.delete(*self.forbidden_tree.get_children())
        for pair_id, name1, name2, is_hard in self.db.list_forbidden_pairs_detailed():
            self.forbidden_tree.insert(
                "",
                "end",
                iid=str(pair_id),
                values=(name1, name2, "Duro" if is_hard else "Morbido"),
            )

        self.preferred_tree.delete(*self.preferred_tree.get_children())
        for pair_id, auto_name, vig_name, is_hard in self.db.list_preferred_pairs_detailed():
            self.preferred_tree.insert(
                "",
                "end",
                iid=str(pair_id),
                values=(auto_name, vig_name, "Duro" if is_hard else "Morbido"),
            )

    def add_forbidden_pair(self) -> None:
        nome1 = self.forbidden_vigile1.get()
        nome2 = self.forbidden_vigile2.get()
        if not nome1 or not nome2:
            messagebox.showerror("Errore", "Seleziona due vigili.")
            return
        if nome1 == nome2:
            messagebox.showerror("Errore", "Non è possibile creare un vincolo su una sola persona.")
            return
        id1 = self.name_to_id.get(nome1)
        id2 = self.name_to_id.get(nome2)
        if id1 is None or id2 is None:
            messagebox.showerror("Errore", "Vigile non trovato in anagrafica.")
            return
        try:
            self.db.set_forbidden_pair(id1, id2, is_hard=self.forbidden_is_hard.get())
            self.refresh_pairs_lists()
        except ValueError as exc:
            messagebox.showerror("Errore", str(exc))

    def delete_forbidden_pair(self) -> None:
        selection = self.forbidden_tree.selection()
        if not selection:
            messagebox.showwarning("Attenzione", "Seleziona una coppia vietata da eliminare.")
            return
        self.db.delete_forbidden_pair(int(selection[0]))
        self.refresh_pairs_lists()

    def add_preferred_pair(self) -> None:
        autista = self.preferred_autista.get()
        vigile = self.preferred_vigile.get()
        if not autista or not vigile:
            messagebox.showerror("Errore", "Seleziona sia autista sia vigile.")
            return
        aut_id = self.name_to_id.get(autista)
        vig_id = self.name_to_id.get(vigile)
        if aut_id is None or vig_id is None:
            messagebox.showerror("Errore", "Autista o vigile non trovati in anagrafica.")
            return
        try:
            self.db.set_preferred_pair(aut_id, vig_id, is_hard=self.preferred_is_hard.get())
            self.refresh_pairs_lists()
        except ValueError as exc:
            messagebox.showerror("Errore", str(exc))

    def delete_preferred_pair(self) -> None:
        selection = self.preferred_tree.selection()
        if not selection:
            messagebox.showwarning("Attenzione", "Seleziona una coppia preferenziale da eliminare.")
            return
        self.db.delete_preferred_pair(int(selection[0]))
        self.refresh_pairs_lists()

    # ---------------------- Tab: Ferie ---------------------- #
    def _build_vacations_tab(self, container: ttk.Frame) -> None:
        container.columnconfigure(0, weight=1)
        container.rowconfigure(0, weight=1)

        self.vacations_tree = ttk.Treeview(
            container,
            columns=("persona", "inizio", "fine", "nota"),
            show="headings",
            selectmode="browse",
        )
        for col, label in zip(("persona", "inizio", "fine", "nota"), ["Persona", "Dal", "Al", "Nota"]):
            self.vacations_tree.heading(col, text=label)
            width = 140 if col in ("inizio", "fine") else 220
            anchor = "center" if col in ("inizio", "fine") else "w"
            self.vacations_tree.column(col, width=width, anchor=anchor)
        self.vacations_tree.grid(row=0, column=0, sticky="nsew")

        scroll = ttk.Scrollbar(container, orient="vertical", command=self.vacations_tree.yview)
        self.vacations_tree.configure(yscrollcommand=scroll.set)
        scroll.grid(row=0, column=1, sticky="ns")

        form = ttk.Frame(container)
        form.grid(row=1, column=0, columnspan=2, sticky="ew", pady=8)
        for c in range(6):
            form.columnconfigure(c, weight=1 if c in (1, 3, 5) else 0)

        ttk.Label(form, text="Persona").grid(row=0, column=0, padx=4, pady=2, sticky="w")
        self.vacation_person = tk.StringVar()
        self.vacation_person_combo = ttk.Combobox(form, textvariable=self.vacation_person, state="readonly")
        self.vacation_person_combo.grid(row=0, column=1, padx=4, pady=2, sticky="ew")

        ttk.Label(form, text="Inizio (YYYY-MM-DD)").grid(row=0, column=2, padx=4, pady=2, sticky="w")
        self.vacation_start = tk.StringVar()
        ttk.Entry(form, textvariable=self.vacation_start).grid(row=0, column=3, padx=4, pady=2, sticky="ew")

        ttk.Label(form, text="Fine (YYYY-MM-DD)").grid(row=0, column=4, padx=4, pady=2, sticky="w")
        self.vacation_end = tk.StringVar()
        ttk.Entry(form, textvariable=self.vacation_end).grid(row=0, column=5, padx=4, pady=2, sticky="ew")

        ttk.Label(form, text="Nota").grid(row=1, column=0, padx=4, pady=2, sticky="w")
        self.vacation_note = tk.StringVar()
        ttk.Entry(form, textvariable=self.vacation_note).grid(row=1, column=1, columnspan=3, padx=4, pady=2, sticky="ew")

        ttk.Button(form, text="Aggiungi ferie", command=self.add_vacation).grid(row=1, column=4, padx=4, pady=2, sticky="ew")
        ttk.Button(form, text="Elimina selezionate", command=self.delete_vacation).grid(row=1, column=5, padx=4, pady=2, sticky="ew")

        helper = ttk.Label(
            container,
            text="Le ferie escludono automaticamente la persona dalle assegnazioni nelle date indicate.",
            foreground="#555555",
        )
        helper.grid(row=2, column=0, columnspan=2, sticky="w", padx=4, pady=(6, 0))

    def refresh_vacations(self) -> None:
        self.vacations_tree.delete(*self.vacations_tree.get_children())
        for row in self.db.list_vacations():
            self.vacations_tree.insert(
                "",
                "end",
                iid=str(row["id"]),
                values=(row["person_name"], row["start_date"], row["end_date"], row["note"] or ""),
            )
        all_names = sorted(self.name_to_id.keys())
        self.vacation_person_combo["values"] = all_names

    def add_vacation(self) -> None:
        nome = self.vacation_person.get()
        if not nome:
            messagebox.showerror("Errore", "Seleziona una persona.")
            return
        person_id = self.name_to_id.get(nome)
        if person_id is None:
            messagebox.showerror("Errore", "Persona non trovata.")
            return
        try:
            start = datetime.strptime(self.vacation_start.get().strip(), "%Y-%m-%d").date()
            end = datetime.strptime(self.vacation_end.get().strip(), "%Y-%m-%d").date()
        except ValueError:
            messagebox.showerror("Errore", "Date non valide. Usa il formato YYYY-MM-DD.")
            return
        nota = self.vacation_note.get().strip() or None
        try:
            self.db.add_vacation(person_id, start, end, nota)
            self.vacation_start.set("")
            self.vacation_end.set("")
            self.vacation_note.set("")
            self.refresh_vacations()
        except ValueError as exc:
            messagebox.showerror("Errore", str(exc))

    def delete_vacation(self) -> None:
        selection = self.vacations_tree.selection()
        if not selection:
            messagebox.showwarning("Attenzione", "Seleziona un periodo di ferie da eliminare.")
            return
        self.db.remove_vacation(int(selection[0]))
        self.refresh_vacations()

    # ---------------------- Tab: Impostazioni generali ---------------------- #
    def _build_settings_tab(self, container: ttk.Frame) -> None:
        for i in range(6):
            container.rowconfigure(i, weight=0)
        container.columnconfigure(1, weight=1)

        ttk.Label(container, text="Autista speciale (solo venerdì)").grid(row=0, column=0, padx=6, pady=6, sticky="w")
        self.setting_autista_varchi = tk.StringVar()
        self.setting_autista_varchi_combo = ttk.Combobox(container, textvariable=self.setting_autista_varchi, state="readonly")
        self.setting_autista_varchi_combo.grid(row=0, column=1, padx=6, pady=6, sticky="ew")

        ttk.Label(container, text="Autista con vincolo Pogliani").grid(row=1, column=0, padx=6, pady=6, sticky="w")
        self.setting_autista_pogliani = tk.StringVar()
        self.setting_autista_pogliani_combo = ttk.Combobox(container, textvariable=self.setting_autista_pogliani, state="readonly")
        self.setting_autista_pogliani_combo.grid(row=1, column=1, padx=6, pady=6, sticky="ew")

        ttk.Label(container, text="Vigile escluso in estate (Luglio/Agosto)").grid(row=2, column=0, padx=6, pady=6, sticky="w")
        self.setting_vigile_estate = tk.StringVar()
        self.setting_vigile_estate_combo = ttk.Combobox(container, textvariable=self.setting_vigile_estate, state="readonly")
        self.setting_vigile_estate_combo.grid(row=2, column=1, padx=6, pady=6, sticky="ew")

        ttk.Label(container, text="Senior minimi per squadra").grid(row=3, column=0, padx=6, pady=6, sticky="w")
        self.setting_min_esperti = tk.IntVar(value=DEFAULT_MIN_ESPERTI)
        ttk.Spinbox(container, from_=0, to=4, textvariable=self.setting_min_esperti, width=5).grid(row=3, column=1, padx=6, pady=6, sticky="w")

        self.setting_varchi_rule = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            container,
            text="Abilita regola speciale Varchi/Pogliani (default consigliato)",
            variable=self.setting_varchi_rule,
        ).grid(row=4, column=0, columnspan=2, padx=6, pady=6, sticky="w")

        giorni_frame = ttk.LabelFrame(container, text="Giorni della settimana da pianificare")
        giorni_frame.grid(row=5, column=0, columnspan=2, sticky="ew", padx=6, pady=10)
        self.setting_weekdays: Dict[int, tk.BooleanVar] = {}
        for idx, dow in enumerate(range(7)):
            var = tk.BooleanVar(value=dow in DEFAULT_ACTIVE_WEEKDAYS)
            self.setting_weekdays[dow] = var
            ttk.Checkbutton(giorni_frame, text=WEEKDAY_LABELS[dow], variable=var).grid(
                row=0, column=idx, padx=4, pady=4, sticky="w"
            )

        btn_frame = ttk.Frame(container)
        btn_frame.grid(row=6, column=0, columnspan=2, pady=12)
        ttk.Button(btn_frame, text="Salva impostazioni", command=self.save_settings).grid(row=0, column=0, padx=4)
        ttk.Button(btn_frame, text="Ricarica", command=self.refresh_settings_inputs).grid(row=0, column=1, padx=4)

        helper = ttk.Label(
            container,
            text=(
                "Suggerimenti:\n"
                "• Seleziona i giorni da includere nel piano (es. weekend, intera settimana, ecc.).\n"
                "• I vincoli duri non vengono mai violati; quelli morbidi possono essere ignorati solo in assenza di alternative."
            ),
            foreground="#555555",
            justify="left",
        )
        helper.grid(row=7, column=0, columnspan=2, sticky="w", padx=6, pady=(8, 0))

    def refresh_settings_inputs(self) -> None:
        values_autisti = [""] + sorted(self.autisti_names)
        values_vigili = [""] + sorted(self.vigili_names)
        self.setting_autista_varchi_combo["values"] = values_autisti
        self.setting_autista_pogliani_combo["values"] = values_autisti
        self.setting_vigile_estate_combo["values"] = values_vigili

        settings = self.db.all_settings()
        self.setting_autista_varchi.set(settings.get("autista_varchi", ""))
        self.setting_autista_pogliani.set(settings.get("autista_pogliani", ""))
        self.setting_vigile_estate.set(settings.get("vigile_escluso_estate", ""))
        try:
            self.setting_min_esperti.set(int(settings.get("min_esperti", DEFAULT_MIN_ESPERTI)))
        except (TypeError, ValueError):
            self.setting_min_esperti.set(DEFAULT_MIN_ESPERTI)
        self.setting_varchi_rule.set(settings.get("enable_varchi_rule", "1") != "0")

        attivi = settings.get(
            "active_weekdays",
            ",".join(str(x) for x in sorted(DEFAULT_ACTIVE_WEEKDAYS)),
        )
        selezionati = {int(token) for token in attivi.split(",") if token.strip().isdigit()}
        if not selezionati:
            selezionati = set(DEFAULT_ACTIVE_WEEKDAYS)
        for dow, var in self.setting_weekdays.items():
            var.set(dow in selezionati)

    def save_settings(self) -> None:
        autisti_validi = self.autisti_names
        vigili_validi = self.vigili_names

        def _normalize(value: str, valid: Set[str]) -> Optional[str]:
            value = value.strip()
            return value if value and value in valid else None

        self.db.set_setting("autista_varchi", _normalize(self.setting_autista_varchi.get(), autisti_validi))
        self.db.set_setting("autista_pogliani", _normalize(self.setting_autista_pogliani.get(), autisti_validi))
        self.db.set_setting("vigile_escluso_estate", _normalize(self.setting_vigile_estate.get(), vigili_validi))
        self.db.set_setting("min_esperti", str(max(0, min(4, self.setting_min_esperti.get()))))
        self.db.set_setting("enable_varchi_rule", "1" if self.setting_varchi_rule.get() else "0")

        weekday_string = ",".join(
            str(dow) for dow, var in self.setting_weekdays.items() if var.get()
        )
        if not weekday_string:
            weekday_string = ",".join(str(x) for x in sorted(DEFAULT_ACTIVE_WEEKDAYS))
        self.db.set_setting("active_weekdays", weekday_string)
        messagebox.showinfo("Impostazioni salvate", "Le impostazioni sono state aggiornate correttamente.")

    # ---------------------- Tab: Generazione turni ---------------------- #
    def _build_generation_tab(self, container: ttk.Frame) -> None:
        for c in range(4):
            container.columnconfigure(c, weight=1 if c in (1, 3) else 0)

        row = 0
        ttk.Label(container, text="Anno da pianificare").grid(row=row, column=0, padx=6, pady=6, sticky="w")
        self.gen_year = tk.IntVar(value=datetime.now().year)
        ttk.Spinbox(container, from_=datetime.now().year - 5, to=datetime.now().year + 10, textvariable=self.gen_year, width=8).grid(
            row=row, column=1, padx=6, pady=6, sticky="w"
        )

        ttk.Label(container, text="Seed (facoltativo)").grid(row=row, column=2, padx=6, pady=6, sticky="w")
        self.gen_seed = tk.StringVar()
        ttk.Entry(container, textvariable=self.gen_seed, width=12).grid(row=row, column=3, padx=6, pady=6, sticky="w")

        row += 1
        ttk.Label(container, text="Cartella output").grid(row=row, column=0, padx=6, pady=6, sticky="w")
        self.gen_output_dir = tk.StringVar(value=str((Path.cwd() / "output").resolve()))
        ttk.Entry(container, textvariable=self.gen_output_dir).grid(row=row, column=1, padx=6, pady=6, sticky="ew", columnspan=2)
        ttk.Button(container, text="Sfoglia…", command=self._choose_output_dir).grid(row=row, column=3, padx=6, pady=6, sticky="w")

        row += 1
        self.gen_import_from_text = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            container,
            text="Importa i file legacy nel database prima di generare",
            variable=self.gen_import_from_text,
            command=self._toggle_generation_mode,
        ).grid(row=row, column=0, columnspan=2, padx=6, pady=4, sticky="w")

        self.gen_use_legacy = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            container,
            text="Usa solo i file legacy (senza database)",
            variable=self.gen_use_legacy,
            command=self._toggle_generation_mode,
        ).grid(row=row, column=2, columnspan=2, padx=6, pady=4, sticky="w")

        row += 1
        ttk.Label(container, text="File autisti.txt").grid(row=row, column=0, padx=6, pady=4, sticky="w")
        self.gen_autisti_path = tk.StringVar(value=str((Path.cwd() / "autisti.txt").resolve()))
        self.gen_autisti_entry = ttk.Entry(container, textvariable=self.gen_autisti_path)
        self.gen_autisti_entry.grid(row=row, column=1, padx=6, pady=4, sticky="ew")
        self._browse_autisti = ttk.Button(
            container, text="Sfoglia…", command=lambda: self._choose_file(self.gen_autisti_path)
        )
        self._browse_autisti.grid(row=row, column=2, padx=6, pady=4, sticky="w")

        row += 1
        ttk.Label(container, text="File vigili.txt").grid(row=row, column=0, padx=6, pady=4, sticky="w")
        self.gen_vigili_path = tk.StringVar(value=str((Path.cwd() / "vigili.txt").resolve()))
        self.gen_vigili_entry = ttk.Entry(container, textvariable=self.gen_vigili_path)
        self.gen_vigili_entry.grid(row=row, column=1, padx=6, pady=4, sticky="ew")
        self._browse_vigili = ttk.Button(
            container, text="Sfoglia…", command=lambda: self._choose_file(self.gen_vigili_path)
        )
        self._browse_vigili.grid(row=row, column=2, padx=6, pady=4, sticky="w")

        row += 1
        ttk.Label(container, text="File vigili_senior.txt").grid(row=row, column=0, padx=6, pady=4, sticky="w")
        self.gen_vigili_senior_path = tk.StringVar(value=str((Path.cwd() / "vigili_senior.txt").resolve()))
        self.gen_vigili_senior_entry = ttk.Entry(container, textvariable=self.gen_vigili_senior_path)
        self.gen_vigili_senior_entry.grid(row=row, column=1, padx=6, pady=4, sticky="ew")
        self._browse_vigili_senior = ttk.Button(
            container, text="Sfoglia…", command=lambda: self._choose_file(self.gen_vigili_senior_path, must_exist=False)
        )
        self._browse_vigili_senior.grid(row=row, column=2, padx=6, pady=4, sticky="w")

        row += 1
        ttk.Label(container, text="Percorso database").grid(row=row, column=0, padx=6, pady=6, sticky="w")
        self.gen_db_path = tk.StringVar(value=str(self.db_path.resolve()))
        self.gen_db_entry = ttk.Entry(container, textvariable=self.gen_db_path)
        self.gen_db_entry.grid(row=row, column=1, padx=6, pady=6, sticky="ew")
        self._browse_db = ttk.Button(container, text="Sfoglia…", command=self._choose_db_path)
        self._browse_db.grid(row=row, column=2, padx=6, pady=6, sticky="w")

        row += 1
        btn_frame = ttk.Frame(container)
        btn_frame.grid(row=row, column=0, columnspan=4, pady=12, sticky="ew")
        btn_frame.columnconfigure(0, weight=1)
        btn_frame.columnconfigure(1, weight=0)
        btn_frame.columnconfigure(2, weight=0)

        self.generate_button = ttk.Button(btn_frame, text="Genera turni", command=self.run_generation)
        self.generate_button.grid(row=0, column=0, padx=6, sticky="w")
        ttk.Button(btn_frame, text="Apri cartella output", command=self._open_output_folder).grid(row=0, column=1, padx=6)
        ttk.Button(btn_frame, text="Pulisci log", command=self._clear_generation_output).grid(row=0, column=2, padx=6)

        row += 1
        ttk.Label(container, text="Log generazione").grid(row=row, column=0, padx=6, pady=(0, 4), sticky="w")
        row += 1
        self.generate_output = tk.Text(container, height=12, wrap="word")
        self.generate_output.grid(row=row, column=0, columnspan=4, padx=6, pady=(0, 6), sticky="nsew")
        scroll = ttk.Scrollbar(container, orient="vertical", command=self.generate_output.yview)
        scroll.grid(row=row, column=4, sticky="ns", pady=(0, 6))
        self.generate_output.configure(yscrollcommand=scroll.set)
        container.rowconfigure(row, weight=1)

        self._toggle_generation_mode()

    def _choose_output_dir(self) -> None:
        directory = filedialog.askdirectory(title="Seleziona cartella di output", initialdir=self.gen_output_dir.get())
        if directory:
            self.gen_output_dir.set(directory)

    def _choose_file(self, var: tk.StringVar, must_exist: bool = True) -> None:
        file_path = filedialog.askopenfilename(
            title="Seleziona file",
            initialdir=Path(var.get()).parent if var.get() else Path.cwd(),
            filetypes=[("File di testo", "*.txt"), ("Tutti i file", "*.*")],
        )
        if file_path:
            var.set(file_path)
        elif must_exist and not Path(var.get()).exists():
            messagebox.showwarning("Attenzione", "Il file selezionato deve esistere.")

    def _choose_db_path(self) -> None:
        file_path = filedialog.askopenfilename(
            title="Seleziona database SQLite",
            initialdir=Path(self.gen_db_path.get()).parent if self.gen_db_path.get() else Path.cwd(),
            filetypes=[("Database SQLite", "*.db"), ("Tutti i file", "*.*")],
        )
        if file_path:
            self.gen_db_path.set(file_path)

    def _toggle_generation_mode(self) -> None:
        legacy = self.gen_use_legacy.get()
        need_files = legacy or self.gen_import_from_text.get()
        state_files = "normal" if need_files else "disabled"
        for entry in (
            self.gen_autisti_entry,
            self.gen_vigili_entry,
            self.gen_vigili_senior_entry,
        ):
            entry.configure(state=state_files)
        for button in (
            self._browse_autisti,
            self._browse_vigili,
            self._browse_vigili_senior,
        ):
            button.configure(state=state_files)

        db_state = "disabled" if legacy else "normal"
        self.gen_db_entry.configure(state=db_state)
        self._browse_db.configure(state=db_state)

    def _clear_generation_output(self) -> None:
        self.generate_output.configure(state="normal")
        self.generate_output.delete("1.0", "end")
        self.generate_output.configure(state="normal")

    def _append_generation_output(self, text: str) -> None:
        self.generate_output.configure(state="normal")
        self.generate_output.insert("end", text)
        self.generate_output.see("end")
        self.generate_output.configure(state="normal")

    def run_generation(self) -> None:
        if self.generate_running:
            return
        year = self.gen_year.get()
        out_dir = Path(self.gen_output_dir.get()).expanduser()
        out_dir.mkdir(parents=True, exist_ok=True)

        cmd = [sys.executable, str(Path(__file__).with_name("turnivvf_fixed.py")), "--year", str(year), "--out", str(out_dir)]

        seed = self.gen_seed.get().strip()
        if seed:
            if not seed.isdigit():
                messagebox.showerror("Errore", "Il seed deve essere un numero intero positivo.")
                return
            cmd.extend(["--seed", seed])

        autisti_file = Path(self.gen_autisti_path.get()).expanduser()
        vigili_file = Path(self.gen_vigili_path.get()).expanduser()
        vigili_senior_file = Path(self.gen_vigili_senior_path.get()).expanduser()

        if self.gen_use_legacy.get():
            for file_path, label in ((autisti_file, "autisti.txt"), (vigili_file, "vigili.txt")):
                if not file_path.exists():
                    messagebox.showerror("Errore", f"File {label} non trovato: {file_path}")
                    return
            cmd.extend(
                [
                    "--skip-db",
                    "--autisti",
                    str(autisti_file),
                    "--vigili",
                    str(vigili_file),
                    "--vigili-senior",
                    str(vigili_senior_file),
                ]
            )
        else:
            db_path = Path(self.gen_db_path.get()).expanduser()
            cmd.extend(["--db", str(db_path)])
            if self.gen_import_from_text.get():
                for file_path, label in ((autisti_file, "autisti.txt"), (vigili_file, "vigili.txt")):
                    if not file_path.exists():
                        messagebox.showerror("Errore", f"File {label} non trovato: {file_path}")
                        return
                cmd.append("--import-from-text")
                cmd.extend(
                    [
                        "--autisti",
                        str(autisti_file),
                        "--vigili",
                        str(vigili_file),
                        "--vigili-senior",
                        str(vigili_senior_file),
                    ]
                )

        self._clear_generation_output()
        self._append_generation_output(f"$ {' '.join(cmd)}\n")
        self.generate_button.configure(state="disabled")
        self.generate_running = True
        threading.Thread(target=self._run_generation_thread, args=(cmd,), daemon=True).start()
        self.after(100, self._poll_generation_queue)

    def _run_generation_thread(self, cmd: List[str]) -> None:
        try:
            process = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
            )
            assert process.stdout is not None
            for line in process.stdout:
                self.generate_queue.put(line)
            returncode = process.wait()
            if returncode == 0:
                self.generate_queue.put("Generazione completata con successo.\n")
            else:
                self.generate_queue.put(f"Errore durante la generazione (codice {returncode}).\n")
        except FileNotFoundError:
            self.generate_queue.put("Impossibile avviare Python. Controlla l'installazione.\n")
        except Exception as exc:
            self.generate_queue.put(f"Errore inatteso: {exc}\n")
        finally:
            self.generate_queue.put("__END__")

    def _poll_generation_queue(self) -> None:
        try:
            while True:
                message = self.generate_queue.get_nowait()
                if message == "__END__":
                    self.generate_running = False
                    self.generate_button.configure(state="normal")
                    return
                self._append_generation_output(message)
        except queue.Empty:
            if self.generate_running:
                self.after(100, self._poll_generation_queue)

    def _open_output_folder(self) -> None:
        path = Path(self.gen_output_dir.get())
        if not path.exists():
            messagebox.showinfo("Informazione", "La cartella di output non esiste ancora.")
            return
        try:
            if sys.platform.startswith("win"):
                subprocess.Popen(["explorer", str(path)])
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(path)])
            else:
                subprocess.Popen(["xdg-open", str(path)])
        except Exception as exc:
            messagebox.showerror("Errore", f"Impossibile aprire la cartella: {exc}")

    # ---------------------- Utilità globali ---------------------- #
    def refresh_all(self) -> None:
        self.refresh_people_list()
        self.refresh_pairs_lists()
        self.refresh_vacations()
        self.refresh_settings_inputs()

    def on_close(self) -> None:
        try:
            self.db.close()
        finally:
            self.destroy()


def main() -> None:
    parser = argparse.ArgumentParser(description="Interfaccia grafica per la gestione del database VVF Scheduler.")
    parser.add_argument("--db", type=Path, default=Path("vvf_data.db"), help="Percorso del database SQLite (default: vvf_data.db)")
    args = parser.parse_args()
    app = SchedulerManagerApp(args.db)
    app.mainloop()


if __name__ == "__main__":
    main()
