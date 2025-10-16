#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
VVF Weekend Scheduler – Generatore turni annuali con supporto a database e GUI.

Il programma legge l'anagrafica del personale (autisti e vigili), rispetta vincoli
configurabili (coppie vietate/prefenziali, limiti settimanali, ferie, ecc.) e
produce tre output per l'anno richiesto:
  • turni_<ANNO>.xlsx  → riepilogo mensile + report dei conteggi;
  • turni_<ANNO>.ics   → calendario elettronico (timezone Europe/Rome);
  • log_<ANNO>.txt     → note operative e deroghe applicate.

L'uso consigliato è via database SQLite gestito dalla GUI (`vvf_gui.py`), ma rimane
disponibile una modalità "legacy" basata sui file di testo storici.
"""

from __future__ import annotations

import argparse
import itertools
import random
import re
import sys
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta, timezone
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Set, Tuple

from database import (
    ConstraintRule,
    Database,
    PersonProfile,
    PreferredRule,
    ProgramConfig,
    Vacation,
    DEFAULT_ACTIVE_WEEKDAYS,
    DEFAULT_AUTISTA_POGLIANI,
    DEFAULT_AUTISTA_VARCHI,
    DEFAULT_FORBIDDEN_PAIRS,
    DEFAULT_MIN_ESPERTI,
    DEFAULT_PREFERRED_PAIRS,
    DEFAULT_VIGILE_ESCLUSO_ESTATE,
    DEFAULT_WEEKLY_CAP,
    ROLE_AUTISTA,
    ROLE_AUTISTA_VIGILE,
    ROLE_VIGILE,
)

try:
    import pandas as pd
except ImportError:
    print("Questo script richiede pandas. Installa con: pip install pandas openpyxl", file=sys.stderr)
    raise

# ------------------------------
# Costanti di base e helper
# ------------------------------
TZID = "Europe/Rome"
NOME_GIORNO = {
    0: "Lunedì",
    1: "Martedì",
    2: "Mercoledì",
    3: "Giovedì",
    4: "Venerdì",
    5: "Sabato",
    6: "Domenica",
}

MESI_IT = {
    1: "Gennaio",
    2: "Febbraio",
    3: "Marzo",
    4: "Aprile",
    5: "Maggio",
    6: "Giugno",
    7: "Luglio",
    8: "Agosto",
    9: "Settembre",
    10: "Ottobre",
    11: "Novembre",
    12: "Dicembre",
}

LIV_SENIOR = "SENIOR"
LIV_JUNIOR = "JUNIOR"
SUMMER_EXCLUDED_MONTHS = (7, 8)


def _norm_name(name: str) -> str:
    """Normalizza un nome per confronti robusti (trim + casefold)."""
    return re.sub(r"\s+", " ", name).strip().casefold()


def _match_person_identifier(
    value: Optional[str],
    roster: Iterable[str],
    profiles: Dict[str, PersonProfile],
) -> Optional[str]:
    """Risolvo un identificativo (nome o cognome) nel roster."""
    if not value:
        return None
    target_norm = _norm_name(value)

    for name in roster:
        if _norm_name(name) == target_norm:
            return name

    for name, profile in profiles.items():
        display = profile.display_name
        if display and _norm_name(display) == target_norm:
            return name
        cognome = profile.cognome or ""
        if cognome and _norm_name(cognome) == target_norm:
            return name
    return None


def carica_nomi(path: Path) -> List[str]:
    """Legacy: carica una lista di nomi dal file testuale (una persona per riga)."""
    if not path.exists():
        raise FileNotFoundError(f"File mancante: {path}")
    nomi_raw: List[str] = []
    with path.open("r", encoding="utf-8") as fh:
        for raw in fh:
            nome = raw.strip()
            if nome and not nome.startswith("#"):
                nomi_raw.append(nome)
    # Rimuovo eventuali duplicati mantenendo l'ordine
    visti: Set[str] = set()
    nomi: List[str] = []
    for nome in nomi_raw:
        if nome not in visti:
            visti.add(nome)
            nomi.append(nome)
    return nomi


def build_program_config_from_files(
    file_autisti: Path, file_vigili: Path, file_vigili_senior: Path
) -> ProgramConfig:
    """Crea una ProgramConfig partendo dai file storici (senza passare dal DB)."""
    autisti = carica_nomi(file_autisti)
    vigili_junior = carica_nomi(file_vigili)
    vigili_senior = (
        carica_nomi(file_vigili_senior) if file_vigili_senior.exists() else []
    )

    def _split(full: str) -> Tuple[str, str]:
        chunks = full.split()
        if len(chunks) >= 2:
            return chunks[0], " ".join(chunks[1:])
        return full, ""

    persone: Dict[str, PersonProfile] = {}

    def _ensure_person(
        nome_visualizzato: str,
        ruolo: str,
        grado: str,
        *,
        autista: bool,
        vigile: bool,
        livello: str,
    ):
        first, last = _split(nome_visualizzato)
        profilo = persone.get(nome_visualizzato)
        if profilo is None:
            profilo = PersonProfile(
                id=-(len(persone) + 1),  # id fittizio (non usato in modalità legacy)
                nome=first,
                cognome=last,
                telefono="",
                email="",
                ruolo=ruolo,
                grado=grado,
                is_autista=autista,
                is_vigile=vigile,
                livello=livello,
                weekly_cap=DEFAULT_WEEKLY_CAP,
            )
            persone[nome_visualizzato] = profilo
        else:
            if autista:
                profilo.is_autista = True
            if vigile:
                profilo.is_vigile = True
                profilo.livello = livello
            if grado:
                profilo.grado = grado
            # Aggiorno il ruolo mantenendo AUTISTA+VIGILE quando necessario
            if ruolo == ROLE_AUTISTA_VIGILE:
                profilo.ruolo = ruolo
            elif ruolo == ROLE_AUTISTA and profilo.ruolo != ROLE_AUTISTA_VIGILE:
                profilo.ruolo = ROLE_AUTISTA
            elif ruolo == ROLE_VIGILE and profilo.ruolo != ROLE_AUTISTA_VIGILE:
                profilo.ruolo = ROLE_VIGILE

    for nome in autisti:
        _ensure_person(nome, ROLE_AUTISTA, "", autista=True, vigile=False, livello=LIV_JUNIOR)
    for nome in vigili_junior:
        ruolo = ROLE_VIGILE if nome not in persone or not persone[nome].is_autista else ROLE_AUTISTA_VIGILE
        _ensure_person(nome, ruolo, LIV_JUNIOR, autista=False, vigile=True, livello=LIV_JUNIOR)
    for nome in vigili_senior:
        ruolo = ROLE_VIGILE if nome not in persone or not persone[nome].is_autista else ROLE_AUTISTA_VIGILE
        _ensure_person(nome, ruolo, LIV_SENIOR, autista=False, vigile=True, livello=LIV_SENIOR)

    elenco_autisti = [nome for nome, profilo in persone.items() if profilo.is_autista]
    elenco_vigili = [nome for nome, profilo in persone.items() if profilo.is_vigile]
    esperienza = {
        nome: profilo.livello for nome, profilo in persone.items() if profilo.is_vigile
    }
    weekly_caps = {nome: profilo.weekly_cap for nome, profilo in persone.items()}

    coppie_vietate = [
        ConstraintRule(primo=a, secondo=b, is_hard=True) for a, b in DEFAULT_FORBIDDEN_PAIRS
    ]
    coppie_preferite = [
        PreferredRule(autista=a, vigile=b, is_hard=False) for a, b in DEFAULT_PREFERRED_PAIRS
    ]

    autista_varchi = _match_person_identifier(DEFAULT_AUTISTA_VARCHI, elenco_autisti, persone)
    autista_pogliani = _match_person_identifier(DEFAULT_AUTISTA_POGLIANI, elenco_autisti, persone)
    vigile_estivo = _match_person_identifier(DEFAULT_VIGILE_ESCLUSO_ESTATE, elenco_vigili, persone)

    return ProgramConfig(
        autisti=elenco_autisti,
        vigili=elenco_vigili,
        esperienza_vigili=esperienza,
        weekly_cap=weekly_caps,
        coppie_vietate=coppie_vietate,
        coppie_preferite=coppie_preferite,
        autista_varchi=autista_varchi,
        autista_pogliani=autista_pogliani,
        vigile_escluso_estate=vigile_estivo,
        min_esperti=DEFAULT_MIN_ESPERTI,
        ferie={},
        active_weekdays=set(DEFAULT_ACTIVE_WEEKDAYS),
        people=persone,
        enable_varchi_rule=True,
    )


def date_attive_anno(anno: int, weekdays: Iterable[int]) -> List[date]:
    """Ritorna tutte le date dell'anno appartenenti ai giorni della settimana indicati."""
    giorni_attivi = {int(g) for g in weekdays if 0 <= int(g) <= 6}
    if not giorni_attivi:
        giorni_attivi = set(DEFAULT_ACTIVE_WEEKDAYS)
    giorno = date(anno, 1, 1)
    fine = date(anno, 12, 31)
    risultati: List[date] = []
    one_day = timedelta(days=1)
    while giorno <= fine:
        if giorno.weekday() in giorni_attivi:
            risultati.append(giorno)
        giorno += one_day
    return risultati


# ------------------------------
# Strutture dati conteggi
# ------------------------------
@dataclass
class Conteggi:
    """Tiene traccia delle statistiche di assegnazione per autisti e vigili."""

    annuale: Dict[str, int] = field(default_factory=dict)
    per_mese: Dict[str, Dict[int, int]] = field(default_factory=dict)
    per_mese_giorno: Dict[str, Dict[int, Dict[int, int]]] = field(default_factory=dict)
    per_giorno_anno: Dict[str, Dict[int, int]] = field(default_factory=dict)
    per_settimana: Dict[str, Dict[Tuple[int, int], int]] = field(default_factory=dict)
    ultimo_giorno: Dict[str, Optional[int]] = field(default_factory=dict)

    def assicura_persona(self, nome: str) -> None:
        if nome not in self.annuale:
            self.annuale[nome] = 0
        if nome not in self.per_mese:
            self.per_mese[nome] = {mese: 0 for mese in range(1, 13)}
        if nome not in self.per_mese_giorno:
            self.per_mese_giorno[nome] = {
                mese: {dow: 0 for dow in range(7)} for mese in range(1, 13)
            }
        if nome not in self.per_giorno_anno:
            self.per_giorno_anno[nome] = {dow: 0 for dow in range(7)}
        if nome not in self.per_settimana:
            self.per_settimana[nome] = {}
        if nome not in self.ultimo_giorno:
            self.ultimo_giorno[nome] = None

    def aggiungi(self, nome: str, giorno: date) -> None:
        self.assicura_persona(nome)
        mese = giorno.month
        dow = giorno.weekday()
        week_key = (giorno.isocalendar().year, giorno.isocalendar().week)

        self.annuale[nome] += 1
        self.per_mese[nome][mese] += 1
        self.per_mese_giorno[nome][mese][dow] += 1
        self.per_giorno_anno[nome][dow] += 1
        self.per_settimana[nome][week_key] = self.per_settimana[nome].get(week_key, 0) + 1
        self.ultimo_giorno[nome] = dow

    def tot_mese(self, nome: str, mese: int) -> int:
        self.assicura_persona(nome)
        return self.per_mese[nome][mese]

    def tot_annuale(self, nome: str) -> int:
        self.assicura_persona(nome)
        return self.annuale[nome]

    def tot_mese_giorno(self, nome: str, mese: int, dow: int) -> int:
        self.assicura_persona(nome)
        return self.per_mese_giorno[nome][mese][dow]

    def tot_giorno_anno(self, nome: str, dow: int) -> int:
        self.assicura_persona(nome)
        return self.per_giorno_anno[nome][dow]

    def tot_settimana(self, nome: str, week_key: Tuple[int, int]) -> int:
        self.assicura_persona(nome)
        return self.per_settimana[nome].get(week_key, 0)

    def ultimo_dow(self, nome: str) -> Optional[int]:
        self.assicura_persona(nome)
        return self.ultimo_giorno[nome]


# ------------------------------
# Modello di assegnazione
# ------------------------------
@dataclass
class Assegnazione:
    giorno: date
    autista: Optional[str]
    vigili: Tuple[Optional[str], Optional[str], Optional[str], Optional[str]]


# ------------------------------
# Scheduler
# ------------------------------
class Scheduler:
    """Motore di generazione dei turni che applica tutti i vincoli configurati."""

    def __init__(self, anno: int, config: ProgramConfig):
        self.anno = anno
        self.config = config

        self.autisti: List[str] = sorted(config.autisti)
        self.vigili: List[str] = sorted(config.vigili)
        self.esperienza_vigili = {
            nome: config.esperienza_vigili.get(nome, LIV_JUNIOR) for nome in self.vigili
        }

        self.forbidden_hard: Set[frozenset] = {
            frozenset(rule.as_sorted_tuple()) for rule in config.coppie_vietate if rule.is_hard
        }
        self.forbidden_soft: Set[frozenset] = {
            frozenset(rule.as_sorted_tuple()) for rule in config.coppie_vietate if not rule.is_hard
        }
        self.preferenze_hard: Dict[str, Set[str]] = {}
        self.preferenze_soft: Dict[str, Set[str]] = {}
        for rule in config.coppie_preferite:
            target = self.preferenze_hard if rule.is_hard else self.preferenze_soft
            target.setdefault(rule.autista, set()).add(rule.vigile)

        self.weekly_cap = {nome: max(0, cap) for nome, cap in config.weekly_cap.items()}
        self.default_weekly_cap = max(0, DEFAULT_WEEKLY_CAP)

        self.enable_varchi_rule = bool(config.enable_varchi_rule)
        self.autista_varchi = config.autista_varchi
        self.autista_pogliani = config.autista_pogliani
        self.vigile_escluso_estate = config.vigile_escluso_estate
        self.min_esperti = max(0, config.min_esperti)
        self.active_weekdays = set(config.active_weekdays or DEFAULT_ACTIVE_WEEKDAYS)
        self.date = date_attive_anno(anno, self.active_weekdays)
        self.ferie: Dict[str, List[Vacation]] = {
            nome: list(vac) for nome, vac in config.ferie.items()
        }

        self.cont_aut = Conteggi()
        self.cont_vig = Conteggi()
        for nome in self.autisti:
            self.cont_aut.assicura_persona(nome)
        for nome in self.vigili:
            self.cont_vig.assicura_persona(nome)

        self.squadre_visti: Set[frozenset] = set()
        self.log: List[str] = []
        self.autisti_reali: Dict[date, Optional[str]] = {}

        if not self.enable_varchi_rule:
            # Se la regola è disattivata, non forzo alcun nome speciale
            self.autista_varchi = None
            self.autista_pogliani = None

        self.varchi_is_senior = (
            self.enable_varchi_rule
            and self.autista_varchi is not None
            and self.autista_varchi in self.vigili
            and self.esperienza_vigili.get(self.autista_varchi, LIV_JUNIOR) == LIV_SENIOR
        )

    # --------------------------
    # Helper interni
    # --------------------------
    def _week_key(self, giorno: date) -> Tuple[int, int]:
        iso = giorno.isocalendar()
        return iso.year, iso.week

    def _limite_settimanale(self, nome: str) -> int:
        return self.weekly_cap.get(nome, self.default_weekly_cap)

    def _ha_raggiunto_limite(self, conteggi: Conteggi, nome: str, giorno: date) -> bool:
        cap = self._limite_settimanale(nome)
        if cap <= 0:
            return False
        return conteggi.tot_settimana(nome, self._week_key(giorno)) >= cap

    def _in_ferie(self, nome: str, giorno: date) -> bool:
        for vac in self.ferie.get(nome, []):
            if vac.start <= giorno <= vac.end:
                return True
        return False

    def _preferenze_obbligatorie(self, autista: Optional[str]) -> Set[str]:
        if not autista:
            return set()
        return set(self.preferenze_hard.get(autista, set()))

    def _preferenze_soft(self, autista: Optional[str]) -> Set[str]:
        if not autista:
            return set()
        return set(self.preferenze_soft.get(autista, set()))

    def _numero_vigili_previsti(self, dow: int) -> int:
        return 4

    def _ordine_giorni(self, giorni: Dict[int, date]) -> List[int]:
        ordine: List[int] = []
        for dow in (5, 4, 6):
            if dow in giorni:
                ordine.append(dow)
        for dow in sorted(giorni):
            if dow not in (4, 5, 6):
                ordine.append(dow)
        return ordine

    def _trova_autista_settimanale(
        self, assegnazioni: Dict[date, Assegnazione], giorno: date, target_dow: int
    ) -> Optional[str]:
        week_key = self._week_key(giorno)
        for data, assegnazione in assegnazioni.items():
            if self._week_key(data) == week_key and data.weekday() == target_dow:
                return self.autisti_reali.get(data, assegnazione.autista)
        return None

    # --------------------------
    # Costruzione globale
    # --------------------------
    def costruisci(self) -> List[Assegnazione]:
        assegnazioni: Dict[date, Assegnazione] = {}

        per_settimana: Dict[Tuple[int, int], List[date]] = {}
        for giorno in self.date:
            per_settimana.setdefault(self._week_key(giorno), []).append(giorno)
        for giorni in per_settimana.values():
            giorni.sort()

        for _, giorni in sorted(per_settimana.items()):
            giorni_per_dow = {d.weekday(): d for d in giorni}
            for dow in self._ordine_giorni(giorni_per_dow):
                giorno = giorni_per_dow[dow]
                assegnazioni[giorno] = self._costruisci_per_data(giorno, assegnazioni)

        return [assegnazioni[d] for d in sorted(assegnazioni)]

    # --------------------------
    # Assegnazione giornaliera
    # --------------------------
    def _costruisci_per_data(
        self,
        giorno: date,
        assegnazioni: Dict[date, Assegnazione],
    ) -> Assegnazione:
        dow = giorno.weekday()
        sabato_autista = self._trova_autista_settimanale(assegnazioni, giorno, 5)

        # Gestione autisti
        esclusioni_autista: Set[str] = set()
        if self.autista_varchi and dow != 4:
            esclusioni_autista.add(self.autista_varchi)
        if (
            dow == 4
            and self.autista_varchi
            and self.autista_pogliani
            and sabato_autista == self.autista_pogliani
        ):
            esclusioni_autista.add(self.autista_varchi)
            self._log(
                giorno,
                "AUTISTA",
                f"Regola: sabato guida {self.autista_pogliani} ⇒ venerdì escludo {self.autista_varchi}.",
            )

        autista = self._scegli_autista(giorno, esclusioni_autista)
        self.autisti_reali[giorno] = autista
        display_autista = autista
        if (
            self.varchi_is_senior
            and dow == 5
            and autista == self.autista_pogliani
            and self.autista_varchi
        ):
            display_autista = self.autista_varchi
            self._log(
                giorno,
                "AUTISTA",
                f"Regola speciale: visualizzo {self.autista_varchi} al posto di {self.autista_pogliani} (conteggio attribuito a {self.autista_pogliani}).",
            )

        # Squadra vigili
        vigili_target = self._numero_vigili_previsti(dow)
        include_varchi = (
            dow == 4
            and self.varchi_is_senior
            and self.autista_varchi
            and (self.autista_pogliani is None or sabato_autista != self.autista_pogliani)
        )
        vigili_base = max(0, vigili_target - (1 if include_varchi else 0))

        esclusioni_vigili: Set[str] = set()
        if autista:
            esclusioni_vigili.add(autista)
        if self.autista_varchi and (dow != 4 or not include_varchi):
            esclusioni_vigili.add(self.autista_varchi)

        squadra = self._scegli_squadra_vigili(
            giorno,
            vigili_base,
            autista_corrente=autista,
            esclusioni=esclusioni_vigili,
        )

        if squadra is None:
            self._log(giorno, "VIGILI", "Turno scoperto: impossibile comporre una squadra valida.")
            vigili_list: List[Optional[str]] = [None, None, None, None]
            return Assegnazione(giorno=giorno, autista=autista, vigili=tuple(vigili_list))

        squadra_list = list(squadra)
        if include_varchi:
            squadra_list = self._aggiungi_varchi_venerdi(giorno, squadra_list)

        while len(squadra_list) < 4:
            squadra_list.append(None)

        return Assegnazione(
            giorno=giorno,
            autista=display_autista,
            vigili=tuple(squadra_list[:4]),
        )

    # --------------------------
    # Scelta autista
    # --------------------------
    def _scegli_autista(self, giorno: date, esclusioni: Set[str]) -> Optional[str]:
        mese = giorno.month
        dow = giorno.weekday()
        week_key = self._week_key(giorno)

        candidati: List[str] = []
        for nome in self.autisti:
            if nome in esclusioni:
                continue
            if self._in_ferie(nome, giorno):
                continue
            if self._ha_raggiunto_limite(self.cont_aut, nome, giorno):
                continue
            candidati.append(nome)

        if not candidati:
            self._log(giorno, "AUTISTA", "Nessun autista disponibile rispettando vincoli e limiti settimanali.")
            return None

        preferiti = [
            nome for nome in candidati if self.cont_aut.tot_mese_giorno(nome, mese, dow) < 1
        ]
        if not preferiti:
            pool = candidati
            self._log(
                giorno,
                "AUTISTA",
                "Deroga: rilasso il vincolo un-turno-per mese/giorno sugli autisti per coprire il servizio.",
            )
        else:
            pool = preferiti

        pool.sort(
            key=lambda nome: (
                self.cont_aut.tot_settimana(nome, week_key),
                self.cont_aut.tot_mese(nome, mese),
                self.cont_aut.tot_annuale(nome),
                self.cont_aut.tot_giorno_anno(nome, dow),
                1 if self.cont_aut.ultimo_dow(nome) == dow else 0,
                random.random(),
            )
        )
        scelto = pool[0]
        self.cont_aut.aggiungi(scelto, giorno)
        return scelto

    # --------------------------
    # Scelta vigili
    # --------------------------
    def _scegli_squadra_vigili(
        self,
        giorno: date,
        n_vigili: int,
        *,
        autista_corrente: Optional[str],
        esclusioni: Set[str],
    ) -> Optional[Tuple[str, ...]]:
        if n_vigili <= 0:
            return tuple()

        mese = giorno.month
        dow = giorno.weekday()
        week_key = self._week_key(giorno)

        base: List[str] = []
        for nome in self.vigili:
            if nome in esclusioni:
                continue
            if (
                self.vigile_escluso_estate
                and nome == self.vigile_escluso_estate
                and mese in SUMMER_EXCLUDED_MONTHS
            ):
                continue
            if self._in_ferie(nome, giorno):
                continue
            if self._ha_raggiunto_limite(self.cont_vig, nome, giorno):
                continue
            base.append(nome)

        if len(base) < n_vigili:
            self._log(
                giorno,
                "VIGILI",
                f"Candidati insufficienti ({len(base)}/{n_vigili}) dopo aver applicato ferie, limiti e vincoli.",
            )
            return None

        ci_sono_senior = any(
            self.esperienza_vigili.get(nome, LIV_JUNIOR) == LIV_SENIOR for nome in base
        )

        obbligatori = self._preferenze_obbligatorie(autista_corrente)
        disponibili_obbligatori = [nome for nome in obbligatori if nome in base]
        mancanti = [nome for nome in obbligatori if nome not in base]
        for nome in mancanti:
            self._log(
                giorno,
                "VIGILI",
                f"Vincolo duro autista-vigile non rispettato (manca {nome}). Proseguo scegliendo la migliore alternativa.",
            )

        if len(disponibili_obbligatori) > n_vigili:
            self._log(
                giorno,
                "VIGILI",
                "Vincoli duri autista-vigile eccedono la dimensione squadra: limito al numero di slot disponibili.",
            )
            disponibili_obbligatori = disponibili_obbligatori[:n_vigili]

        slot_rimanenti = max(0, n_vigili - len(disponibili_obbligatori))
        residui = [nome for nome in base if nome not in disponibili_obbligatori]
        if slot_rimanenti > len(residui):
            self._log(
                giorno,
                "VIGILI",
                f"Impossibile completare la squadra ({slot_rimanenti} posti da coprire, {len(residui)} candidati idonei).",
            )
            return None

        combinazioni = (
            itertools.combinations(residui, slot_rimanenti)
            if slot_rimanenti > 0
            else [tuple()]
        )

        soluzioni: List[Tuple[Tuple[float, ...], Tuple[str, ...], Dict[str, int]]] = []
        deroga_senior_loggata = False

        for extra in combinazioni:
            team = tuple(disponibili_obbligatori + list(extra))
            team_set = frozenset(team)

            # Vincoli duri tra vigili
            if any(frozenset(pair) in self.forbidden_hard for pair in itertools.combinations(team, 2)):
                continue

            # Esperienza
            if not self._team_ok_by_experience(team, ci_sono_senior):
                if not ci_sono_senior and not deroga_senior_loggata:
                    self._log(
                        giorno,
                        "VIGILI",
                        "Deroga esperienza: nessun SENIOR disponibile fra i candidati di oggi.",
                    )
                    deroga_senior_loggata = True
                if ci_sono_senior:
                    continue

            violazioni_soft = sum(
                1 for pair in itertools.combinations(team, 2) if frozenset(pair) in self.forbidden_soft
            )
            violazioni_mese_dow = sum(
                1 for nome in team if self.cont_vig.tot_mese_giorno(nome, mese, dow) >= 1
            )
            squadra_nuova = 0 if team_set not in self.squadre_visti else 1
            carico_settimanale = sum(self.cont_vig.tot_settimana(nome, week_key) for nome in team)
            carico_mensile = sum(self.cont_vig.tot_mese(nome, mese) for nome in team)
            carico_annuale = sum(self.cont_vig.tot_annuale(nome) for nome in team)
            carico_giorno = sum(self.cont_vig.tot_giorno_anno(nome, dow) for nome in team)
            ripetizioni_recenti = sum(1 for nome in team if self.cont_vig.ultimo_dow(nome) == dow)
            preferenze_soft = -sum(
                1 for nome in team if nome in self._preferenze_soft(autista_corrente)
            )

            score = (
                violazioni_soft,
                violazioni_mese_dow,
                squadra_nuova,
                carico_settimanale,
                carico_mensile,
                carico_annuale,
                carico_giorno,
                preferenze_soft,
                ripetizioni_recenti,
                random.random(),
            )
            metriche = {
                "violazioni_soft": violazioni_soft,
                "violazioni_mese_dow": violazioni_mese_dow,
                "squadra_nuova": squadra_nuova == 0,
            }
            soluzioni.append((score, team, metriche))

        if not soluzioni:
            self._log(
                giorno,
                "VIGILI",
                "Nessuna squadra soddisfa i vincoli dopo le deroghe concesse.",
            )
            return None

        soluzioni.sort(key=lambda item: item[0])
        _, team, metriche = soluzioni[0]
        team_set = frozenset(team)

        if metriche["violazioni_soft"] > 0:
            self._log(
                giorno,
                "VIGILI",
                f"Deroga morbida: accetto {metriche['violazioni_soft']} coppia/e vietata/e (soft).",
            )
        if metriche["violazioni_mese_dow"] > 0:
            self._log(
                giorno,
                "VIGILI",
                f"Deroga: {metriche['violazioni_mese_dow']} vigili superano il limite mensile per questo giorno.",
            )
        if not metriche["squadra_nuova"]:
            self._log(
                giorno,
                "VIGILI",
                f"Squadra già vista {team}: la riutilizzo per mancanza di alternative migliori.",
            )

        for nome in team:
            self.cont_vig.aggiungi(nome, giorno)
        self.squadre_visti.add(team_set)

        return team

    def _aggiungi_varchi_venerdi(self, giorno: date, squadra: List[Optional[str]]) -> List[Optional[str]]:
        if not self.autista_varchi:
            return squadra
        if self.autista_varchi in squadra:
            return squadra
        if self._in_ferie(self.autista_varchi, giorno):
            self._log(
                giorno,
                "VIGILI",
                f"{self.autista_varchi} è in ferie: venerdì senza vigile speciale.",
            )
            return squadra
        if self._ha_raggiunto_limite(self.cont_vig, self.autista_varchi, giorno):
            self._log(
                giorno,
                "VIGILI",
                f"{self.autista_varchi} ha già raggiunto il limite settimanale: niente quarto SENIOR speciale.",
            )
            return squadra

        squadra.append(self.autista_varchi)
        self.cont_vig.aggiungi(self.autista_varchi, giorno)
        self.squadre_visti.add(frozenset(x for x in squadra if x))
        self._log(
            giorno,
            "VIGILI",
            f"Venerdì speciale: aggiungo {self.autista_varchi} come quarto vigile SENIOR.",
        )
        return squadra

    def _team_ok_by_experience(
        self, team: Tuple[str, ...], ci_sono_senior: bool
    ) -> bool:
        if self.min_esperti <= 0:
            return True
        if not ci_sono_senior:
            return True
        senior_presenti = sum(
            1 for nome in team if self.esperienza_vigili.get(nome, LIV_JUNIOR) == LIV_SENIOR
        )
        return senior_presenti >= self.min_esperti

    def _log(self, giorno: date, categoria: str, messaggio: str) -> None:
        stamp = giorno.strftime("%Y-%m-%d (%a)")
        self.log.append(f"[{stamp}] [{categoria}] {messaggio}")


# ------------------------------
# Excel / ICS
# ------------------------------
VTIMEZONE_EUROPE_ROME = """BEGIN:VTIMEZONE
TZID:Europe/Rome
X-LIC-LOCATION:Europe/Rome
BEGIN:DAYLIGHT
TZOFFSETFROM:+0100
TZOFFSETTO:+0200
TZNAME:CEST
DTSTART:19700329T020000
RRULE:FREQ=YEARLY;BYMONTH=3;BYDAY=-1SU
END:DAYLIGHT
BEGIN:STANDARD
TZOFFSETFROM:+0200
TZOFFSETTO:+0100
TZNAME:CET
DTSTART:19701025T030000
RRULE:FREQ=YEARLY;BYMONTH=10;BYDAY=-1SU
END:STANDARD
END:VTIMEZONE"""


def scrivi_excel(
    assegnazioni: List[Assegnazione],
    autisti: List[str],
    vigili: List[str],
    cont_aut: Conteggi,
    cont_vig: Conteggi,
    anno: int,
    out_path: Path,
) -> None:
    """Esporta l'esito dei turni in formato Excel (uno sheet per mese + report)."""
    per_mese: Dict[int, List[Assegnazione]] = {mese: [] for mese in range(1, 13)}
    for assegnazione in assegnazioni:
        per_mese[assegnazione.giorno.month].append(assegnazione)

    def _build_report_table(nomi: List[str], cont: Conteggi) -> pd.DataFrame:
        colonne = ["Nome", "Turni totali"]
        for mese in range(1, 13):
            colonne.extend(
                [
                    MESI_IT[mese],
                    f"{MESI_IT[mese]} Lun",
                    f"{MESI_IT[mese]} Mar",
                    f"{MESI_IT[mese]} Mer",
                    f"{MESI_IT[mese]} Gio",
                    f"{MESI_IT[mese]} Ven",
                    f"{MESI_IT[mese]} Sab",
                    f"{MESI_IT[mese]} Dom",
                ]
            )
        colonne.extend(["Lunedì", "Martedì", "Mercoledì", "Giovedì", "Venerdì", "Sabato", "Domenica"])

        righe = []
        for nome in nomi:
            tot_annuale = cont.tot_annuale(nome)
            valori: List[int] = []
            for mese in range(1, 13):
                tot_mese = cont.tot_mese(nome, mese)
                valori.append(tot_mese)
                for dow in range(7):
                    valori.append(cont.per_mese_giorno[nome][mese][dow])
            valori.extend(cont.per_giorno_anno[nome][dow] for dow in range(7))
            righe.append([nome, tot_annuale] + valori)
        return pd.DataFrame(righe, columns=colonne)

    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        for mese in range(1, 13):
            righe = []
            for assegnazione in sorted(per_mese[mese], key=lambda a: a.giorno):
                dow = assegnazione.giorno.weekday()
                righe.append(
                    {
                        "Data": assegnazione.giorno.strftime("%Y-%m-%d"),
                        "Giorno": NOME_GIORNO.get(dow, str(dow)),
                        "Autista": assegnazione.autista or "",
                        "Vigile1": assegnazione.vigili[0] or "",
                        "Vigile2": assegnazione.vigili[1] or "",
                        "Vigile3": assegnazione.vigili[2] or "",
                        "Vigile4": assegnazione.vigili[3] or "",
                    }
                )
            df = pd.DataFrame(
                righe, columns=["Data", "Giorno", "Autista", "Vigile1", "Vigile2", "Vigile3", "Vigile4"]
            )
            nome_foglio = MESI_IT[mese]
            df.to_excel(writer, sheet_name=nome_foglio, index=False)

        report_vig = _build_report_table(vigili, cont_vig)
        report_aut = _build_report_table(autisti, cont_aut)
        report_vig.to_excel(writer, sheet_name="Report", index=False, startrow=1)
        offset = len(report_vig) + 4
        report_aut.to_excel(writer, sheet_name="Report", index=False, startrow=offset)


def scrivi_ics(assegnazioni: List[Assegnazione], anno: int, out_path: Path) -> None:
    """Crea un file ICS con gli eventi per autisti e vigili."""
    righe: List[str] = [
        "BEGIN:VCALENDAR",
        "PRODID:-//VVF Scheduler//IT",
        "VERSION:2.0",
        "CALSCALE:GREGORIAN",
        "METHOD:PUBLISH",
        f"X-WR-CALNAME:Turni VVF {anno}",
        f"X-WR-TIMEZONE:{TZID}",
        VTIMEZONE_EUROPE_ROME,
    ]

    def _fmt_dt_locale(dt: datetime) -> str:
        return dt.strftime("%Y%m%dT%H%M%S")

    def _aggiungi_evento(nome: str, giorno: date, ora_inizio: int) -> None:
        from uuid import uuid4

        start = datetime(giorno.year, giorno.month, giorno.day, ora_inizio, 0, 0)
        end = start + timedelta(hours=1)
        uid = f"{uuid4()}@vvf-scheduler"
        righe.extend(
            [
            "BEGIN:VEVENT",
            f"UID:{uid}",
            f"DTSTAMP:{datetime.now(timezone.utc).strftime('%Y%m%dT%H%M%SZ')}",
            f"DTSTART;TZID={TZID}:{_fmt_dt_locale(start)}",
            f"DTEND;TZID={TZID}:{_fmt_dt_locale(end)}",
            f"SUMMARY:{nome}",
            "END:VEVENT",
            ]
        )

    for assegnazione in assegnazioni:
        if assegnazione.autista:
            _aggiungi_evento(assegnazione.autista, assegnazione.giorno, 11)
        for indice, nome in enumerate(assegnazione.vigili):
            if nome:
                _aggiungi_evento(nome, assegnazione.giorno, 12 + indice)

    righe.append("END:VCALENDAR")
    out_path.write_text("\n".join(righe), encoding="utf-8")


# ------------------------------
# Entry point
# ------------------------------
def esegui(
    anno: int,
    config: ProgramConfig,
    out_dir: Path,
    seed: Optional[int] = None,
) -> Tuple[Path, Path, Path]:
    if seed is not None:
        random.seed(seed)

    if not config.autisti:
        raise RuntimeError("Serve almeno un autista per generare il piano turni.")
    if len(config.vigili) < 1:
        raise RuntimeError("Serve almeno un vigile per generare il piano turni.")

    scheduler = Scheduler(anno, config)
    assegnazioni = scheduler.costruisci()

    out_dir.mkdir(parents=True, exist_ok=True)
    xlsx_path = out_dir / f"turni_{anno}.xlsx"
    ics_path = out_dir / f"turni_{anno}.ics"
    log_path = out_dir / f"log_{anno}.txt"

    scrivi_excel(
        assegnazioni=assegnazioni,
        autisti=config.autisti,
        vigili=config.vigili,
        cont_aut=scheduler.cont_aut,
        cont_vig=scheduler.cont_vig,
        anno=anno,
        out_path=xlsx_path,
    )
    scrivi_ics(assegnazioni, anno, ics_path)

    senior_count = sum(1 for livello in config.esperienza_vigili.values() if livello == LIV_SENIOR)
    header = [
        f"VVF Weekend Scheduler – anno {anno}",
        f"Autisti: {len(config.autisti)} ({', '.join(config.autisti)})",
        f"Vigili : {len(config.vigili)} ({', '.join(config.vigili)})",
        f"Vigili senior configurati: {senior_count}",
        f"Giorni pianificati: {len(config.active_weekdays)} ({', '.join(NOME_GIORNO[d] for d in sorted(config.active_weekdays))})",
    ]
    log_path.write_text(
        "\n".join(header + ["", "Registro decisioni/deroghe:"] + scheduler.log),
        encoding="utf-8",
    )

    return xlsx_path, ics_path, log_path


def main() -> None:
    parser = argparse.ArgumentParser(
        description="VVF Weekend Scheduler – turni weekend → Excel + ICS + Log (IT)"
    )
    parser.add_argument("--year", type=int, default=datetime.now().year, help="Anno di riferimento (default: anno corrente)")
    parser.add_argument("--db", type=Path, default=Path("vvf_data.db"), help="Percorso del database SQLite (default: vvf_data.db)")
    parser.add_argument("--import-from-text", action="store_true", help="Importa i file legacy nel database prima di generare i turni")
    parser.add_argument("--skip-db", action="store_true", help="Usa esclusivamente i file legacy senza database")
    parser.add_argument("--autisti", type=Path, default=Path("autisti.txt"), help="File autisti.txt (per import/legacy)")
    parser.add_argument("--vigili", type=Path, default=Path("vigili.txt"), help="File vigili.txt (JUNIOR, per import/legacy)")
    parser.add_argument("--vigili-senior", type=Path, default=Path("vigili_senior.txt"), help="File vigili_senior.txt (SENIOR, per import/legacy)")
    parser.add_argument("--out", type=Path, default=Path("output"), help="Cartella di output")
    parser.add_argument("--seed", type=int, default=None, help="Seed RNG per risultati ripetibili")
    args = parser.parse_args()

    if args.skip_db:
        config = build_program_config_from_files(args.autisti, args.vigili, args.vigili_senior)
    else:
        with Database(args.db) as db:
            if args.import_from_text:
                db.import_from_text_files(
                    autisti_path=args.autisti,
                    vigili_path=args.vigili,
                    vigili_senior_path=args.vigili_senior,
                    set_defaults=True,
                )
            config = db.load_program_config()
        if not config.autisti or not config.vigili:
            raise RuntimeError(
                "Il database non contiene autisti/vigili sufficienti. Popola i dati dalla GUI oppure usa --import-from-text."
            )

    xlsx_path, ics_path, log_path = esegui(args.year, config, args.out, args.seed)
    print("Operazione completata. File generati:")
    print(f"- {xlsx_path}")
    print(f"- {ics_path}")
    print(f"- {log_path}")


if __name__ == "__main__":
    main()
