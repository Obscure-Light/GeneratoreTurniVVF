"""Utility per costruire ProgramConfig da file legacy."""

from __future__ import annotations

import re
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Set, Tuple

from database import (
    ConstraintRule,
    PersonProfile,
    PreferredRule,
    ProgramConfig,
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

from .constants import LIV_JUNIOR, LIV_SENIOR
from .rules import build_default_rules


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
                id=-(len(persone) + 1),
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
            if ruolo == ROLE_AUTISTA_VIGILE:
                profilo.ruolo = ruolo
            elif ruolo == ROLE_AUTISTA and profilo.ruolo != ROLE_AUTISTA_VIGILE:
                profilo.ruolo = ROLE_AUTISTA
            elif ruolo == ROLE_VIGILE and profilo.ruolo != ROLE_AUTISTA_VIGILE:
                profilo.ruolo = ROLE_VIGILE

    for nome in autisti:
        _ensure_person(nome, ROLE_AUTISTA, "", autista=True, vigile=False, livello=LIV_JUNIOR)
    for nome in vigili_junior:
        _ensure_person(nome, ROLE_VIGILE, "", autista=False, vigile=True, livello=LIV_JUNIOR)
    for nome in vigili_senior:
        _ensure_person(nome, ROLE_VIGILE, "", autista=False, vigile=True, livello=LIV_SENIOR)

    elenco_autisti = sorted(persone)
    elenco_vigili = sorted(persone)
    esperienza = {nome: profilo.livello for nome, profilo in persone.items()}
    weekly_caps = {nome: profilo.weekly_cap for nome, profilo in persone.items()}

    coppie_vietate: List[ConstraintRule] = []
    for primo, secondo in DEFAULT_FORBIDDEN_PAIRS:
        ids = (_match_person_identifier(primo, elenco_vigili, persone), _match_person_identifier(secondo, elenco_vigili, persone))
        if all(ids):
            coppie_vietate.append(
                ConstraintRule(primo=ids[0], secondo=ids[1], is_hard=True)
            )

    coppie_preferite: List[PreferredRule] = []
    for autista_nome, vigile_nome in DEFAULT_PREFERRED_PAIRS:
        autista_id = _match_person_identifier(autista_nome, elenco_autisti, persone)
        vigile_id = _match_person_identifier(vigile_nome, elenco_vigili, persone)
        if autista_id and vigile_id:
            coppie_preferite.append(
                PreferredRule(autista=autista_id, vigile=vigile_id, is_hard=False)
            )

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
        generation_rules=build_default_rules(),
    )
