"""Motore di generazione turni VVF."""

from __future__ import annotations

import itertools
import logging
import random
from dataclasses import dataclass, field
from datetime import date, timedelta
from typing import Dict, Iterable, List, Optional, Set, Tuple

from database import (
    ProgramConfig,
    Vacation,
    DEFAULT_ACTIVE_WEEKDAYS,
    DEFAULT_WEEKLY_CAP,
)

from .constants import LIV_JUNIOR, LIV_SENIOR, SUMMER_EXCLUDED_MONTHS
from .rules import RuleMode, merge_with_defaults


def date_attive_anno(
    anno: int,
    weekdays: Iterable[int],
    months: Optional[Iterable[int]] = None,
) -> List[date]:
    """Ritorna tutte le date dell'anno appartenenti ai giorni/mesi indicati."""
    giorni_attivi = {int(g) for g in weekdays if 0 <= int(g) <= 6}
    if not giorni_attivi:
        giorni_attivi = set(DEFAULT_ACTIVE_WEEKDAYS)
    if months:
        mesi_attivi = {int(m) for m in months if 1 <= int(m) <= 12}
        if not mesi_attivi:
            mesi_attivi = set(range(1, 13))
    else:
        mesi_attivi = set(range(1, 13))
    giorno = date(anno, 1, 1)
    fine = date(anno, 12, 31)
    risultati: List[date] = []
    one_day = timedelta(days=1)
    while giorno <= fine:
        if giorno.weekday() in giorni_attivi and giorno.month in mesi_attivi:
            risultati.append(giorno)
        giorno += one_day
    return risultati


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


@dataclass
class Assegnazione:
    giorno: date
    autista: Optional[str]
    vigili: Tuple[Optional[str], Optional[str], Optional[str], Optional[str]]


class Scheduler:
    """Motore di generazione dei turni che applica tutti i vincoli configurati."""

    def __init__(
        self,
        anno: int,
        config: ProgramConfig,
        months: Optional[Iterable[int]] = None,
    ):
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

        self.rules = merge_with_defaults(config.generation_rules)
        self.rule_min_senior = self.rules["min_senior"]
        self.rule_weekly_cap = self.rules["weekly_cap"]
        self.rule_summer = self.rules["summer_exclusion"]
        self.rule_varchi = self.rules["varchi_rotation"]

        self.enable_varchi_rule = bool(config.enable_varchi_rule) and self.rule_varchi.mode != RuleMode.OFF
        self.autista_varchi = config.autista_varchi
        self.autista_pogliani = config.autista_pogliani
        self.vigile_escluso_estate = config.vigile_escluso_estate if self.rule_summer.mode != RuleMode.OFF else None
        min_rule_value = self.rule_min_senior.value if self.rule_min_senior.value is not None else config.min_esperti
        self.min_esperti = max(0, min_rule_value)
        self.active_weekdays = set(config.active_weekdays or DEFAULT_ACTIVE_WEEKDAYS)
        self.active_months: Set[int] = (
            {int(m) for m in months if 1 <= int(m) <= 12} if months else set(range(1, 13))
        )
        if not self.active_months:
            self.active_months = set(range(1, 13))
        self.date = date_attive_anno(anno, self.active_weekdays, self.active_months)
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
        self.logger = logging.getLogger("vvf.scheduler")

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

    def _week_key(self, giorno: date) -> Tuple[int, int]:
        iso = giorno.isocalendar()
        return iso.year, iso.week

    def _limite_settimanale(self, nome: str) -> int:
        return self.weekly_cap.get(nome, self.default_weekly_cap)

    def _ha_raggiunto_limite(self, conteggi: Conteggi, nome: str, giorno: date) -> bool:
        if self.rule_weekly_cap.mode == RuleMode.OFF:
            return False
        raggiunto = self._limite_raggiunto_raw(conteggi, nome, giorno)
        if raggiunto and self.rule_weekly_cap.mode == RuleMode.SOFT:
            return False
        return raggiunto

    def _limite_raggiunto_raw(self, conteggi: Conteggi, nome: str, giorno: date) -> bool:
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

    def _costruisci_per_data(
        self,
        giorno: date,
        assegnazioni: Dict[date, Assegnazione],
    ) -> Assegnazione:
        sabato_autista = self._trova_autista_settimanale(assegnazioni, giorno, 5)
        assegnazione, autista_reale = self._costruisci_per_data_internal(
            giorno,
            assegnazioni,
            sabato_autista=sabato_autista,
            apply_varchi=self.enable_varchi_rule,
        )
        if (
            self.enable_varchi_rule
            and self.rule_varchi.mode == RuleMode.SOFT
            and self._turno_incompleto(assegnazione)
        ):
            # Se la regola speciale impedisce di coprire il turno, riprovo senza imporla
            self._log(
                giorno,
                "AUTISTA",
                "Deroga regola Varchi/Pogliani: ricompongo il turno senza il vincolo speciale.",
            )
            assegnazione, autista_reale = self._costruisci_per_data_internal(
                giorno,
                assegnazioni,
                sabato_autista=sabato_autista,
                apply_varchi=False,
            )
        self.autisti_reali[giorno] = autista_reale
        return assegnazione

    def _costruisci_per_data_internal(
        self,
        giorno: date,
        assegnazioni: Dict[date, Assegnazione],
        *,
        sabato_autista: Optional[str],
        apply_varchi: bool,
    ) -> Tuple[Assegnazione, Optional[str]]:
        dow = giorno.weekday()

        esclusioni_autista: Set[str] = set()
        if apply_varchi and self.autista_varchi and dow != 4:
            esclusioni_autista.add(self.autista_varchi)
        if (
            apply_varchi
            and dow == 4
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
        display_autista = autista
        include_varchi = (
            apply_varchi
            and self.varchi_is_senior
            and dow == 4
            and self.autista_varchi
            and (self.autista_pogliani is None or sabato_autista != self.autista_pogliani)
        )
        if (
            apply_varchi
            and self.varchi_is_senior
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

        vigili_target = self._numero_vigili_previsti(dow)
        vigili_base = max(0, vigili_target - (1 if include_varchi else 0))

        esclusioni_vigili: Set[str] = set()
        if autista:
            esclusioni_vigili.add(autista)
        if apply_varchi and self.autista_varchi:
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
            return Assegnazione(giorno=giorno, autista=autista, vigili=tuple(vigili_list)), autista

        squadra_list = list(squadra)
        if include_varchi:
            squadra_list = self._aggiungi_varchi_venerdi(giorno, squadra_list)

        while len(squadra_list) < 4:
            squadra_list.append(None)

        assegnazione = Assegnazione(
            giorno=giorno,
            autista=display_autista,
            vigili=tuple(squadra_list[:4]),
        )
        return assegnazione, autista

    @staticmethod
    def _turno_incompleto(assegnazione: Assegnazione) -> bool:
        if assegnazione.autista is None:
            return True
        return any(v is None for v in assegnazione.vigili)

    def _scegli_autista(self, giorno: date, esclusioni: Set[str]) -> Optional[str]:
        mese = giorno.month
        dow = giorno.weekday()
        week_key = self._week_key(giorno)

        candidati_info: List[Tuple[str, bool]] = []
        for nome in self.autisti:
            if nome in esclusioni:
                continue
            if self._in_ferie(nome, giorno):
                continue
            limit_raw = self._limite_raggiunto_raw(self.cont_aut, nome, giorno)
            if self.rule_weekly_cap.mode == RuleMode.HARD and limit_raw:
                continue
            candidati_info.append((nome, limit_raw))

        if not candidati_info:
            self._log(giorno, "AUTISTA", "Nessun autista disponibile rispettando vincoli e limiti settimanali.")
            return None

        limit_map = {nome: limit for nome, limit in candidati_info}
        candidati = [nome for nome, _ in candidati_info]
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
                1 if limit_map[nome] else 0,
                self.cont_aut.tot_settimana(nome, week_key),
                self.cont_aut.tot_mese(nome, mese),
                self.cont_aut.tot_annuale(nome),
                self.cont_aut.tot_giorno_anno(nome, dow),
                1 if self.cont_aut.ultimo_dow(nome) == dow else 0,
                random.random(),
            )
        )
        scelto = pool[0]
        if self.rule_weekly_cap.mode == RuleMode.SOFT and limit_map[scelto]:
            self._log(
                giorno,
                "AUTISTA",
                f"Deroga limite settimanale: assegno {scelto} oltre il proprio limite.",
            )
        self.cont_aut.aggiungi(scelto, giorno)
        return scelto

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

        candidati_base: List[str] = []
        fallback_candidates: List[str] = []
        limit_map: Dict[str, bool] = {}
        summer_map: Dict[str, bool] = {}
        for nome in self.vigili:
            # Raccolgo i candidati distinguendo quelli che violerebbero vincoli soft
            if nome in esclusioni:
                continue
            if self._in_ferie(nome, giorno):
                continue
            summer_block = (
                self.vigile_escluso_estate
                and nome == self.vigile_escluso_estate
                and mese in SUMMER_EXCLUDED_MONTHS
                and self.rule_summer.mode != RuleMode.OFF
            )
            limit_raw = self._limite_raggiunto_raw(self.cont_vig, nome, giorno)
            if self.rule_summer.mode == RuleMode.HARD and summer_block:
                continue
            if self.rule_weekly_cap.mode == RuleMode.HARD and limit_raw:
                continue
            limit_map[nome] = limit_raw
            summer_map[nome] = summer_block
            if summer_block or limit_raw:
                fallback_candidates.append(nome)
            else:
                candidati_base.append(nome)

        disponibili = list(candidati_base)
        if len(disponibili) < n_vigili and fallback_candidates:
            reason_parts: List[str] = []
            if any(limit_map[n] for n in fallback_candidates):
                reason_parts.append("limite settimanale")
            if any(summer_map[n] for n in fallback_candidates):
                reason_parts.append("regola estiva")
            descr = " e ".join(reason_parts) if reason_parts else "vincoli soft"
            self._log(
                giorno,
                "VIGILI",
                f"Deroga {descr}: includo {', '.join(fallback_candidates)} fra i candidati.",
            )
            disponibili.extend(fallback_candidates)

        if len(disponibili) < n_vigili:
            self._log(
                giorno,
                "VIGILI",
                f"Candidati insufficienti ({len(disponibili)}/{n_vigili}) dopo aver applicato ferie, limiti e vincoli.",
            )
            return None

        ci_sono_senior = any(
            self.esperienza_vigili.get(nome, LIV_JUNIOR) == LIV_SENIOR for nome in disponibili
        )

        obbligatori = self._preferenze_obbligatorie(autista_corrente)
        disponibili_obbligatori = [nome for nome in obbligatori if nome in disponibili]
        mancanti = [nome for nome in obbligatori if nome not in disponibili]
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
        residui = [nome for nome in disponibili if nome not in disponibili_obbligatori]
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

            if any(frozenset(pair) in self.forbidden_hard for pair in itertools.combinations(team, 2)):
                continue

            senior_count = sum(
                1 for nome in team if self.esperienza_vigili.get(nome, LIV_JUNIOR) == LIV_SENIOR
            )
            if self.min_esperti > 0:
                if not ci_sono_senior and not deroga_senior_loggata:
                    self._log(
                        giorno,
                        "VIGILI",
                        "Deroga esperienza: nessun SENIOR disponibile fra i candidati di oggi.",
                    )
                    deroga_senior_loggata = True
                if ci_sono_senior and senior_count < self.min_esperti:
                    if self.rule_min_senior.mode == RuleMode.SOFT:
                        if not deroga_senior_loggata:
                            self._log(
                                giorno,
                                "VIGILI",
                                f"Deroga esperienza: squadra con {senior_count} SENIOR (<{self.min_esperti}).",
                            )
                            deroga_senior_loggata = True
                    else:
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

            punteggio = (
                violazioni_soft,
                violazioni_mese_dow,
                squadra_nuova,
                carico_settimanale,
                carico_mensile,
                carico_annuale,
                carico_giorno,
                ripetizioni_recenti,
                preferenze_soft,
                random.random(),
            )
            soluzioni.append((punteggio, team, {"violazioni_soft": violazioni_soft}))

        if not soluzioni:
            self._log(
                giorno,
                "VIGILI",
                "Nessuna combinazione di squadra soddisfa i vincoli duri + soft gestibili.",
            )
            return None

        soluzioni.sort(key=lambda item: item[0])
        team = soluzioni[0][1]
        info = soluzioni[0][2]
        if info.get("violazioni_soft"):
            self._log(
                giorno,
                "VIGILI",
                f"Deroga: utilizzo squadra con {info['violazioni_soft']} coppia/e sconsigliate.",
            )

        if frozenset(team) in self.squadre_visti:
            self._log(
                giorno,
                "VIGILI",
                f"Squadra già vista {team}: la riutilizzo per mancanza di alternative migliori.",
            )

        if self.rule_weekly_cap.mode == RuleMode.SOFT:
            for nome in team:
                if limit_map.get(nome):
                    self._log(
                        giorno,
                        "VIGILI",
                        f"Deroga limite settimanale: associo {nome} oltre il proprio limite.",
                    )
        if self.rule_summer.mode == RuleMode.SOFT:
            for nome in team:
                if summer_map.get(nome):
                    self._log(
                        giorno,
                        "VIGILI",
                        f"Deroga regola estiva: includo {nome} nonostante il vincolo.",
                    )

        for nome in team:
            self.cont_vig.aggiungi(nome, giorno)
        self.squadre_visti.add(frozenset(team))

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

    def _log(self, giorno: date, categoria: str, messaggio: str) -> None:
        stamp = giorno.strftime("%Y-%m-%d (%a)")
        self.logger.info("[%s] %s", categoria, messaggio)
        self.log.append(f"[{stamp}] [{categoria}] {messaggio}")
