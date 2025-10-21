"""Funzioni di esportazione (Excel, ICS, log)."""

from __future__ import annotations

from datetime import date, datetime, timedelta, timezone
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Sequence

try:
    import pandas as pd
except ImportError as exc:  # pragma: no cover - dipendenza opzionale
    raise RuntimeError(
        "Questo script richiede pandas. Installa con: pip install pandas openpyxl"
    ) from exc

from .constants import MESI_IT, NOME_GIORNO, TZID
from .core import Assegnazione, Conteggi

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
    selected_months: Optional[Iterable[int]] = None,
) -> None:
    """Esporta l'esito dei turni in formato Excel (uno sheet per mese + report)."""
    per_mese: Dict[int, List[Assegnazione]] = {mese: [] for mese in range(1, 13)}
    for assegnazione in assegnazioni:
        per_mese[assegnazione.giorno.month].append(assegnazione)

    def _build_report_table(nomi: List[str], cont: Conteggi, mesi_rilevanti: Sequence[int]) -> pd.DataFrame:
        colonne = ["Nome", "Turni totali"]
        for mese in mesi_rilevanti:
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
            for mese in mesi_rilevanti:
                tot_mese = cont.tot_mese(nome, mese)
                valori.append(tot_mese)
                for dow in range(7):
                    valori.append(cont.per_mese_giorno[nome][mese][dow])
            valori.extend(cont.per_giorno_anno[nome][dow] for dow in range(7))
            righe.append([nome, tot_annuale] + valori)
        return pd.DataFrame(righe, columns=colonne)

    mesi_excel = (
        sorted({int(m) for m in selected_months if 1 <= int(m) <= 12})
        if selected_months
        else list(range(1, 13))
    )
    if not mesi_excel:
        mesi_excel = list(range(1, 13))

    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        for mese in mesi_excel:
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

        report_vig = _build_report_table(vigili, cont_vig, mesi_excel)
        report_aut = _build_report_table(autisti, cont_aut, mesi_excel)
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
