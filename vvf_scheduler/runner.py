"""Funzioni di alto livello per eseguire il generatore di turni."""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Optional, Sequence, Tuple

from database import ProgramConfig

from .constants import MESI_IT, NOME_GIORNO, LIV_SENIOR
from .core import Scheduler
from .exports import scrivi_excel, scrivi_ics

logger = logging.getLogger(__name__)


def esegui(
    anno: int,
    config: ProgramConfig,
    out_dir: Path,
    months: Optional[Sequence[int]] = None,
    seed: Optional[int] = None,
) -> Tuple[Path, Path, Path, Scheduler]:
    scheduler = Scheduler(anno, config, months)
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
        selected_months=scheduler.active_months,
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
    if scheduler.active_months and len(scheduler.active_months) < 12:
        header.append(
            f"Mesi pianificati: {', '.join(MESI_IT[m] for m in sorted(scheduler.active_months))}"
        )
    log_path.write_text(
        "\n".join(header + ["", "Registro decisioni/deroghe:"] + scheduler.log),
        encoding="utf-8",
    )
    logger.info("Generazione completata: %s", out_dir)
    return xlsx_path, ics_path, log_path, scheduler
