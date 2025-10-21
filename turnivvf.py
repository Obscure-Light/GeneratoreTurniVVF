#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""VVF Weekend Scheduler – entrypoint CLI."""

from __future__ import annotations

import argparse
import logging
import random
import sys
from datetime import datetime
from pathlib import Path
from typing import List, Optional, Sequence

from database import Database
from vvf_scheduler.config import build_program_config_from_files
from vvf_scheduler.runner import esegui


def _setup_logging(verbose: bool = False) -> None:
    level = logging.DEBUG if verbose else logging.INFO
    logging.basicConfig(
        level=level,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )


def _parse_months(values: Optional[Sequence[int]], parser: argparse.ArgumentParser) -> Optional[List[int]]:
    if not values:
        return None
    mesi: List[int] = []
    for raw in values:
        if raw < 1 or raw > 12:
            parser.error("I mesi devono essere compresi tra 1 e 12.")
        if raw not in mesi:
            mesi.append(raw)
    return mesi


def main(argv: Optional[Sequence[str]] = None) -> None:
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
    parser.add_argument(
        "--months",
        type=int,
        nargs="+",
        help="Lista di mesi (1-12) da generare; se omesso vengono inclusi tutti i mesi.",
    )
    parser.add_argument("--verbose", action="store_true", help="Abilita log dettagliati su stdout")
    args = parser.parse_args(argv)

    _setup_logging(verbose=args.verbose)
    logging.debug("Argomenti CLI: %s", args)

    months = _parse_months(args.months, parser)
    if args.seed is not None:
        random.seed(args.seed)

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

    xlsx_path, ics_path, log_path, _ = esegui(args.year, config, args.out, months, args.seed)
    print("Operazione completata. File generati:")
    print(f"- {xlsx_path}")
    print(f"- {ics_path}")
    print(f"- {log_path}")


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:  # pragma: no cover - CLI fallback
        logging.getLogger("vvf").exception("Errore non gestito")
        print(f"Errore: {exc}", file=sys.stderr)
        sys.exit(1)
