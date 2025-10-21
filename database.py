from __future__ import annotations

import re
import sqlite3
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Sequence, Set, Tuple

from vvf_scheduler.rules import (
    GenerationRuleConfig,
    RULE_DEFINITIONS,
    RuleMode,
    build_default_rules,
)


DEFAULT_FORBIDDEN_PAIRS: Set[Tuple[str, str]] = {
    ("Copellini", "Gallicchio"),
    ("Pila", "Garzaro"),
}
DEFAULT_PREFERRED_PAIRS: Set[Tuple[str, str]] = {
    ("Mascaretti", "Frangipane"),
}
DEFAULT_AUTISTA_VARCHI = "Varchi"
DEFAULT_AUTISTA_POGLIANI = "Pogliani"
DEFAULT_VIGILE_ESCLUSO_ESTATE = "Lodigiani"
DEFAULT_MIN_ESPERTI = 1
DEFAULT_ACTIVE_WEEKDAYS: Set[int] = {4, 5, 6}  # Venerdì, Sabato, Domenica
DEFAULT_WEEKLY_CAP = 1  # un turno per settimana salvo diversa indicazione

ROLE_AUTISTA = "AUTISTA"
ROLE_VIGILE = "VIGILE"
ROLE_AUTISTA_VIGILE = "AUTISTA+VIGILE"


def _normalize_whitespace(value: str) -> str:
    """Strip leading/trailing spaces and compress internal whitespace."""
    return re.sub(r"\s+", " ", value).strip()


def _parse_date(value: str) -> date:
    return datetime.strptime(value, "%Y-%m-%d").date()


def _format_date(value: date) -> str:
    return value.strftime("%Y-%m-%d")


@dataclass
class Vacation:
    start: date
    end: date
    note: Optional[str] = None


@dataclass
class ConstraintRule:
    primo: str
    secondo: str
    is_hard: bool = True

    def as_sorted_tuple(self) -> Tuple[str, str]:
        return tuple(sorted((self.primo, self.secondo)))


@dataclass
class PreferredRule:
    autista: str
    vigile: str
    is_hard: bool = False


@dataclass
class PersonProfile:
    id: int
    nome: str
    cognome: Optional[str]
    telefono: Optional[str]
    email: Optional[str]
    ruolo: Optional[str]
    grado: Optional[str]
    is_autista: bool
    is_vigile: bool
    livello: str
    weekly_cap: int

    @property
    def display_name(self) -> str:
        if self.nome and self.cognome:
            return f"{self.nome} {self.cognome}".strip()
        return self.nome or self.cognome or ""


def _match_person_identifier(
    value: Optional[str],
    roster: Iterable[str],
    profiles: Dict[str, PersonProfile],
) -> Optional[str]:
    """Risolvo un identificativo (nome o cognome) al nome completo presente nelle liste."""
    if not value:
        return None
    target_norm = _normalize_whitespace(value).casefold()

    # match diretto sul nome completo
    for name in roster:
        if _normalize_whitespace(name).casefold() == target_norm:
            return name

    # match su display name o cognome
    for name, profile in profiles.items():
        display_norm = _normalize_whitespace(profile.display_name).casefold()
        if display_norm == target_norm:
            return name
        cognome = (profile.cognome or "").strip()
        if cognome and _normalize_whitespace(cognome).casefold() == target_norm:
            return name

    return None


@dataclass
class ProgramConfig:
    autisti: List[str]
    vigili: List[str]
    esperienza_vigili: Dict[str, str]
    weekly_cap: Dict[str, int] = field(default_factory=dict)
    coppie_vietate: List[ConstraintRule] = field(default_factory=list)
    coppie_preferite: List[PreferredRule] = field(default_factory=list)
    autista_varchi: Optional[str] = None
    autista_pogliani: Optional[str] = None
    vigile_escluso_estate: Optional[str] = None
    min_esperti: int = 1
    ferie: Dict[str, List[Vacation]] = field(default_factory=dict)
    active_weekdays: Set[int] = field(default_factory=lambda: set(DEFAULT_ACTIVE_WEEKDAYS))
    people: Dict[str, PersonProfile] = field(default_factory=dict)
    enable_varchi_rule: bool = True
    generation_rules: Dict[str, GenerationRuleConfig] = field(default_factory=dict)


class Database:
    """SQLite helper to manage scheduler data."""

    def __init__(self, path: Path):
        self.path = Path(path)
        self.conn = sqlite3.connect(str(self.path))
        self.conn.row_factory = sqlite3.Row
        self.conn.execute("PRAGMA foreign_keys = ON;")
        self._ensure_schema()

    # ------------------------------------------------------------------ #
    # Schema
    # ------------------------------------------------------------------ #
    def _ensure_schema(self):
        self.conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS people (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL UNIQUE COLLATE NOCASE,
                is_autista INTEGER NOT NULL DEFAULT 0,
                is_vigile INTEGER NOT NULL DEFAULT 0,
                livello TEXT NOT NULL DEFAULT 'JUNIOR',
                note TEXT
            );

            CREATE TABLE IF NOT EXISTS settings (
                key TEXT PRIMARY KEY,
                value TEXT
            );

            CREATE TABLE IF NOT EXISTS forbidden_pairs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                first_id INTEGER NOT NULL,
                second_id INTEGER NOT NULL,
                UNIQUE(first_id, second_id),
                FOREIGN KEY(first_id) REFERENCES people(id) ON DELETE CASCADE,
                FOREIGN KEY(second_id) REFERENCES people(id) ON DELETE CASCADE
            );

            CREATE TABLE IF NOT EXISTS preferred_pairs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                autista_id INTEGER NOT NULL,
                vigile_id INTEGER NOT NULL,
                UNIQUE(autista_id, vigile_id),
                FOREIGN KEY(autista_id) REFERENCES people(id) ON DELETE CASCADE,
                FOREIGN KEY(vigile_id) REFERENCES people(id) ON DELETE CASCADE
            );

            CREATE TABLE IF NOT EXISTS vacations (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                person_id INTEGER NOT NULL,
                start_date TEXT NOT NULL,
                end_date TEXT NOT NULL,
                note TEXT,
                FOREIGN KEY(person_id) REFERENCES people(id) ON DELETE CASCADE
            );
            """
        )
        # Migrazioni incrementali: aggiungo colonne se mancanti
        self._ensure_column("people", "first_name", "TEXT")
        self._ensure_column("people", "last_name", "TEXT")
        self._ensure_column("people", "phone", "TEXT")
        self._ensure_column("people", "email", "TEXT")
        self._ensure_column("people", "ruolo", "TEXT DEFAULT ''")
        self._ensure_column("people", "grado", "TEXT")
        self._ensure_column("people", "weekly_cap", f"INTEGER DEFAULT {DEFAULT_WEEKLY_CAP}")

        self._ensure_column("forbidden_pairs", "is_hard", "INTEGER NOT NULL DEFAULT 1")
        self._ensure_column("preferred_pairs", "is_hard", "INTEGER NOT NULL DEFAULT 0")

        self.conn.commit()

    def _ensure_column(self, table: str, column: str, definition: str):
        cur = self.conn.execute(f"PRAGMA table_info({table})")
        columns = {row[1] for row in cur.fetchall()}
        if column not in columns:
            self.conn.execute(f"ALTER TABLE {table} ADD COLUMN {column} {definition}")

    # ------------------------------------------------------------------ #
    # Person management
    # ------------------------------------------------------------------ #
    def upsert_person(
        self,
        name: str,
        *,
        first_name: Optional[str] = None,
        last_name: Optional[str] = None,
        phone: Optional[str] = None,
        email: Optional[str] = None,
        ruolo: Optional[str] = None,
        grado: Optional[str] = None,
        is_autista: Optional[bool] = None,
        is_vigile: Optional[bool] = None,
        livello: Optional[str] = None,
        weekly_cap: Optional[int] = None,
    ) -> int:
        """Insert or update a person and return the person id."""
        base_name = _normalize_whitespace(name)
        if not base_name and first_name:
            base_name = _normalize_whitespace(first_name)
        if not base_name and last_name:
            base_name = _normalize_whitespace(last_name)
        if first_name is not None:
            first_name = _normalize_whitespace(first_name)
        if last_name is not None:
            last_name = _normalize_whitespace(last_name)
        if phone is not None:
            phone = _normalize_whitespace(phone)
        if email is not None:
            email = _normalize_whitespace(email)
        if ruolo is not None:
            ruolo = _normalize_whitespace(ruolo)
        if grado is not None:
            grado = _normalize_whitespace(grado)
        if not base_name:
            raise ValueError("Il nome della persona non può essere vuoto.")

        cur = self.conn.execute(
            "SELECT id, is_autista, is_vigile, livello FROM people WHERE name = ? COLLATE NOCASE",
            (base_name,),
        )
        row = cur.fetchone()

        insert_weekly_cap = weekly_cap if weekly_cap is not None else DEFAULT_WEEKLY_CAP

        if row is None:
            self.conn.execute(
                """
                INSERT INTO people (name, first_name, last_name, phone, email, ruolo, grado,
                                    is_autista, is_vigile, livello, weekly_cap)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    base_name,
                    first_name or "",
                    last_name or "",
                    phone or "",
                    email or "",
                    ruolo or "",
                    grado or "",
                    int(is_autista or False),
                    int(is_vigile or False),
                    livello or "JUNIOR",
                    insert_weekly_cap,
                ),
            )
            self.conn.commit()
            return int(self.conn.execute("SELECT last_insert_rowid()").fetchone()[0])

        updates: Dict[str, int | str] = {}
        if first_name is not None:
            updates["first_name"] = first_name
        if last_name is not None:
            updates["last_name"] = last_name
        if phone is not None:
            updates["phone"] = phone
        if email is not None:
            updates["email"] = email
        if ruolo is not None:
            updates["ruolo"] = ruolo
        if grado is not None:
            updates["grado"] = grado
        if is_autista is not None:
            updates["is_autista"] = int(is_autista)
        if is_vigile is not None:
            updates["is_vigile"] = int(is_vigile)
        if livello is not None:
            updates["livello"] = livello
        if weekly_cap is not None:
            updates["weekly_cap"] = int(weekly_cap)

        if updates:
            updates["name"] = base_name
            set_clause = ", ".join(f"{field} = ?" for field in updates.keys())
            params: List[int | str] = list(updates.values())
            params.append(row["id"])
            self.conn.execute(
                f"UPDATE people SET {set_clause} WHERE id = ?", params
            )
            self.conn.commit()

        return int(row["id"])

    def list_people(self) -> List[sqlite3.Row]:
        cur = self.conn.execute("SELECT * FROM people ORDER BY id")
        return list(cur.fetchall())

    def get_person_id(self, name: str) -> Optional[int]:
        cleaned = _normalize_whitespace(name)
        cur = self.conn.execute(
            "SELECT id FROM people WHERE name = ? COLLATE NOCASE", (cleaned,)
        )
        row = cur.fetchone()
        return int(row["id"]) if row else None

    def update_person(
        self,
        person_id: int,
        *,
        name: str,
        first_name: Optional[str],
        last_name: Optional[str],
        phone: Optional[str],
        email: Optional[str],
        ruolo: Optional[str],
        grado: Optional[str],
        is_autista: bool,
        is_vigile: bool,
        livello: str,
        weekly_cap: int,
    ):
        cleaned = _normalize_whitespace(name)
        if not cleaned:
            raise ValueError("Il nome della persona non può essere vuoto.")
        first_name = _normalize_whitespace(first_name) if first_name else ""
        last_name = _normalize_whitespace(last_name) if last_name else ""
        phone = _normalize_whitespace(phone) if phone else ""
        email = _normalize_whitespace(email) if email else ""
        ruolo = _normalize_whitespace(ruolo) if ruolo else ""
        grado = _normalize_whitespace(grado) if grado else ""

        cur = self.conn.execute("SELECT id FROM people WHERE id = ?", (person_id,))
        if cur.fetchone() is None:
            raise ValueError("Persona non trovata.")

        dup = self.conn.execute(
            "SELECT id FROM people WHERE name = ? COLLATE NOCASE AND id <> ?",
            (cleaned, person_id),
        ).fetchone()
        if dup:
            raise ValueError("Esiste già una persona con lo stesso nome.")

        self.conn.execute(
            """
            UPDATE people
            SET name = ?, first_name = ?, last_name = ?, phone = ?, email = ?, ruolo = ?, grado = ?,
                is_autista = ?, is_vigile = ?, livello = ?, weekly_cap = ?
            WHERE id = ?
            """,
            (
                cleaned,
                first_name,
                last_name,
                phone,
                email,
                ruolo,
                grado,
                int(is_autista),
                int(is_vigile),
                livello or "JUNIOR",
                int(weekly_cap),
                person_id,
            ),
        )
        self.conn.commit()

    def delete_person(self, person_id: int):
        cur = self.conn.execute("SELECT name FROM people WHERE id = ?", (person_id,))
        row = cur.fetchone()
        if not row:
            return
        name = row["name"]
        self.conn.execute("DELETE FROM people WHERE id = ?", (person_id,))
        # Se una configurazione punta al nome eliminato, la ripuliamo.
        self.conn.execute("DELETE FROM settings WHERE value = ?", (name,))
        self.conn.commit()

    # ------------------------------------------------------------------ #
    # Settings helpers
    # ------------------------------------------------------------------ #
    def set_setting(self, key: str, value: Optional[str]):
        if value is None:
            self.conn.execute("DELETE FROM settings WHERE key = ?", (key,))
        else:
            self.conn.execute(
                """
                INSERT INTO settings (key, value)
                VALUES (?, ?)
                ON CONFLICT(key) DO UPDATE SET value = excluded.value
                """,
                (key, value),
            )
        self.conn.commit()

    def get_setting(self, key: str, default: Optional[str] = None) -> Optional[str]:
        cur = self.conn.execute("SELECT value FROM settings WHERE key = ?", (key,))
        row = cur.fetchone()
        return row["value"] if row else default

    def all_settings(self) -> Dict[str, str]:
        cur = self.conn.execute("SELECT key, value FROM settings")
        return {row["key"]: row["value"] for row in cur.fetchall()}

    def _load_generation_rules(self) -> Dict[str, GenerationRuleConfig]:
        rules = build_default_rules()
        for key, definition in RULE_DEFINITIONS.items():
            mode_raw = self.get_setting(f"rule.{key}.mode")
            mode = RuleMode.from_value(mode_raw)
            value = rules[key].value
            if definition.has_value:
                val_raw = self.get_setting(f"rule.{key}.value")
                if val_raw is not None:
                    try:
                        parsed = int(val_raw)
                    except ValueError:
                        parsed = definition.default_value
                    if definition.min_value is not None:
                        parsed = max(definition.min_value, parsed)
                    if definition.max_value is not None:
                        parsed = min(definition.max_value, parsed)
                    value = parsed
            rules[key] = GenerationRuleConfig(mode=mode, value=value)
        return rules

    def load_generation_rules_config(self) -> Dict[str, GenerationRuleConfig]:
        return self._load_generation_rules()

    def save_generation_rule(self, key: str, config: GenerationRuleConfig) -> None:
        definition = RULE_DEFINITIONS.get(key)
        if definition is None:
            raise KeyError(f"Regola sconosciuta: {key}")
        self.set_setting(f"rule.{key}.mode", config.mode.value)
        if definition.has_value:
            if config.value is not None:
                self.set_setting(f"rule.{key}.value", str(config.value))
            else:
                self.set_setting(f"rule.{key}.value", None)
        else:
            self.set_setting(f"rule.{key}.value", None)

    def reset_generation_rules_to_defaults(self) -> None:
        defaults = build_default_rules()
        for key, cfg in defaults.items():
            self.save_generation_rule(key, cfg)

    # ------------------------------------------------------------------ #
    # Pairs
    # ------------------------------------------------------------------ #
    def set_forbidden_pair(self, first_id: int, second_id: int, is_hard: bool = True):
        a, b = sorted((first_id, second_id))
        if a == b:
            raise ValueError("Non è possibile creare una coppia vietata con la stessa persona.")
        self.conn.execute(
            """
            INSERT INTO forbidden_pairs (first_id, second_id, is_hard)
            VALUES (?, ?, ?)
            ON CONFLICT(first_id, second_id) DO UPDATE SET is_hard = excluded.is_hard
            """,
            (a, b, int(is_hard)),
        )
        self.conn.commit()

    def remove_forbidden_pair(self, first_id: int, second_id: int):
        a, b = sorted((first_id, second_id))
        self.conn.execute(
            "DELETE FROM forbidden_pairs WHERE first_id = ? AND second_id = ?", (a, b)
        )
        self.conn.commit()

    def list_forbidden_pairs(self) -> List[Tuple[int, int, bool]]:
        cur = self.conn.execute(
            "SELECT first_id, second_id, COALESCE(is_hard, 1) AS is_hard FROM forbidden_pairs ORDER BY first_id, second_id"
        )
        return [
            (int(row["first_id"]), int(row["second_id"]), bool(row["is_hard"]))
            for row in cur.fetchall()
        ]

    def list_forbidden_pairs_detailed(self) -> List[Tuple[int, str, str, bool]]:
        cur = self.conn.execute(
            """
            SELECT fp.id, p1.name AS first_name, p2.name AS second_name, COALESCE(fp.is_hard, 1) AS is_hard
            FROM forbidden_pairs fp
            JOIN people p1 ON p1.id = fp.first_id
            JOIN people p2 ON p2.id = fp.second_id
            ORDER BY first_name, second_name
            """
        )
        return [
            (int(row["id"]), row["first_name"], row["second_name"], bool(row["is_hard"]))
            for row in cur.fetchall()
        ]

    def delete_forbidden_pair(self, pair_id: int):
        self.conn.execute("DELETE FROM forbidden_pairs WHERE id = ?", (pair_id,))
        self.conn.commit()

    def set_preferred_pair(self, autista_id: int, vigile_id: int, is_hard: bool = False):
        if autista_id == vigile_id:
            raise ValueError("Una coppia preferita richiede due persone distinte.")
        self.conn.execute(
            """
            INSERT INTO preferred_pairs (autista_id, vigile_id, is_hard)
            VALUES (?, ?, ?)
            ON CONFLICT(autista_id, vigile_id) DO UPDATE SET is_hard = excluded.is_hard
            """,
            (autista_id, vigile_id, int(is_hard)),
        )
        self.conn.commit()

    def remove_preferred_pair(self, autista_id: int, vigile_id: int):
        self.conn.execute(
            "DELETE FROM preferred_pairs WHERE autista_id = ? AND vigile_id = ?",
            (autista_id, vigile_id),
        )
        self.conn.commit()

    def list_preferred_pairs(self) -> List[Tuple[int, int, bool]]:
        cur = self.conn.execute(
            "SELECT autista_id, vigile_id, COALESCE(is_hard, 0) AS is_hard FROM preferred_pairs ORDER BY autista_id, vigile_id"
        )
        return [
            (int(row["autista_id"]), int(row["vigile_id"]), bool(row["is_hard"]))
            for row in cur.fetchall()
        ]

    def list_preferred_pairs_detailed(self) -> List[Tuple[int, str, str, bool]]:
        cur = self.conn.execute(
            """
            SELECT pp.id, pa.name AS autista_name, pv.name AS vigile_name, COALESCE(pp.is_hard, 0) AS is_hard
            FROM preferred_pairs pp
            JOIN people pa ON pa.id = pp.autista_id
            JOIN people pv ON pv.id = pp.vigile_id
            ORDER BY autista_name, vigile_name
            """
        )
        return [
            (int(row["id"]), row["autista_name"], row["vigile_name"], bool(row["is_hard"]))
            for row in cur.fetchall()
        ]

    def delete_preferred_pair(self, pair_id: int):
        self.conn.execute("DELETE FROM preferred_pairs WHERE id = ?", (pair_id,))
        self.conn.commit()

    # ------------------------------------------------------------------ #
    # Vacations
    # ------------------------------------------------------------------ #
    def add_vacation(
        self, person_id: int, start: date, end: date, note: Optional[str] = None
    ):
        if end < start:
            raise ValueError("La data di fine ferie non può precedere la data di inizio.")
        self.conn.execute(
            """
            INSERT INTO vacations (person_id, start_date, end_date, note)
            VALUES (?, ?, ?, ?)
            """,
            (person_id, _format_date(start), _format_date(end), note),
        )
        self.conn.commit()

    def remove_vacation(self, vacation_id: int):
        self.conn.execute("DELETE FROM vacations WHERE id = ?", (vacation_id,))
        self.conn.commit()

    def list_vacations(self) -> List[sqlite3.Row]:
        cur = self.conn.execute(
            """
            SELECT v.id, v.person_id, p.name as person_name, v.start_date, v.end_date, v.note
            FROM vacations v
            JOIN people p ON p.id = v.person_id
            ORDER BY v.start_date, p.name
            """
        )
        return list(cur.fetchall())

    # ------------------------------------------------------------------ #
    # Data loading utilities
    # ------------------------------------------------------------------ #
    def load_program_config(self) -> ProgramConfig:
        people_rows = self.list_people()
        id_to_name: Dict[int, str] = {}
        autisti: List[str] = []
        vigili: List[str] = []
        esperienza: Dict[str, str] = {}
        weekly_caps: Dict[str, int] = {}
        profiles: Dict[str, PersonProfile] = {}

        for row in people_rows:
            name = row["name"]
            id_to_name[int(row["id"])] = name
            first_name = row["first_name"] or ""
            last_name = row["last_name"] or ""
            # fallback smart: se non presente ma name contiene spazi, provo a splittare
            if not first_name and name:
                parts = name.split(" ", 1)
                first_name = parts[0]
                if len(parts) > 1 and not last_name:
                    last_name = parts[1]
            weekly_cap = (
                int(row["weekly_cap"])
                if row["weekly_cap"] is not None
                else DEFAULT_WEEKLY_CAP
            )
            profile = PersonProfile(
                id=int(row["id"]),
                nome=first_name or name,
                cognome=last_name or "",
                telefono=row["phone"] or "",
                email=row["email"] or "",
                ruolo=row["ruolo"] or "",
                grado=row["grado"] or "",
                is_autista=bool(row["is_autista"]),
                is_vigile=bool(row["is_vigile"]),
                livello=row["livello"] or "JUNIOR",
                weekly_cap=weekly_cap,
            )
            profiles[name] = profile
            weekly_caps[name] = weekly_cap
            if profile.is_autista:
                autisti.append(name)
            if profile.is_vigile:
                vigili.append(name)
                esperienza[name] = profile.livello

        forbidden_pairs: List[ConstraintRule] = []
        for first_id, second_id, is_hard in self.list_forbidden_pairs():
            name1 = id_to_name.get(first_id)
            name2 = id_to_name.get(second_id)
            if name1 and name2:
                forbidden_pairs.append(
                    ConstraintRule(primo=name1, secondo=name2, is_hard=is_hard)
                )

        preferred_pairs: List[PreferredRule] = []
        for autista_id, vigile_id, is_hard in self.list_preferred_pairs():
            aut_name = id_to_name.get(autista_id)
            vig_name = id_to_name.get(vigile_id)
            if aut_name and vig_name:
                preferred_pairs.append(
                    PreferredRule(autista=aut_name, vigile=vig_name, is_hard=is_hard)
                )

        ferie: Dict[str, List[Vacation]] = {}
        for row in self.list_vacations():
            vac = Vacation(
                start=_parse_date(row["start_date"]),
                end=_parse_date(row["end_date"]),
                note=row["note"],
            )
            ferie.setdefault(row["person_name"], []).append(vac)

        key_autista_varchi = self.get_setting("autista_varchi")
        key_autista_pogliani = self.get_setting("autista_pogliani")
        key_vigile_estate = self.get_setting("vigile_escluso_estate")
        min_esperti_value = int(
            self.get_setting("min_esperti", str(DEFAULT_MIN_ESPERTI))
        )
        weekdays_setting = self.get_setting(
            "active_weekdays",
            ",".join(str(x) for x in sorted(DEFAULT_ACTIVE_WEEKDAYS)),
        )
        enable_varchi_rule_str = self.get_setting("enable_varchi_rule", "1")
        active_weekdays: Set[int] = set()
        for token in (weekdays_setting or "").split(","):
            token = token.strip()
            if token.isdigit():
                val = int(token)
                if 0 <= val <= 6:
                    active_weekdays.add(val)
        if not active_weekdays:
            active_weekdays = set(DEFAULT_ACTIVE_WEEKDAYS)
        varchi_rule_enabled = enable_varchi_rule_str != "0"
        rules = self._load_generation_rules()
        rule_min_senior = rules.get("min_senior")
        if rule_min_senior:
            if rule_min_senior.value is None:
                rule_min_senior.value = min_esperti_value
            else:
                min_esperti_value = rule_min_senior.value
        if rules.get("varchi_rotation") and rules["varchi_rotation"].mode == RuleMode.OFF:
            varchi_rule_enabled = False

        autista_varchi_name = _match_person_identifier(key_autista_varchi, autisti, profiles)
        autista_pogliani_name = _match_person_identifier(key_autista_pogliani, autisti, profiles)
        vigile_estivo_name = _match_person_identifier(key_vigile_estate, vigili, profiles)

        if varchi_rule_enabled and not autista_varchi_name:
            autista_varchi_name = _match_person_identifier(DEFAULT_AUTISTA_VARCHI, autisti, profiles)
        if varchi_rule_enabled and not autista_pogliani_name:
            autista_pogliani_name = _match_person_identifier(DEFAULT_AUTISTA_POGLIANI, autisti, profiles)
        if not vigile_estivo_name:
            vigile_estivo_name = _match_person_identifier(DEFAULT_VIGILE_ESCLUSO_ESTATE, vigili, profiles)
        if rules.get("summer_exclusion") and rules["summer_exclusion"].mode == RuleMode.OFF:
            vigile_estivo_name = None

        if not varchi_rule_enabled:
            autista_varchi_name = None
            autista_pogliani_name = None

        return ProgramConfig(
            autisti=autisti,
            vigili=vigili,
            esperienza_vigili=esperienza,
            weekly_cap=weekly_caps,
            coppie_vietate=forbidden_pairs,
            coppie_preferite=preferred_pairs,
            autista_varchi=autista_varchi_name,
            autista_pogliani=autista_pogliani_name,
            vigile_escluso_estate=vigile_estivo_name,
            min_esperti=min_esperti_value,
            ferie=ferie,
            active_weekdays=active_weekdays,
            people=profiles,
            enable_varchi_rule=varchi_rule_enabled,
            generation_rules=rules,
        )

    # ------------------------------------------------------------------ #
    # Import utility
    # ------------------------------------------------------------------ #
    def import_from_text_files(
        self,
        *,
        autisti_path: Path,
        vigili_path: Path,
        vigili_senior_path: Optional[Path] = None,
        set_defaults: bool = True,
    ):
        def _load_lines(path: Path) -> List[str]:
            if not path.exists():
                return []
            names: List[str] = []
            with path.open("r", encoding="utf-8") as handle:
                for raw in handle:
                    val = _normalize_whitespace(raw)
                    if not val or val.startswith("#"):
                        continue
                    if val not in names:
                        names.append(val)
            return names

        autisti = _load_lines(autisti_path)
        vigili_junior = _load_lines(vigili_path)
        vigili_senior = _load_lines(vigili_senior_path) if vigili_senior_path else []

        def _split(full_name: str) -> Tuple[str, str]:
            chunks = full_name.split()
            if len(chunks) >= 2:
                return chunks[0], " ".join(chunks[1:])
            return full_name, ""

        ruolo_map: Dict[str, str] = {}

        with self.conn:
            for name in autisti:
                first, last = _split(name)
                ruolo_map[name] = ROLE_AUTISTA
                self.upsert_person(
                    name,
                    first_name=first,
                    last_name=last,
                    ruolo=ROLE_AUTISTA,
                    is_autista=True,
                    is_vigile=False,
                    weekly_cap=DEFAULT_WEEKLY_CAP,
                )
            for name in vigili_junior:
                first, last = _split(name)
                ruolo = ruolo_map.get(name, ROLE_VIGILE)
                if ruolo == ROLE_AUTISTA:
                    ruolo = ROLE_AUTISTA_VIGILE
                ruolo_map[name] = ruolo
                self.upsert_person(
                    name,
                    first_name=first,
                    last_name=last,
                    ruolo=ruolo,
                    grado="JUNIOR",
                    is_vigile=True,
                    livello="JUNIOR",
                    weekly_cap=DEFAULT_WEEKLY_CAP,
                )
            for name in vigili_senior:
                first, last = _split(name)
                ruolo = ruolo_map.get(name, ROLE_VIGILE)
                if ruolo == ROLE_AUTISTA:
                    ruolo = ROLE_AUTISTA_VIGILE
                ruolo_map[name] = ruolo
                self.upsert_person(
                    name,
                    first_name=first,
                    last_name=last,
                    ruolo=ruolo,
                    grado="SENIOR",
                    is_vigile=True,
                    livello="SENIOR",
                    weekly_cap=DEFAULT_WEEKLY_CAP,
                )

        if set_defaults:
            # Try to populate settings based on common names
            if DEFAULT_AUTISTA_VARCHI in autisti:
                self.set_setting("autista_varchi", DEFAULT_AUTISTA_VARCHI)
            if DEFAULT_AUTISTA_POGLIANI in autisti:
                self.set_setting("autista_pogliani", DEFAULT_AUTISTA_POGLIANI)
            if DEFAULT_VIGILE_ESCLUSO_ESTATE in vigili_junior:
                self.set_setting(
                    "vigile_escluso_estate", DEFAULT_VIGILE_ESCLUSO_ESTATE
                )
            self.set_setting("enable_varchi_rule", "1")

            # Populate default constraints
            for a_name, b_name in DEFAULT_FORBIDDEN_PAIRS:
                a_id = self.get_person_id(a_name)
                b_id = self.get_person_id(b_name)
                if a_id and b_id:
                    self.set_forbidden_pair(a_id, b_id)
            for aut_name, vig_name in DEFAULT_PREFERRED_PAIRS:
                aut_id = self.get_person_id(aut_name)
                vig_id = self.get_person_id(vig_name)
                if aut_id and vig_id:
                    self.set_preferred_pair(aut_id, vig_id, is_hard=False)

    def close(self):
        self.conn.close()

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, tb):
        self.close()
