"""Definizione e gestione delle regole di generazione configurabili."""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from typing import Dict, Optional


class RuleMode(str, Enum):
    HARD = "hard"
    SOFT = "soft"
    OFF = "off"

    @classmethod
    def from_value(cls, value: Optional[str]) -> "RuleMode":
        if value is None:
            return cls.HARD
        try:
            return cls(value)
        except ValueError:
            return cls.HARD


@dataclass
class RuleDefinition:
    key: str
    label: str
    description: str
    default_mode: RuleMode = RuleMode.HARD
    has_value: bool = False
    default_value: Optional[int] = None
    min_value: Optional[int] = None
    max_value: Optional[int] = None


@dataclass
class GenerationRuleConfig:
    mode: RuleMode = RuleMode.HARD
    value: Optional[int] = None

    def as_strings(self) -> Dict[str, str]:
        data = {"mode": self.mode.value}
        if self.value is not None:
            data["value"] = str(self.value)
        return data


RULE_DEFINITIONS: Dict[str, RuleDefinition] = {
    "min_senior": RuleDefinition(
        key="min_senior",
        label="Minimo SENIOR in squadra",
        description="Numero minimo di vigili SENIOR nella squadra selezionata.",
        default_mode=RuleMode.HARD,
        has_value=True,
        default_value=1,
        min_value=0,
        max_value=4,
    ),
    "weekly_cap": RuleDefinition(
        key="weekly_cap",
        label="Limite turni settimanali",
        description="Rispetta il limite di turni settimanali per ogni persona.",
        default_mode=RuleMode.HARD,
    ),
    "summer_exclusion": RuleDefinition(
        key="summer_exclusion",
        label="Esclusione estiva dedicata",
        description="Esclude il vigile configurato da luglio e agosto.",
        default_mode=RuleMode.HARD,
    ),
    "varchi_rotation": RuleDefinition(
        key="varchi_rotation",
        label="Regola Varchi / Pogliani",
        description="Applica la regola speciale per autisti Varchi/Pogliani.",
        default_mode=RuleMode.HARD,
    ),
}


def build_default_rules() -> Dict[str, GenerationRuleConfig]:
    return {
        key: GenerationRuleConfig(
            mode=definition.default_mode,
            value=definition.default_value if definition.has_value else None,
        )
        for key, definition in RULE_DEFINITIONS.items()
    }


def merge_with_defaults(
    custom: Optional[Dict[str, GenerationRuleConfig]]
) -> Dict[str, GenerationRuleConfig]:
    base = build_default_rules()
    if not custom:
        return base
    merged: Dict[str, GenerationRuleConfig] = {}
    for key, definition in RULE_DEFINITIONS.items():
        cfg = custom.get(key)
        if cfg is None:
            merged[key] = base[key]
            continue
        value = cfg.value if definition.has_value else None
        if definition.has_value and value is None:
            value = definition.default_value
        merged[key] = GenerationRuleConfig(mode=cfg.mode, value=value)
    return merged
