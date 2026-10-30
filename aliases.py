"""Recognised title variants (aliases) and forbidden pairs (exceptions).

Add new aliases here, or at run time via ``AliasTable.add`` / the JSON box in
the web app.  Each canonical title maps to a list of variants that should be
treated as the same show.  Numbers (season labels) are handled separately, so
list only the base title ("Bigg Boss", not "Bigg Boss 20").

JSON format accepted by ``load_custom_rules``::

    {
      "aliases":    {"Bigg Boss": ["BB", "Big Boss"]},
      "exceptions": [["Bigg Boss", "Bigg Boss OTT"]]
    }

Exceptions are pairs that must NEVER be matched, however similar they look.
"""
from __future__ import annotations

import json
from typing import Dict, Iterable, List, Tuple

from .normalize import compact_key, normalize_title, strip_trailing_number, compact

DEFAULT_ALIASES: Dict[str, List[str]] = {
    "Bigg Boss": ["Big Boss", "Bigboss", "BB"],
    "The Kapil Sharma Show": ["TKSS"],
    "Kaun Banega Crorepati": ["KBC", "Kaun Banega Crorpati"],
    "Taarak Mehta Ka Ooltah Chashmah": [
        "TMKOC",
        "Tarak Mehta Ka Ulta Chashma",
        "Taarak Mehta",
        "Tarak Mehta",
    ],
    "Yeh Rishta Kya Kehlata Hai": ["YRKKH"],
    "Khatron Ke Khiladi": ["KKK", "Fear Factor Khatron Ke Khiladi"],
    "Game of Thrones": ["GOT"],
    "Money Heist": ["La Casa de Papel"],
    "MTV Roadies": ["Roadies"],
    "MTV Splitsvilla": ["Splitsvilla"],
    "Naagin": ["Nagin"],
    "Anupamaa": ["Anupama"],
}

# Pairs that must never match (checked before any other rule).
DEFAULT_EXCEPTIONS: List[Tuple[str, str]] = []


class AliasTable:
    """Maps compact title keys to a shared group id."""

    def __init__(self, groups: Dict[str, Iterable[str]] | None = None):
        self._key_to_group: Dict[str, str] = {}
        for canonical, variants in (groups or {}).items():
            self.add(canonical, variants)

    def add(self, canonical: str, variants: Iterable[str] = ()) -> None:
        base = strip_trailing_number(normalize_title(canonical))
        gid = compact(base)
        if not gid:
            return
        gid = self._key_to_group.get(gid, gid)
        for text in [canonical, *variants]:
            key = compact(strip_trailing_number(normalize_title(text)))
            if key:
                self._key_to_group[key] = gid

    def group_of(self, key: str) -> str:
        return self._key_to_group.get(key, key)

    def __len__(self) -> int:
        return len(self._key_to_group)


def default_alias_table() -> AliasTable:
    return AliasTable(DEFAULT_ALIASES)


def load_custom_rules(text: str) -> Tuple[Dict[str, List[str]], List[Tuple[str, str]]]:
    """Parse the optional JSON rules.  Raises ``ValueError`` with a clear message."""
    text = (text or "").strip()
    if not text:
        return {}, []
    try:
        data = json.loads(text)
    except json.JSONDecodeError as exc:
        raise ValueError(f"Custom rules are not valid JSON: {exc}") from exc
    if not isinstance(data, dict):
        raise ValueError("Custom rules must be a JSON object with 'aliases' and/or 'exceptions'.")
    aliases = data.get("aliases", {}) or {}
    exceptions = data.get("exceptions", []) or []
    if not isinstance(aliases, dict) or not all(
        isinstance(v, list) for v in aliases.values()
    ):
        raise ValueError("'aliases' must map a title to a list of variants.")
    if not isinstance(exceptions, list) or not all(
        isinstance(p, (list, tuple)) and len(p) == 2 for p in exceptions
    ):
        raise ValueError("'exceptions' must be a list of [titleA, titleB] pairs.")
    return (
        {str(k): [str(x) for x in v] for k, v in aliases.items()},
        [(str(a), str(b)) for a, b in exceptions],
    )
