"""Text normalisation for OTT titles.

``normalize_title`` produces a canonical, space-separated lowercase form.
``compact_key`` additionally removes spaces so that joined/separated words
("BiggBoss" / "Bigg Boss") become identical.
"""
from __future__ import annotations

import re
import unicodedata
from typing import Tuple

_APOSTROPHES = dict.fromkeys(map(ord, "'’‘`´"), None)
_ORDINAL_SEASON = re.compile(r"\b(\d+)(?:st|nd|rd|th)\s+season\b")
_SEASON_WORD = re.compile(r"\b(?:seasons?|ssn|sn|s)\s*0*(\d+(?:\.\d+)?)\b")
_TRAILING_NUMBER = re.compile(r"\s*\d+(?:\.\d+)?$")
_DIGITS = re.compile(r"\d+(?:\.\d+)?")


def _strip_latin_accents(text: str) -> str:
    """Remove accents from Latin letters only (keeps Devanagari etc. intact)."""
    out = []
    prev = ""
    for ch in unicodedata.normalize("NFKD", text):
        if unicodedata.combining(ch) and prev and prev.isascii() and prev.isalpha():
            continue
        out.append(ch)
        prev = ch
    return unicodedata.normalize("NFC", "".join(out))


def normalize_title(text) -> str:
    """Canonical form: case-folded, punctuation-free, season labels unified.

    Examples
    --------
    >>> normalize_title("  The  Kapil Sharma Show! ")
    'kapil sharma show'
    >>> normalize_title("Bigg Boss Season 20")
    'bigg boss 20'
    """
    if text is None:
        return ""
    s = unicodedata.normalize("NFKC", str(text))
    s = _strip_latin_accents(s).casefold()
    s = s.translate(_APOSTROPHES)
    s = s.replace("&", " and ")
    s = re.sub(r"(?<=\d)\.(?=\d)", "\x00", s)  # keep decimals such as 1.5
    s = "".join(
        ch if (unicodedata.category(ch)[0] in "LNM" or ch == "\x00") else " " for ch in s
    )
    s = s.replace("\x00", ".")
    s = re.sub(r"\s+", " ", s).strip()
    s = _ORDINAL_SEASON.sub(r"\1", s)
    s = _SEASON_WORD.sub(lambda m: m.group(1).lstrip("0") or "0", s)
    s = re.sub(r"\s+", " ", s).strip()
    if s.startswith("the ") and len(s.split()) > 1:
        s = s[4:]
    return s


def compact(normalized: str) -> str:
    """Remove all spaces from an already-normalised string."""
    return normalized.replace(" ", "")


def compact_key(text) -> str:
    return compact(normalize_title(text))


def digit_tokens(normalized: str) -> Tuple[str, ...]:
    """All numbers in a normalised title, leading zeros removed."""
    out = []
    for d in _DIGITS.findall(normalized):
        if "." not in d:
            d = d.lstrip("0") or "0"
        out.append(d)
    return tuple(out)


def strip_trailing_number(normalized: str) -> str:
    return _TRAILING_NUMBER.sub("", normalized).strip()
