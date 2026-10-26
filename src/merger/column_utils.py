"""
Utilities for reconciling inconsistent column names across many files
and for parsing dates leniently.
"""
from __future__ import annotations

import re
import unicodedata
from datetime import datetime

import pandas as pd

from . import config


def normalize_col_name(name: str) -> str:
    """Lowercase, strip, collapse internal whitespace/punctuation for matching."""
    name = str(name)
    name = unicodedata.normalize("NFKD", name)
    name = name.strip().lower()
    name = re.sub(r"[\s_\-/]+", " ", name)
    name = re.sub(r"[^\w ]", "", name)
    return name.strip()


def build_alias_lookup() -> dict:
    """Map normalized alias -> canonical column name, from config.COLUMN_ALIASES."""
    lookup = {}
    for canonical, aliases in config.COLUMN_ALIASES.items():
        lookup[normalize_col_name(canonical)] = canonical
        for alias in aliases:
            lookup[normalize_col_name(alias)] = canonical
    return lookup


_ALIAS_LOOKUP = build_alias_lookup()


def canonicalize_column(col_name: str, first_seen_casing: dict) -> str:
    """
    Return the canonical name for a column. Known aliases (config.py) map to
    their configured canonical name. Unlisted columns are matched
    case/whitespace-insensitively against any column already seen earlier
    in this same merge run, so 'Customer' and 'customer' collapse together
    using whichever casing appeared first.
    """
    norm = normalize_col_name(col_name)
    if norm in _ALIAS_LOOKUP:
        return _ALIAS_LOOKUP[norm]
    if norm in first_seen_casing:
        return first_seen_casing[norm]
    canon = str(col_name).strip()
    first_seen_casing[norm] = canon
    return canon


# Persists across the whole run so the SAME unlisted column name (e.g.
# "customer" vs "Customer") always canonicalizes to one casing, no matter
# which file is read first.
_GLOBAL_CASING_MAP: dict = {}


def canonicalize_columns(columns) -> list:
    seen = {}
    result = []
    for c in columns:
        canon = canonicalize_column(c, _GLOBAL_CASING_MAP)
        if canon in seen:
            seen[canon] += 1
            canon = f"{canon} ({seen[canon]})"
        else:
            seen[canon] = 0
        result.append(canon)
    return result


def reset_casing_map():
    """Call at the start of a run if reusing this module across multiple merges."""
    _GLOBAL_CASING_MAP.clear()


def parse_date_value(value):
    """Best-effort parse of a date/time value from mixed-format sources.

    Returns a (datetime_or_None, is_ambiguous) tuple. is_ambiguous is True
    when the text could plausibly mean two different dates depending on
    day-first vs month-first interpretation (e.g. '07/01/2026'), so the
    caller can warn the user rather than silently picking one.
    """
    if value is None:
        return None, False
    if isinstance(value, datetime):
        return value, False
    if isinstance(value, pd.Timestamp):
        if pd.isna(value):
            return None, False
        return value.to_pydatetime(), False
    if isinstance(value, (int, float)):
        try:
            if 20000 < value < 80000:
                return pd.to_datetime(value, unit="D", origin="1899-12-30").to_pydatetime(), False
        except Exception:
            return None, False
        return None, False

    text = str(value).strip()
    if not text or text.lower() in ("nan", "none", "nat"):
        return None, False

    # ISO format (yyyy-mm-dd...) is never day/month ambiguous.
    is_iso = bool(re.match(r"^\d{4}-\d{1,2}-\d{1,2}([ T].*)?$", text))

    import warnings
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        try:
            chosen = pd.to_datetime(text, errors="raise", dayfirst=config.DATE_DAYFIRST)
        except Exception:
            return None, False
        if pd.isna(chosen):
            return None, False

        is_ambiguous = False
        if not is_iso:
            try:
                other = pd.to_datetime(text, errors="raise", dayfirst=not config.DATE_DAYFIRST)
                if not pd.isna(other) and other != chosen:
                    is_ambiguous = True
            except Exception:
                pass

    return chosen.to_pydatetime(), is_ambiguous
