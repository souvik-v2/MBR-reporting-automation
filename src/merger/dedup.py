"""
Deduplication: collapse rows that share the same ID (config.DEDUP_ID_COLUMN),
keeping only the one with the latest parsed date (config.DEDUP_DATE_COLUMN).
"""
from __future__ import annotations

from typing import List

import pandas as pd

from . import config
from .column_utils import parse_date_value


def dedup_latest(records: List[dict], log: list) -> List[dict]:
    id_col = config.DEDUP_ID_COLUMN
    date_col = config.DEDUP_DATE_COLUMN

    parsed = []
    ambiguous_examples = []
    for row in records:
        row = dict(row)
        raw_date = row.get(date_col)
        dt, is_ambiguous = parse_date_value(raw_date)
        row["__parsed_date"] = dt
        if is_ambiguous and len(ambiguous_examples) < 8:
            ambiguous_examples.append(f"'{raw_date}' in '{row.get('__source_file')}'")
        parsed.append(row)

    if ambiguous_examples:
        log.append(
            f"[WARN] {len(ambiguous_examples)}+ date value(s) are ambiguous (could be read as "
            f"day-first or month-first) and were parsed using DATE_DAYFIRST="
            f"{config.DATE_DAYFIRST} in config.py. Examples: {'; '.join(ambiguous_examples)}. "
            f"Double-check these are correct, or set DATE_DAYFIRST to match your data."
        )

    n_unparsed = sum(1 for r in parsed if r.get(date_col) not in (None, "") and r["__parsed_date"] is None)
    if n_unparsed:
        log.append(
            f"[WARN] {n_unparsed} row(s) had a value in '{date_col}' that could not be "
            f"parsed as a date; see KEEP_UNPARSEABLE_DATES in config.py."
        )

    if not config.KEEP_UNPARSEABLE_DATES:
        before = len(parsed)
        parsed = [
            r for r in parsed
            if r["__parsed_date"] is not None or r.get(date_col) in (None, "")
        ]
        removed = before - len(parsed)
        if removed:
            log.append(f"[INFO] Dropped {removed} row(s) with unparseable dates (KEEP_UNPARSEABLE_DATES=False).")

    groups: dict = {}
    no_id_rows = []
    for row in parsed:
        id_val = row.get(id_col)
        has_id = id_val is not None and str(id_val).strip() != ""
        if not has_id:
            if config.NO_ID_POLICY == "keep":
                no_id_rows.append(row)
            else:
                log.append(
                    f"[INFO] Dropped a row with no '{id_col}' value from "
                    f"'{row.get('__source_file')}' (NO_ID_POLICY=drop)."
                )
            continue
        key = str(id_val).strip().lower()
        groups.setdefault(key, []).append(row)

    deduped = []
    dup_count = 0
    for key, rows in groups.items():
        if len(rows) == 1:
            deduped.append(rows[0])
            continue
        dup_count += len(rows) - 1
        rows_with_date = [r for r in rows if r["__parsed_date"] is not None]
        if rows_with_date:
            winner = max(rows_with_date, key=lambda r: r["__parsed_date"])
        else:
            winner = rows[-1]
        deduped.append(winner)
        sources = ", ".join(sorted(set(r.get("__source_file", "?") for r in rows)))
        log.append(
            f"[INFO] ID '{key}': {len(rows)} duplicate row(s) found across [{sources}]; "
            f"kept the one from '{winner.get('__source_file')}'"
            + (f" dated {winner['__parsed_date']}" if winner["__parsed_date"] else " (no parseable date, kept last seen)")
            + "."
        )

    if dup_count:
        log.append(f"[SUMMARY] Removed {dup_count} duplicate row(s) across {len(groups)} unique ID(s).")

    result = deduped + no_id_rows
    for row in result:
        row.pop("__parsed_date", None)
    return result
