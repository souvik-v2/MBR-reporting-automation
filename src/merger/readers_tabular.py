"""
Readers for Excel (.xlsx/.xlsm/.xls) and CSV/TSV files.

Every row read is returned as a dict, tagged with __source_file and
__source_sheet, with column names already canonicalized.
"""
from __future__ import annotations

from pathlib import Path
from typing import List

import pandas as pd

from .column_utils import canonicalize_columns


def _frame_to_records(df: pd.DataFrame, source_file: str, source_sheet: str) -> List[dict]:
    if df.empty:
        return []
    df = df.copy()
    df.columns = canonicalize_columns(df.columns)
    df = df.dropna(how="all")
    records = []
    for row in df.to_dict(orient="records"):
        row = {k: (None if pd.isna(v) else v) for k, v in row.items()}
        row["__source_file"] = source_file
        row["__source_sheet"] = source_sheet
        records.append(row)
    return records


def read_excel_file(path: Path, log: list) -> List[dict]:
    records = []
    try:
        sheets = pd.read_excel(path, sheet_name=None, dtype=object)
    except Exception as exc:
        log.append(f"[ERROR] Could not read Excel file '{path.name}': {exc}")
        return []

    for sheet_name, df in sheets.items():
        try:
            recs = _frame_to_records(df, path.name, sheet_name)
            records.extend(recs)
            if not recs:
                log.append(f"[INFO] Sheet '{sheet_name}' in '{path.name}' was empty, skipped.")
        except Exception as exc:
            log.append(f"[ERROR] Failed parsing sheet '{sheet_name}' in '{path.name}': {exc}")
    return records


def read_csv_file(path: Path, log: list) -> List[dict]:
    sep = "\t" if path.suffix.lower() == ".tsv" else None
    try:
        df = pd.read_csv(path, sep=sep, engine="python", dtype=object)
    except Exception as exc:
        log.append(f"[ERROR] Could not read CSV file '{path.name}': {exc}")
        return []
    try:
        recs = _frame_to_records(df, path.name, "csv")
        if not recs:
            log.append(f"[INFO] CSV file '{path.name}' had no data rows, skipped.")
        return recs
    except Exception as exc:
        log.append(f"[ERROR] Failed parsing CSV file '{path.name}': {exc}")
        return []
