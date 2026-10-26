"""
Extract weekly delivery metrics from LT-SimplifiedMetrics-{project}.xlsx files.

THE BUG THIS REPLACES
----------------------
The original implementation called `pd.read_excel(file_path, header=1)`,
which reads whichever sheet pandas opens by default (the *first* sheet in
the workbook). Every `LT-SimplifiedMetrics-*.xlsx` file in this project is
actually a full copy of the same master workbook containing a sheet per
project (SEOD, Flex, CanopyNext, "Phoenix - Forms", "Consumer Web Kanban",
"WPM Scrum ", etc.) plus many unrelated sheets ("Detail1".."Detail7",
"SnapCap", "Voltron", ...). The first sheet is never the metrics sheet, so
the old code silently read garbage and returned the *same* hard-coded-looking
numbers for every project (you can see this artifact baked into the
project's own example docs, e.g. FINAL_SUMMARY.md).

THE FIX
-------
1. Find the correct sheet by its *header signature* ("Week Starting On" in
   column R / index 17 of row 0) rather than by string-matching the sheet
   name against the filename slug — sheet names ("Phoenix - Forms",
   "Consumer Web Kanban") don't match filename slugs ("phoenix",
   "consumerweb") closely enough for naive substring matching to be safe.
2. Read the *entire* weekly time series for that project (every week row),
   not just one row — downstream chart-building needs the full series.
3. Resolve "this month's" row(s) by parsing the week label
   ("SEOD 03/10 - 03/24") instead of assuming the data row is fixed at a
   particular spreadsheet row number.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Any

import pandas as pd

# Column layout of the metrics table found on every real per-project sheet.
# (0-based positions, matching the raw, header=None read of the sheet.)
METRICS_HEADER_ROW = 0
METRICS_FIRST_DATA_ROW = 1
COL_WEEK = 17  # "Week Starting On"
COL_PLANNED_SP = 18
COL_ACTUAL_DELIVERED = 19
COL_V2 = 20
COL_LT = 21
COL_REWORK = 22
COL_REOPENED_SP_LT = 23
COL_V2_CARRYOVER = 24
COL_LT_CARRYOVER = 25
COL_QA_PASSED = 26

# Key/value header block with project metadata (manager names, etc.)
# Label is in column C (index 2), value in column D (index 3). The row
# offset isn't perfectly fixed across sheets, so we *scan* for the label
# text instead of hard-coding a row number.
META_LABEL_COL = 2
META_VALUE_COL = 3
META_SCAN_ROWS = 20  # how many rows from the top to scan for labels


@dataclass
class WeekRow:
    week_label: str
    week_start: "date | None"
    week_end: "date | None"
    planned_sp: float
    actual_delivered: float
    v2: float
    lt: float
    rework: float
    reopened_sp_lt: float
    v2_carryover: float
    lt_carryover: float
    qa_passed: float


@dataclass
class ProjectGraphData:
    project_key: str  # filename slug, e.g. "seod"
    display_name: str  # e.g. "SEOD"
    source_file: str
    sheet_name: "str | None" = None
    metadata: dict[str, str] = field(default_factory=dict)
    weeks: list[WeekRow] = field(default_factory=list)
    error: "str | None" = None

    @property
    def project_manager(self) -> str:
        for label in ("Project Manager/Lead", "Project Manager", "Lead"):
            if label in self.metadata and self.metadata[label]:
                return self.metadata[label]
        return "TBD"

    @property
    def delivery_manager(self) -> str:
        return self.metadata.get("Delivery Manager") or "TBD"

    @property
    def lt_manager(self) -> "str | None":
        """The LendingTree-side account/engagement manager for this
        project. Stored in the "Client POC" field of the Graph workbook
        (that field has no other use in this sheet -- every project's
        own "Client Feedback" field, not this one, is where actual
        client feedback notes go). Returns None if blank, so callers can
        fall back to a default rather than display the literal word
        "TBD" in a name field."""
        value = self.metadata.get("Client POC")
        if value and value.strip() and value.strip() != "-":
            return value.strip()
        return None

    @property
    def sub_project(self) -> "str | None":
        """The "Sub Project" field from the Graph workbook, used as the
        dynamic project display name on slide titles and table rows.

        Shown exactly as entered in Excel, with no correction or
        cross-checking against the sheet name -- if this field is wrong
        or generic in the source file, the output will show that same
        wrong value. This is intentional, by explicit request: keeping
        this field accurate in the Graph workbook is the data owner's
        responsibility, not something this pipeline second-guesses.
        """
        value = self.metadata.get("Sub Project")
        if value and value.strip():
            return value.strip()
        return None

    @property
    def sub_project_name(self) -> "str | None":
        """This project's display name, as set in the "Sub Project"
        field of the Graph workbook (e.g. "Consumer Web", "SEOD").
        Drives the dynamic slide-title rename in build_mbr.py -- edit
        this cell in sample_data/Graph/ to rename a project everywhere
        it appears in the generated deck. Returns None if blank."""
        value = self.metadata.get("Sub Project")
        if value and value.strip():
            return value.strip()
        return None

    def weeks_in_month(self, month_name: str) -> list[WeekRow]:
        """Return weeks whose start OR end date falls in the given month
        (case-insensitive month name, e.g. "March"). Falls back to matching
        the month name as a literal token in the week label if dates
        couldn't be parsed (defensive: some week labels have stray spaces
        or unexpected formats)."""
        month_num = _MONTH_NAME_TO_NUM.get(month_name.strip().lower())
        matches = []
        for w in self.weeks:
            if month_num and (
                (w.week_start and w.week_start.month == month_num)
                or (w.week_end and w.week_end.month == month_num)
            ):
                matches.append(w)
        if matches:
            return matches
        # Fallback: nothing parsed as a date, but maybe the label mentions
        # the month name directly (defensive only; not expected in this
        # data format, which uses MM/DD ranges with no month name).
        return [w for w in self.weeks if month_name.lower() in w.week_label.lower()]

    def latest_week(self) -> "WeekRow | None":
        return self.weeks[-1] if self.weeks else None


_MONTH_NAME_TO_NUM = {
    "january": 1, "february": 2, "march": 3, "april": 4,
    "may": 5, "june": 6, "july": 7, "august": 8,
    "september": 9, "october": 10, "november": 11, "december": 12,
}

_FILENAME_RE = re.compile(r"LT-SimplifiedMetrics-(.+)\.xlsx", re.IGNORECASE)

# Reference "today" used to disambiguate two-digit week-range dates that
# straddle a year boundary (e.g. "12/16 - 12/30" vs "01/13 - 01/27"). We
# infer the year by walking from the most recent week backwards.
_DATE_RANGE_RE = re.compile(r"(\d{1,2})/(\d{1,2})\s*-\s*(\d{1,2})/(\d{1,2})")
# Some week labels embed an explicit 4-digit year up front, e.g.
# "2025 20: CAN 10/02 - 10/16" or "2026: LCW 03/01 - 03/31". When present
# this is authoritative and should override the chronological-walk guess
# below, since some projects' sprint-numbering resets create date ranges
# that look like (but aren't) a year wraparound (see CanopyNext's CAN -> SLE
# sprint-naming transition in real data, where sprint dates briefly overlap).
_LEADING_YEAR_RE = re.compile(r"^\s*(20\d{2})\b")


def extract_project_slug(filename: str) -> str:
    """'LT-SimplifiedMetrics-seod.xlsx' -> 'seod'"""
    match = _FILENAME_RE.search(filename)
    return match.group(1) if match else Path(filename).stem


def _safe_float(value: Any) -> float:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return 0.0
    try:
        return float(value)
    except (TypeError, ValueError):
        return 0.0


def _find_metrics_sheet(xl: pd.ExcelFile, prefer_name_hint: "str | None" = None) -> "str | None":
    """Find the sheet that has the real weekly metrics table, identified by
    the header signature at (row 0, col 17) == 'Week Starting On'.

    If multiple sheets qualify (the workbook contains every project's
    sheet), prefer the one whose name best matches `prefer_name_hint`
    (typically the filename slug), since that's the project this specific
    file is "about". Falls back to the first qualifying sheet found.
    """
    candidates = []
    for sheet in xl.sheet_names:
        try:
            df = xl.parse(sheet, header=None, nrows=1)
        except Exception:
            continue
        if df.shape[1] > COL_WEEK:
            header_val = df.iat[0, COL_WEEK]
            if isinstance(header_val, str) and header_val.strip().lower() == "week starting on":
                candidates.append(sheet)

    if not candidates:
        return None
    if not prefer_name_hint:
        return candidates[0]

    hint = re.sub(r"[^a-z0-9]", "", prefer_name_hint.lower())
    best, best_score = candidates[0], -1
    for sheet in candidates:
        norm = re.sub(r"[^a-z0-9]", "", sheet.lower())
        if hint in norm or norm in hint:
            score = min(len(hint), len(norm))
            if score > best_score:
                best, best_score = sheet, score
    return best


def _parse_week_range_raw(label: str):
    """Parse 'SEOD 03/10 - 03/24' -> (3, 10, 3, 24) with no year attached yet."""
    m = _DATE_RANGE_RE.search(label)
    if not m:
        return None
    return tuple(int(x) for x in m.groups())


def _leading_year(label: str):
    """Pull an explicit 4-digit year off the front of a label, if present
    (e.g. "2025 20: CAN 10/02 - 10/16" -> 2025). Returns None if absent."""
    m = _LEADING_YEAR_RE.match(label)
    return int(m.group(1)) if m else None


def _assign_years_to_weeks(raw_ranges, as_of: date, labels=None):
    """Resolve a real calendar year for every week row.

    Two strategies, in priority order:

    1. Explicit year in the label. Some projects' week labels embed the
       year directly ("2025 20: CAN 10/02 - 10/16", "2026: LCW 03/01 -
       03/31"). When present this is authoritative: trust it outright,
       because sprint/week numbering can reset or briefly overlap when a
       project renames its cadence (observed in real data -- CanopyNext's
       week labels switch from "CAN" to "SLE" sprint names mid-series with
       dates that briefly go backwards), which would otherwise be
       misread as a year wraparound by strategy 2.

    2. Chronological walk. For labels with no embedded year
       ("SEOD 03/10 - 03/24"), walk the sequence top-to-bottom
       (oldest-to-newest) and bump the year forward every time we see a
       drop in month number that isn't explained by an explicit year
       already assigned nearby. The whole relative timeline is then
       anchored so the last dated row lands on or before `as_of`.
    """
    n = len(raw_ranges)
    labels = labels or [None] * n
    explicit_years = [_leading_year(lbl) if lbl else None for lbl in labels]

    # End-year offsets, relative to a yet-unknown base year, used only for
    # rows that have no explicit year of their own.
    end_year_offset = [0] * n
    for i in range(n):
        cur = raw_ranges[i]
        if cur is None:
            end_year_offset[i] = end_year_offset[i - 1] if i > 0 else 0
            continue
        m1, _, m2, _ = cur
        offset = end_year_offset[i - 1] if i > 0 else 0
        if i > 0 and raw_ranges[i - 1] is not None and explicit_years[i] is None:
            _, _, prev_end_month, _ = raw_ranges[i - 1]
            if m1 < prev_end_month and m1 <= m2:
                offset += 1
        if m1 > m2:
            offset += 1
        end_year_offset[i] = offset

    # Anchor the *relative* (no-explicit-year) rows to a base year, chosen
    # so the last relative row's end-date doesn't land after `as_of`.
    relative_idx = [i for i in range(n) if raw_ranges[i] is not None and explicit_years[i] is None]
    base_year = as_of.year
    if relative_idx:
        last_idx = relative_idx[-1]
        _, _, end_month, end_day = raw_ranges[last_idx]
        try:
            candidate_end = date(base_year + end_year_offset[last_idx], end_month, end_day)
        except ValueError:
            candidate_end = None
        if candidate_end and candidate_end > as_of:
            base_year -= 1

    results = []
    for i, raw in enumerate(raw_ranges):
        if raw is None:
            results.append((None, None))
            continue
        m1, d1, m2, d2 = raw
        if explicit_years[i] is not None:
            year = explicit_years[i]
        else:
            year = base_year + end_year_offset[i]
        start_year = year - 1 if m1 > m2 else year
        try:
            start = date(start_year, m1, d1)
            end = date(year, m2, d2)
        except ValueError:
            start, end = None, None
        results.append((start, end))
    return results


def _read_metadata_block(df_raw: pd.DataFrame) -> dict[str, str]:
    """Scan the first META_SCAN_ROWS rows of column C/D for label/value pairs."""
    meta: dict[str, str] = {}
    n_rows = min(META_SCAN_ROWS, df_raw.shape[0])
    if df_raw.shape[1] <= META_VALUE_COL:
        return meta
    for r in range(n_rows):
        label = df_raw.iat[r, META_LABEL_COL]
        if not isinstance(label, str) or not label.strip():
            continue
        value = df_raw.iat[r, META_VALUE_COL]
        if isinstance(value, float) and pd.isna(value):
            value = ""
        meta[label.strip()] = str(value).strip() if value not in (None, "") else ""
    return meta


def read_graph_data(file_path: Path, reference_year: "int | None" = None) -> ProjectGraphData:
    """Read the full weekly metrics series + project metadata for one
    LT-SimplifiedMetrics-{project}.xlsx file."""
    slug = extract_project_slug(file_path.name)
    result = ProjectGraphData(
        project_key=slug,
        display_name=slug.replace("_", " ").upper() if slug.isupper() or slug.islower() else slug,
        source_file=file_path.name,
    )

    try:
        xl = pd.ExcelFile(file_path)
    except Exception as exc:  # corrupt / unreadable file
        result.error = f"Could not open workbook: {exc}"
        return result

    sheet_name = _find_metrics_sheet(xl, prefer_name_hint=slug)
    if sheet_name is None:
        result.error = "No sheet found with a 'Week Starting On' metrics table"
        return result

    return _read_sheet_into(result, xl, sheet_name, reference_year)


def read_graph_data_for_sheet(
    file_path: Path,
    sheet_name: str,
    display_name: "str | None" = None,
    reference_year: "int | None" = None,
) -> ProjectGraphData:
    """Read the weekly metrics series for an *explicitly named* sheet
    (e.g. from project_registry.PROJECT_REGISTRY), bypassing the
    filename-based fuzzy matching used by `read_graph_data`. Any
    LT-SimplifiedMetrics-*.xlsx file works as `file_path` since they are
    all copies of the same master workbook containing every project's
    sheet -- the registry, not the filename, determines which sheet is
    "this project"."""
    result = ProjectGraphData(
        project_key=sheet_name.strip(),
        display_name=display_name or sheet_name.strip(),
        source_file=file_path.name,
    )
    try:
        xl = pd.ExcelFile(file_path)
    except Exception as exc:
        result.error = f"Could not open workbook: {exc}"
        return result
    if sheet_name not in xl.sheet_names:
        result.error = f"Sheet '{sheet_name}' not found in workbook"
        return result
    return _read_sheet_into(result, xl, sheet_name, reference_year)


def _read_sheet_into(
    result: ProjectGraphData,
    xl: pd.ExcelFile,
    sheet_name: str,
    reference_year: "int | None",
) -> ProjectGraphData:
    """Shared logic: parse metadata + weekly rows from `sheet_name` in an
    already-open workbook, writing results onto `result`."""
    result.sheet_name = sheet_name

    try:
        df_raw = xl.parse(sheet_name, header=None)
    except Exception as exc:
        result.error = f"Could not read sheet '{sheet_name}': {exc}"
        return result

    result.metadata = _read_metadata_block(df_raw)

    as_of = date(reference_year, 12, 31) if reference_year else date.today()

    # Pass 1: collect every week row's raw fields (label still unparsed for
    # date purposes) in top-to-bottom (oldest-to-newest) order.
    row_data = []
    r = METRICS_FIRST_DATA_ROW
    while r < df_raw.shape[0]:
        week_label = df_raw.iat[r, COL_WEEK] if df_raw.shape[1] > COL_WEEK else None
        if not isinstance(week_label, str) or not week_label.strip():
            r += 1
            # Allow a couple of blank rows before giving up, in case of a
            # stray gap row, but stop once we hit a clearly empty region.
            if r < df_raw.shape[0]:
                next_label = df_raw.iat[r, COL_WEEK] if df_raw.shape[1] > COL_WEEK else None
                if not isinstance(next_label, str) or not next_label.strip():
                    break
            continue
        week_label = week_label.strip()
        row_data.append((r, week_label))
        r += 1

    # Pass 2: resolve real calendar years for every row at once (needs the
    # full sequence to detect the Dec->Jan wraparound correctly, and to
    # prefer any year embedded directly in the label).
    raw_ranges = [_parse_week_range_raw(label) for _, label in row_data]
    week_labels_only = [label for _, label in row_data]
    date_pairs = _assign_years_to_weeks(raw_ranges, as_of, labels=week_labels_only)

    weeks: list[WeekRow] = []
    for (r, week_label), (start, end) in zip(row_data, date_pairs):
        weeks.append(
            WeekRow(
                week_label=week_label,
                week_start=start,
                week_end=end,
                planned_sp=_safe_float(df_raw.iat[r, COL_PLANNED_SP]) if df_raw.shape[1] > COL_PLANNED_SP else 0.0,
                actual_delivered=_safe_float(df_raw.iat[r, COL_ACTUAL_DELIVERED]) if df_raw.shape[1] > COL_ACTUAL_DELIVERED else 0.0,
                v2=_safe_float(df_raw.iat[r, COL_V2]) if df_raw.shape[1] > COL_V2 else 0.0,
                lt=_safe_float(df_raw.iat[r, COL_LT]) if df_raw.shape[1] > COL_LT else 0.0,
                rework=_safe_float(df_raw.iat[r, COL_REWORK]) if df_raw.shape[1] > COL_REWORK else 0.0,
                reopened_sp_lt=_safe_float(df_raw.iat[r, COL_REOPENED_SP_LT]) if df_raw.shape[1] > COL_REOPENED_SP_LT else 0.0,
                v2_carryover=_safe_float(df_raw.iat[r, COL_V2_CARRYOVER]) if df_raw.shape[1] > COL_V2_CARRYOVER else 0.0,
                lt_carryover=_safe_float(df_raw.iat[r, COL_LT_CARRYOVER]) if df_raw.shape[1] > COL_LT_CARRYOVER else 0.0,
                qa_passed=_safe_float(df_raw.iat[r, COL_QA_PASSED]) if df_raw.shape[1] > COL_QA_PASSED else 0.0,
            )
        )

    result.weeks = weeks
    if not weeks:
        result.error = "Metrics sheet found but no week rows could be parsed"
    return result


def extract_all_graph_data(
    graph_folder: str = "sample_data/Graph",
    reference_year: "int | None" = None,
) -> list[ProjectGraphData]:
    """Read every LT-SimplifiedMetrics-*.xlsx file in `graph_folder`."""
    graph_path = Path(graph_folder)
    results: list[ProjectGraphData] = []

    for file_path in sorted(graph_path.glob("LT-SimplifiedMetrics-*.xlsx")):
        if file_path.name.startswith("~$"):
            continue
        results.append(read_graph_data(file_path, reference_year=reference_year))

    return results


def to_summary_dataframe(projects: list[ProjectGraphData], month: "str | None" = None) -> pd.DataFrame:
    """One row per project, summarizing either the latest week or the
    matching week(s) for `month` if given (averaging/summing duplicate
    weeks isn't appropriate for these metrics, so when a month has more
    than one matching week we report the *last* one, consistent with
    "as of end of month")."""
    rows = []
    for p in projects:
        if p.error and not p.weeks:
            rows.append({
                "Project": p.display_name,
                "Source File": p.source_file,
                "Sheet": p.sheet_name,
                "Error": p.error,
            })
            continue

        week = None
        if month:
            matches = p.weeks_in_month(month)
            week = matches[-1] if matches else None
        if week is None:
            week = p.latest_week()

        rows.append({
            "Project": p.display_name,
            "Source File": p.source_file,
            "Sheet": p.sheet_name,
            "Week": week.week_label if week else None,
            "Planned SP": week.planned_sp if week else None,
            "Actual Delivered": week.actual_delivered if week else None,
            "V2": week.v2 if week else None,
            "LT": week.lt if week else None,
            "Rework": week.rework if week else None,
            "Reopened SP (LT)": week.reopened_sp_lt if week else None,
            "V2 Carryover": week.v2_carryover if week else None,
            "LT Carryover": week.lt_carryover if week else None,
            "QA Passed": week.qa_passed if week else None,
            "Project Manager/Lead": p.project_manager,
            "Delivery Manager": p.delivery_manager,
            "Error": p.error,
        })
    return pd.DataFrame(rows)


if __name__ == "__main__":
    import sys

    folder = sys.argv[1] if len(sys.argv) > 1 else "sample_data/Graph"
    month = sys.argv[2] if len(sys.argv) > 2 else None

    projects = extract_all_graph_data(folder)
    df = to_summary_dataframe(projects, month=month)

    print("Extracted Graph Data:")
    print("=" * 100)
    print(df.to_string(index=False))
    print("\nTotal projects:", len(df))
    for p in projects:
        print(f"  {p.display_name}: sheet={p.sheet_name!r}, weeks parsed={len(p.weeks)}, error={p.error}")
