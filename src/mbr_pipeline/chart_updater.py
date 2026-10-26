"""
Update the native PowerPoint charts on each "Project Metrics - {Project}"
slide using python-pptx's `Chart.replace_data()` API.

WHY THIS REPLACES THE PROJECT'S EARLIER APPROACHES
----------------------------------------------------
The project's documentation (AUTOMATION_SUMMARY.md, CONTINUE_PLAN.md,
FINAL_MBR_AUTOMATION_README.md, MBR_STATUS.md) spent multiple iterations
trying to solve "how do we get the Excel charts into the PowerPoint" via:
  - Windows COM automation (`pptx_auto_consolidator.py`) -- Windows-only,
    requires PowerPoint installed, fragile in practice per the docs
    ("Win32COM automation has compatibility issues with Python 3.14").
  - Manual copy/paste instructions -- not automation at all.
  - matplotlib-recreated PNG images -- proposed but never built; would
    also mean the chart is a flat image, not an editable PPT chart.

None of this was necessary: the *template itself* already contains live,
editable python-pptx chart objects (`shape.has_chart`) on every Project
Metrics slide, titled exactly "Rework Trend", "Throughput Distribution",
and (on some slides) "Carryover Trend". These can be updated in place
with `chart.replace_data(...)` -- pure python-pptx, no COM, no Windows
dependency, no image recreation, and the result is still a fully native,
editable PowerPoint chart.

USAGE
-----
    update_project_charts(slide, project_graph_data, as_of_week_count=6)
"""
from __future__ import annotations

from pptx.chart.data import CategoryChartData

# Recognized chart titles -> which WeekRow fields make up each series.
# Each entry maps a chart title (case-insensitive, matched by substring)
# to a list of (series_name, field_name) pairs to pull from WeekRow.
CHART_FIELD_MAP = {
    "rework trend": [("V2", "v2_rework"), ("LT", "lt_rework")],
    "throughput distribution": [("V2", "v2"), ("LT", "lt")],
    "carryover trend": [("V2", "v2_carryover"), ("LT", "lt_carryover")],
}


def _v2_rework(week) -> float:
    """The source data has a single combined 'Rework' column, not a V2/LT
    split. The template's 'Rework Trend' chart, however, has two series
    (V2 and LT) like the other charts. Real SEOD data shows rework is
    consistently attributed to LT (V2's reopen/rework figures are 0 in
    every sampled week), so until a project provides a genuine V2/LT
    rework split we report all rework under LT and 0 under V2 rather
    than guessing a split that isn't in the source data.
    """
    return 0.0


def _lt_rework(week) -> float:
    return week.rework


_FIELD_GETTERS = {
    "v2": lambda w: w.v2,
    "lt": lambda w: w.lt,
    "v2_carryover": lambda w: w.v2_carryover,
    "lt_carryover": lambda w: w.lt_carryover,
    "v2_rework": _v2_rework,
    "lt_rework": _lt_rework,
}


def _match_chart_field_map(chart_title: str):
    title_lower = chart_title.strip().lower()
    for key, mapping in CHART_FIELD_MAP.items():
        if key in title_lower:
            return mapping
    return None


def build_chart_data(weeks, mapping) -> "CategoryChartData | None":
    """Build a CategoryChartData (categories=week labels, one series per
    mapping entry) from a list of WeekRow objects, oldest-to-newest."""
    if not weeks:
        return None
    chart_data = CategoryChartData()
    chart_data.categories = [w.week_label.strip() for w in weeks]
    for series_name, field_name in mapping:
        getter = _FIELD_GETTERS.get(field_name)
        if getter is None:
            continue
        values = [getter(w) for w in weeks]
        chart_data.add_series(series_name, values)
    return chart_data


def select_window(weeks, as_of_month: "str | None", window_size: int = 6):
    """Pick `window_size` consecutive weeks ending at the last week that
    falls in `as_of_month` (or the latest available week if no month is
    given / no weeks match)."""
    if not weeks:
        return []
    end_idx = len(weeks) - 1
    if as_of_month:
        month_idx = [i for i, w in enumerate(weeks) if w.week_start and _month_matches(w, as_of_month)]
        if month_idx:
            end_idx = month_idx[-1]
    start_idx = max(0, end_idx - window_size + 1)
    return weeks[start_idx:end_idx + 1]


def _month_matches(week, month_name: str) -> bool:
    from .graph_data_extractor import _MONTH_NAME_TO_NUM  # local import avoids a cycle at module load

    month_num = _MONTH_NAME_TO_NUM.get(month_name.strip().lower())
    if not month_num:
        return False
    return (week.week_start and week.week_start.month == month_num) or (
        week.week_end and week.week_end.month == month_num
    )


def _detach_external_data_link(chart) -> None:
    """Remove the chart's <c:externalData> link (if any) before calling
    replace_data().

    Some source templates have charts whose backing workbook is an
    *external* link (TargetMode="External", e.g. a SharePoint URL)
    rather than an embedded .xlsx part. python-pptx's
    `Chart.replace_data()` assumes an embedded workbook and raises
    `ValueError: .target_part property on _Relationship is undefined
    when target-mode is external` if it isn't. Removing the
    <c:externalData> element makes python-pptx treat the chart as having
    no existing workbook, so replace_data() adds a fresh *embedded* one
    instead -- which is what we want anyway, since the new data no longer
    corresponds to whatever lived at that external URL.
    """
    chart_space = chart._chartSpace
    chart_space._remove_externalData()


def update_project_charts(slide, project_graph_data, month: "str | None" = None, window_size: int = 6) -> list[str]:
    """Update every recognized native chart on `slide` using
    `project_graph_data.weeks`. Returns a list of human-readable log lines
    describing what was updated (or skipped, with why)."""
    log: list[str] = []
    weeks_window = select_window(project_graph_data.weeks, month, window_size=window_size)

    if not weeks_window:
        log.append(f"  (no week data available -- charts left unchanged)")
        return log

    for shape in slide.shapes:
        if not shape.has_chart:
            continue
        chart = shape.chart
        title = chart.chart_title.text_frame.text if chart.has_title else ""
        mapping = _match_chart_field_map(title)
        if mapping is None:
            log.append(f"  - {shape.name} ('{title}'): no field mapping known, left unchanged")
            continue

        chart_data = build_chart_data(weeks_window, mapping)
        if chart_data is None:
            log.append(f"  - {shape.name} ('{title}'): no data to apply, left unchanged")
            continue

        _detach_external_data_link(chart)
        chart.replace_data(chart_data)
        log.append(
            f"  - {shape.name} ('{title}'): updated with weeks "
            f"{weeks_window[0].week_label.strip()} .. {weeks_window[-1].week_label.strip()}"
        )

    return log


if __name__ == "__main__":
    import sys
    from pathlib import Path

    sys.path.insert(0, str(Path(__file__).resolve().parent))
    from pptx import Presentation
    from graph_data_extractor import read_graph_data_for_sheet
    from project_registry import PROJECT_REGISTRY, find_entry_by_slide_title

    template_path = sys.argv[1] if len(sys.argv) > 1 else "sample_data/final-data/06-01-2026.pptx"
    graph_folder = sys.argv[2] if len(sys.argv) > 2 else "sample_data/Graph"
    month = sys.argv[3] if len(sys.argv) > 3 else "March"
    out_path = sys.argv[4] if len(sys.argv) > 4 else "output/_chart_update_test.pptx"

    # Any LT-SimplifiedMetrics-*.xlsx file works as the source workbook --
    # they're all identical copies containing every project's sheet.
    any_graph_file = next(Path(graph_folder).glob("LT-SimplifiedMetrics-*.xlsx"), None)
    if any_graph_file is None:
        raise SystemExit(f"No LT-SimplifiedMetrics-*.xlsx files found in {graph_folder}")

    prs = Presentation(template_path)

    for slide in prs.slides:
        if not slide.shapes.title:
            continue
        title_text = slide.shapes.title.text
        if "project metrics" not in title_text.lower():
            continue
        entry = find_entry_by_slide_title(title_text)
        print(f"Slide: {title_text!r} -> registry match: {entry.sheet_name if entry else None}")
        if entry is None:
            continue
        project = read_graph_data_for_sheet(any_graph_file, entry.sheet_name, display_name=entry.label())
        if project.error:
            print(f"  ERROR reading sheet: {project.error}")
            continue
        for line in update_project_charts(slide, project, month=month):
            print(line)

    Path(out_path).parent.mkdir(parents=True, exist_ok=True)
    prs.save(out_path)
    print(f"\nSaved: {out_path}")
