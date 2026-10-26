"""
build_mbr.py -- fully automated Monthly Business Review deck builder.

USAGE (no arguments needed -- auto-detects everything from sample_data/)
--------------------------------------------------------------------------
    python3 build_mbr.py

This automatically finds the base template (the .pptx in
sample_data/final-data/), the reporting month (majority vote across
sample_data/ppt-data/ filenames), and writes to
output/MBR_<Month>_<Year>.pptx. Re-run the exact same command any time
new files are added to sample_data/ -- there is nothing to edit.

Override any individual piece if your layout differs from sample_data/:

    python3 build_mbr.py \\
        --template path/to/template.pptx \\
        --graph-dir path/to/Graph \\
        --ppt-data-dir path/to/ppt-data \\
        --month March \\
        --out output/MBR_March_2026.pptx

WHAT THIS DOES
---------------
1. Reads each project's weekly metrics from the Graph workbooks
   (graph_data_extractor.py), using the verified PROJECT_REGISTRY
   mapping (project_registry.py) rather than fragile name matching.
2. Updates every native chart ("Rework Trend", "Throughput
   Distribution", "Carryover Trend") on each "Project Metrics - ..."
   slide with real data for the target month (chart_updater.py).
3. For each project with a team deck, copies that team's "Goals
   Achieved" / "In Progress" slides into the consolidated deck,
   converting .odp decks to .pptx first as needed (goals_consolidator.py,
   odp_converter.py).
4. Refreshes the "Key Highlights Across Workstreams" table and each
   project's "Key Achievements" text box with extractive summaries
   pulled from that same source material (key_highlights_extractor.py,
   key_highlights_table.py) -- real wording lifted from the source
   slides, not invented or paraphrased.
5. Re-fits every "Key Achievements" text box to its slide's fixed layout
   (textbox_fitter.py / text_fit.py), shrinking font size as needed so
   a long month's content never visually overflows past the chart row
   above it or the footer line below it.
6. Saves the result and prints a full summary, including anything that
   still needs human attention (missing decks, near-miss slide titles).
"""
from __future__ import annotations

import argparse
import re
import sys
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path

_HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(_HERE))

from pptx import Presentation

from src.mbr_pipeline.chart_updater import update_project_charts
from src.mbr_pipeline.goals_consolidator import consolidate_project_slides, find_deck_for_entry, consolidate_key_achievements_only
from src.mbr_pipeline.graph_data_extractor import read_graph_data_for_sheet
from src.mbr_pipeline.project_registry import PROJECT_REGISTRY, ProjectEntry
from src.mbr_pipeline.auto_discovery import discover_inputs, default_output_path, DiscoveryError
from src.mbr_pipeline.key_highlights_extractor import extract_completed_and_in_progress
from src.mbr_pipeline.key_highlights_table import TableRowUpdate, refresh_key_highlights_tables
from src.mbr_pipeline.textbox_fitter import fit_textbox_in_place
from src.mbr_pipeline.odp_converter import ensure_pptx, ConversionError
from src.mbr_pipeline.pdf_export import export_pptx_to_pdf, PdfExportError


@dataclass
class ProjectBuildLog:
    entry: ProjectEntry
    chart_lines: list[str] = field(default_factory=list)
    manager: "str | None" = None
    deck_used: "str | None" = None
    copied_slides: list[str] = field(default_factory=list)
    near_miss_titles: list[str] = field(default_factory=list)
    key_highlights_status: "str | None" = None
    text_fit_info: "dict | None" = None
    renamed_title: "tuple[str, str] | None" = None  # (old_title, new_title), if renamed
    errors: list[str] = field(default_factory=list)


_TITLE_DATE_RE = re.compile(r"^\d{2}-[A-Za-z]+-\d{4}$")

_MONTH_NUM = {
    "january": 1, "february": 2, "march": 3, "april": 4, "may": 5, "june": 6,
    "july": 7, "august": 8, "september": 9, "october": 10, "november": 11, "december": 12,
}


def infer_report_date_from_graph_data(workbook_path: Path, month: str) -> "date | None":
    """The title slide's date should reflect the *reporting month*
    (e.g. "31-March-2026" for a March MBR), not whatever day the script
    happens to be run on -- those are frequently different (a March
    report is often built in April). The month name alone has no year,
    so the year is taken from the same Graph workbook data the rest of
    the pipeline already trusts: find any project's weeks that fall in
    `month` and use the year of the latest matching week's end date.
    Returns None if no project has data for that month (falls back to
    not touching the date stamp at all rather than guessing a year).
    """
    import calendar

    from src.mbr_pipeline.graph_data_extractor import read_graph_data_for_sheet
    from src.mbr_pipeline.project_registry import PROJECT_REGISTRY

    month_num = _MONTH_NUM.get(month.strip().lower())
    if month_num is None:
        return None

    for entry in PROJECT_REGISTRY:
        project_data = read_graph_data_for_sheet(workbook_path, entry.sheet_name, display_name=entry.label())
        if project_data.error:
            continue
        matches = project_data.weeks_in_month(month)
        for w in reversed(matches):
            if w.week_end and w.week_end.month == month_num:
                year = w.week_end.year
                last_day = calendar.monthrange(year, month_num)[1]
                return date(year, month_num, last_day)
            if w.week_start and w.week_start.month == month_num:
                year = w.week_start.year
                last_day = calendar.monthrange(year, month_num)[1]
                return date(year, month_num, last_day)

    return None


def update_title_slide_date(prs, report_date: date) -> "str | None":
    """The title slide carries a plain date string (e.g. "01-June-2026")
    in its own text box, separate from any chart data. Unlike the
    narrative text boxes this pipeline deliberately leaves alone, a date
    stamp has no editorial judgment involved -- it's safe to update
    mechanically. Returns the old date string if one was found and
    replaced, else None."""
    if not prs.slides:
        return None
    title_slide = prs.slides[0]
    new_text = report_date.strftime("%d-%B-%Y")
    for shape in title_slide.shapes:
        if not shape.has_text_frame:
            continue
        text = shape.text_frame.text.strip()
        if _TITLE_DATE_RE.match(text):
            old_text = text
            # Replace just the first run's text to preserve formatting
            # rather than rebuilding the paragraph.
            para = shape.text_frame.paragraphs[0]
            if para.runs:
                para.runs[0].text = new_text
                for run in para.runs[1:]:
                    run.text = ""
            else:
                shape.text_frame.text = new_text
            return old_text
    return None


def find_project_details_slide_index(prs) -> "int | None":
    for i, slide in enumerate(prs.slides):
        if slide.shapes.title and slide.shapes.title.text.strip().lower() == "project details":
            return i
    return None


def refresh_key_achievements_textbox(shape, bullets: list[str], max_bullets: int = 6) -> bool:
    """Replace the body bullets under the "Key Achievements and Project
    Roadmap" heading in `shape`'s text frame with real `bullets`,
    leaving the heading itself and any "Risks / Constraints" section
    (and its own heading) untouched. Returns True if a replacement was
    made, False if the expected heading structure wasn't found (in
    which case the box is left completely alone rather than guessing).
    """
    import copy as _copy

    tf = shape.text_frame
    paragraphs = list(tf.paragraphs)
    if not paragraphs:
        return False

    # Paragraph 0 should be the bold "Key Achievements..." heading.
    heading = paragraphs[0]
    if not heading.runs or not heading.runs[0].font.bold:
        return False

    # Find where the next bold paragraph (the "Risks / Constraints..."
    # sub-heading) starts; body bullets to replace are everything in
    # between. If there's no second bold heading, replace everything
    # after the first paragraph.
    risks_start = None
    for i in range(1, len(paragraphs)):
        p = paragraphs[i]
        if p.runs and p.runs[0].font.bold:
            risks_start = i
            break

    body_paragraphs = paragraphs[1:risks_start] if risks_start is not None else paragraphs[1:]
    if not body_paragraphs:
        return False

    template_pPr = body_paragraphs[0]._p.find(
        "{http://schemas.openxmlformats.org/drawingml/2006/main}pPr"
    )
    template_size = body_paragraphs[0].runs[0].font.size if body_paragraphs[0].runs else None

    # Remove the old body paragraphs.
    for p in body_paragraphs:
        p._p.getparent().remove(p._p)

    # Insert new ones, in order, right after the heading paragraph.
    anchor_el = heading._p
    for line in bullets[:max_bullets]:
        new_p_el = _copy.deepcopy(heading._p)  # clone heading purely for its <a:p> container/pPr shape
        # Strip the cloned heading's own runs and bold pPr override isn't
        # relevant here since we attach the *body's* pPr below instead.
        for run_el in new_p_el.findall("{http://schemas.openxmlformats.org/drawingml/2006/main}r"):
            new_p_el.remove(run_el)
        old_pPr = new_p_el.find("{http://schemas.openxmlformats.org/drawingml/2006/main}pPr")
        if old_pPr is not None:
            new_p_el.remove(old_pPr)
        if template_pPr is not None:
            new_p_el.insert(0, _copy.deepcopy(template_pPr))
        anchor_el.addnext(new_p_el)
        anchor_el = new_p_el

        from pptx.text.text import _Paragraph
        new_p = _Paragraph(new_p_el, tf)
        run = new_p.add_run()
        run.text = line
        if template_size is not None:
            run.font.size = template_size

    return True


_METRICS_TITLE_PREFIX_RE = re.compile(r"^(Project Metrics)\s*[-–]\s*(.+)$", re.IGNORECASE)


def rename_metrics_slide_title(slide, new_project_name: str) -> "str | None":
    """Rewrite a "Project Metrics - <name>" slide title's <name> portion
    using `new_project_name`, preserving the "Project Metrics" prefix,
    a consistent " - " separator, and the title run's existing
    formatting (font, size, bold, etc -- the text is replaced in place
    on the existing run rather than rebuilding the paragraph).

    This always re-matches against the *original* template's title text
    every run (build_mbr() loads a fresh, unmodified copy of the
    template every time -- see module docstring), so renaming here
    never breaks next month's slide-matching: there's nothing to "un-
    rename" because the rename never persists anywhere except this
    run's output file.

    Returns the old title text if a rename was made, else None (the
    title didn't match the expected "Project Metrics - ..." pattern, so
    it was left untouched rather than guessed at).
    """
    if not slide.shapes.title:
        return None
    title_shape = slide.shapes.title
    old_title = title_shape.text_frame.text
    match = _METRICS_TITLE_PREFIX_RE.match(old_title.strip())
    if not match:
        return None

    prefix = match.group(1)
    new_title = f"{prefix} - {new_project_name}"
    if new_title == old_title:
        return None

    # Replace text on the first run only, leaving its formatting (font,
    # size, bold, color) untouched; clear any additional runs in the
    # same paragraph so no leftover old text remains.
    paragraphs = title_shape.text_frame.paragraphs
    if not paragraphs or not paragraphs[0].runs:
        title_shape.text_frame.text = new_title
        return old_title

    first_run = paragraphs[0].runs[0]
    first_run.text = new_title
    for extra_run in paragraphs[0].runs[1:]:
        extra_run.text = ""
    for p in paragraphs[1:]:
        for r in p.runs:
            r.text = ""

    return old_title


def find_metrics_slide(prs, entry: ProjectEntry):
    for slide in prs.slides:
        if not slide.shapes.title:
            continue
        title = slide.shapes.title.text
        if "project metrics" in title.lower() and entry.slide_title_hint.lower() in title.lower():
            return slide
    return None


def any_graph_workbook(graph_dir: Path) -> "Path | None":
    """Fallback only: the first LT-SimplifiedMetrics-*.xlsx file found,
    used when no project-specific file matches (see
    `graph_workbook_for_entry`, which should be preferred everywhere a
    specific project's data is being read)."""
    return next(graph_dir.glob("LT-SimplifiedMetrics-*.xlsx"), None)


def graph_workbook_for_entry(graph_dir: Path, entry: ProjectEntry) -> "Path | None":
    """Pick the right workbook for `entry`.

    Every LT-SimplifiedMetrics-*.xlsx file *should* be an identical copy
    of the same master workbook (every project's sheet present in every
    file) -- but that's only true until someone edits one file without
    updating the others, which is an entirely normal thing to happen in
    practice (a team updates their own Excel; nobody re-syncs the rest).
    If this pipeline always grabbed "any" file via glob()'s alphabetical
    first match, an edit made to e.g. LT-SimplifiedMetrics-seod.xlsx
    could be silently ignored in favor of stale data still sitting in
    LT-SimplifiedMetrics-CanopyNext.xlsx's copy of the SEOD sheet.

    To avoid that, prefer the file whose name slug matches this
    project's sheet name or slide title hint (e.g. "seod" ->
    LT-SimplifiedMetrics-seod.xlsx), and only fall back to "any file" if
    no project-specific filename match exists.
    """
    import re as _re

    candidates = sorted(graph_dir.glob("LT-SimplifiedMetrics-*.xlsx"))
    if not candidates:
        return None

    target_tokens = {
        _re.sub(r"[^a-z0-9]", "", entry.sheet_name.lower()),
        _re.sub(r"[^a-z0-9]", "", entry.slide_title_hint.lower()),
    }
    for path in candidates:
        slug = path.stem.replace("LT-SimplifiedMetrics-", "")
        slug_norm = _re.sub(r"[^a-z0-9]", "", slug.lower())
        if slug_norm and any(slug_norm in t or t in slug_norm for t in target_tokens if t):
            return path

    return candidates[0]


def build_key_highlights_update(
    ppt_data_dir: Path, entry: ProjectEntry, manager: "str | None", odp_cache_dir: Path,
    lt_manager: "str | None" = None, sub_project_name: "str | None" = None,
) -> "tuple[TableRowUpdate, str]":
    """Load this project's team deck (converting .odp if needed) and
    build a TableRowUpdate with extractive Completed/In-progress
    summaries for the Key Highlights table. Returns (update, status)
    where status is a short human-readable description for logging."""
    deck_path = find_deck_for_entry(ppt_data_dir, entry.deck_glob)
    if deck_path is None:
        return (
            TableRowUpdate(
                project_label=entry.label(),
                manager_text=manager,
                lt_manager_text=lt_manager or entry.lt_manager,
                project_name_text=sub_project_name,
                completed=None,
                in_progress=None,
                note="(no team report submitted this month)",
                match_keys=[entry.slide_title_hint, entry.sheet_name],
            ),
            "no deck configured/found",
        )

    try:
        pptx_path = ensure_pptx(deck_path, odp_cache_dir)
        src_prs = Presentation(pptx_path)
    except (ConversionError, Exception) as exc:
        return (
            TableRowUpdate(
                project_label=entry.label(),
                manager_text=manager,
                lt_manager_text=lt_manager or entry.lt_manager,
                project_name_text=sub_project_name,
                completed=None,
                in_progress=None,
                note="(this month's deck could not be read -- see review checklist)",
                match_keys=[entry.slide_title_hint, entry.sheet_name],
            ),
            f"error reading deck: {exc}",
        )

    completed, in_progress = extract_completed_and_in_progress(src_prs)
    update = TableRowUpdate(
        project_label=entry.label(),
        manager_text=manager,
        lt_manager_text=lt_manager or entry.lt_manager,
        project_name_text=sub_project_name,
        completed=completed,
        in_progress=in_progress,
        match_keys=[entry.slide_title_hint, entry.sheet_name],
    )
    status = f"extracted ({len(completed.bullets) if completed else 0} completed, " \
              f"{len(in_progress.bullets) if in_progress else 0} in-progress)"
    return update, status


def build_mbr(
    template_path: Path,
    graph_dir: Path,
    ppt_data_dir: Path,
    month: str,
    out_path: Path,
    odp_cache_dir: Path,
    window_size: int = 6,
    report_date: "date | None" = None,
    insert_goal_slides: bool = False,
) -> "tuple[list[ProjectBuildLog], str | None]":
    prs = Presentation(template_path)

    old_title_date = None
    if report_date is not None:
        old_title_date = update_title_slide_date(prs, report_date)

    workbook_path = any_graph_workbook(graph_dir)
    if workbook_path is None:
        raise SystemExit(f"No LT-SimplifiedMetrics-*.xlsx files found in {graph_dir}")

    project_details_idx = find_project_details_slide_index(prs)
    if insert_goal_slides and project_details_idx is None:
        print(
            "WARNING: no slide titled exactly 'Project Details' was found; "
            "Goals Achieved / In Progress slides will not be inserted.",
            file=sys.stderr,
        )

    insert_position = project_details_idx
    logs: list[ProjectBuildLog] = []
    table_updates: dict[str, TableRowUpdate] = {}

    for entry in PROJECT_REGISTRY:
        log = ProjectBuildLog(entry=entry)

        # --- 1. Chart data -------------------------------------------------
        project_workbook = graph_workbook_for_entry(graph_dir, entry) or workbook_path
        project_data = read_graph_data_for_sheet(project_workbook, entry.sheet_name, display_name=entry.label())
        metrics_slide = None
        if project_data.error:
            log.errors.append(f"Graph data: {project_data.error}")
        else:
            log.manager = project_data.project_manager
            metrics_slide = find_metrics_slide(prs, entry)
            if metrics_slide is None:
                log.errors.append("No matching 'Project Metrics' slide found in template")
            else:
                # Rename the slide title to match the Graph workbook's
                # "Sub Project" field, shown exactly as entered there --
                # by explicit instruction, this pipeline does not
                # validate or second-guess that field. If it's wrong or
                # generic in the source Excel, the slide title will show
                # that same value; keeping it accurate is the data
                # owner's responsibility, not this code's.
                if project_data.sub_project:
                    old_title = rename_metrics_slide_title(metrics_slide, project_data.sub_project)
                    if old_title:
                        log.renamed_title = (old_title, metrics_slide.shapes.title.text)
                log.chart_lines = update_project_charts(
                    metrics_slide, project_data, month=month, window_size=window_size
                )

        # --- 2. Goals / In Progress slides (optional -- off by default)
        #        and the Key Achievements slide (always copied in if the
        #        deck has one, right after wherever the Goals/In-Progress
        #        content would go -- see consolidate_key_achievements_only
        #        for the case where the full set is skipped). The Key
        #        Highlights table (step 3) is refreshed regardless. -----
        ka_bullets = None
        if project_details_idx is not None:
            if insert_goal_slides:
                result = consolidate_project_slides(
                    prs, ppt_data_dir, entry, insert_position, odp_cache_dir,
                    include_key_achievements=True,
                )
                log.deck_used = result.deck_path.name if result.deck_path else None
                log.copied_slides = result.copied_slide_titles
                log.near_miss_titles = result.near_miss_titles
                ka_bullets = result.key_achievements_bullets
                if result.error:
                    if result.error != "No team deck configured/found for this project":
                        log.errors.append(f"Goals slides: {result.error}")
                insert_position += len(result.copied_slide_titles)
            else:
                result = consolidate_key_achievements_only(
                    prs, ppt_data_dir, entry, insert_position, odp_cache_dir,
                )
                log.deck_used = result.deck_path.name if result.deck_path else None
                log.copied_slides = result.copied_slide_titles
                ka_bullets = result.key_achievements_bullets
                if result.error and result.error != "No team deck configured/found for this project":
                    log.errors.append(f"Key Achievements slide: {result.error}")
                insert_position += len(result.copied_slide_titles)

        # --- 3. Key Highlights extractive summary (table + this project's
        #        own "Key Achievements" text isn't rewritten -- only the
        #        shared cross-project table is, since the per-slide
        #        narrative is reviewed contextually via the text-fit step
        #        below using whatever the template already says). -------
        update, kh_status = build_key_highlights_update(
            ppt_data_dir, entry, log.manager, odp_cache_dir,
            lt_manager=project_data.lt_manager, sub_project_name=project_data.sub_project,
        )
        table_updates[entry.label()] = update
        log.key_highlights_status = kh_status

        # --- 4. Refresh the "Key Achievements" text box's body bullets
        #        with real content from this project's Key Achievements
        #        slide (if one was found), then re-fit the box to its
        #        fixed size so this month's (possibly longer/shorter)
        #        text never overflows. If no real Key Achievements slide
        #        exists for this project, the box's existing text is left
        #        as-is (still re-fitted, never rewritten on guesswork). -
        if metrics_slide is not None:
            for shape in metrics_slide.shapes:
                if shape.has_text_frame and "Key Achievements" in shape.text_frame.text:
                    if ka_bullets:
                        refresh_key_achievements_textbox(shape, ka_bullets)
                    log.text_fit_info = fit_textbox_in_place(shape)
                    break

        logs.append(log)

    # --- 5. Refresh the Key Highlights table(s) now that every project's
    #        extractive summary has been built. -----------------------------
    table_log = refresh_key_highlights_tables(prs, table_updates)
    for line in table_log:
        print(line)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(out_path)
    return logs, old_title_date


def print_summary(logs: list[ProjectBuildLog], out_path: Path, old_title_date: "str | None" = None) -> None:
    print("=" * 78)
    print("MBR BUILD SUMMARY")
    print("=" * 78)

    if old_title_date:
        print(f"\nTitle slide date updated (was: {old_title_date})")

    needs_review = []

    for log in logs:
        print(f"\n[{log.entry.label()}]")
        print(f"  Project Manager/Lead: {log.manager or '(not found)'}")
        if log.renamed_title:
            old_title, new_title = log.renamed_title
            print(f"  Slide title renamed: {old_title!r} -> {new_title!r}")
        if log.chart_lines:
            print("  Charts:")
            for line in log.chart_lines:
                print(f"  {line}")
        if log.deck_used:
            print(f"  Team deck: {log.deck_used}")
            print(f"  Slides copied: {log.copied_slides or '(none matched)'}")
        else:
            print("  Team deck: (none configured for this project)")
        if log.key_highlights_status:
            print(f"  Key Highlights table: {log.key_highlights_status}")
            if log.key_highlights_status.startswith("error reading deck"):
                needs_review.append((log.entry.label(), "deck read error", log.key_highlights_status))
        if log.text_fit_info:
            info = log.text_fit_info
            print(f"  Key Achievements text box: fit at {info['font_size_pt']}pt"
                  + (" (truncated -- see source slide)" if info["truncated"] else ""))
            if info["truncated"]:
                needs_review.append((log.entry.label(), "text truncated", "Key Achievements text box still "
                                      "didn't fit at the minimum font size; trailing bullets were dropped "
                                      "with a pointer to the source slide."))
        if log.near_miss_titles:
            print(f"  >> REVIEW: near-miss slide titles not auto-copied: {log.near_miss_titles}")
            needs_review.append((log.entry.label(), "near-miss slides", log.near_miss_titles))
        if log.errors:
            for err in log.errors:
                print(f"  !! {err}")
            needs_review.append((log.entry.label(), "errors", log.errors))

    print("\n" + "=" * 78)
    if needs_review:
        print(f"ITEMS NEEDING HUMAN REVIEW ({len(needs_review)}):")
        for label, kind, detail in needs_review:
            print(f"  - {label} [{kind}]: {detail}")
    else:
        print("No items flagged for review.")
    print("=" * 78)
    print(f"\nSaved: {out_path}")


def write_review_checklist(logs: list[ProjectBuildLog], out_path: Path, month: str, insert_goal_slides: bool = False) -> Path:
    """Write a short markdown checklist of everything in this build that
    needs a human's judgment call, alongside the generated .pptx."""
    checklist_path = out_path.with_suffix("").with_name(out_path.stem + "_REVIEW_CHECKLIST.md")

    lines = [
        f"# MBR Build Review Checklist -- {month}",
        "",
        f"Generated deck: `{out_path.name}`",
        "",
        "This file lists everything in this build that genuinely needs a",
        "human's judgment call, or that an automated step had to fall back on",
        "(truncation, a missing deck, etc). Nothing below caused the build to",
        "fail -- the deck is complete and usable as produced.",
        "",
        "## \"Key Highlights\" table and \"Key Achievements\" text",
        "",
        "Both are now auto-refreshed from each team's actual Goals Achieved /",
        "In Progress slide content -- but the table cells are an *extractive*",
        "summary (the first few real bullet lines, verbatim, not a rewritten",
        "paraphrase) capped at 3 bullets per cell. If a team's most important",
        "update wasn't in the first 3 lines of their Goals Achieved slide, it",
        "won't appear in the table even though it's still in the inserted",
        "source slide further down the deck. Skim the table once; edit by",
        "hand if a different bullet should have been surfaced.",
        "",
        "## Items flagged during this build",
        "",
    ]

    any_flag = False
    for log in logs:
        flags = []
        if insert_goal_slides and not log.deck_used:
            flags.append("No team deck is configured for this project, so no Goals Achieved / "
                          "In Progress slides could be inserted. If this project has a report "
                          "this month, add it to ppt-data and register its filename pattern in "
                          "project_registry.py.")
        if log.near_miss_titles:
            flags.append(
                "This team's deck has slide(s) that look relevant but weren't auto-copied "
                f"(title didn't match the standard pattern): {log.near_miss_titles}. Check "
                "whether they should be added to the deck by hand."
            )
        if log.key_highlights_status and log.key_highlights_status.startswith("error reading deck"):
            flags.append(
                f"Could not read this project's team deck for the Key Highlights table "
                f"({log.key_highlights_status}). If this is a .odp file, install LibreOffice "
                f"(see requirements.txt) so it can be converted, or ask the team for a .pptx instead."
            )
        for err in log.errors:
            if "Goals slides" not in err:  # already covered by the "no deck" flag above
                flags.append(err)
        if flags:
            any_flag = True
            lines.append(f"### {log.entry.label()}")
            for f in flags:
                lines.append(f"- {f}")
            lines.append("")

    if not any_flag:
        lines.append("(Nothing flagged this run.)")
        lines.append("")

    checklist_path.write_text("\n".join(lines))
    return checklist_path


def main() -> None:
    parser = argparse.ArgumentParser(description="Build a fully-automated Monthly Business Review deck.")
    parser.add_argument(
        "--root",
        type=Path,
        default=Path("sample_data"),
        help="Root folder to auto-discover Graph/, ppt-data/, final-data/ under (default: sample_data)",
    )
    parser.add_argument("--template", type=Path, default=None, help="Override: path to the starting .pptx template")
    parser.add_argument("--graph-dir", type=Path, default=None, help="Override: folder of LT-SimplifiedMetrics-*.xlsx files")
    parser.add_argument("--ppt-data-dir", type=Path, default=None, help="Override: folder of team .pptx/.odp decks")
    parser.add_argument("--month", default=None, help="Override: reporting month, e.g. 'March' (default: auto-detected)")
    parser.add_argument("--out", type=Path, default=None, help="Override: output .pptx path (default: output/MBR_<Month>_<Year>.pptx)")
    parser.add_argument("--odp-cache-dir", type=Path, default=Path("output/_odp_cache"))
    parser.add_argument("--window-size", type=int, default=6, help="Number of weeks of chart history to show")
    parser.add_argument(
        "--insert-goal-slides",
        action="store_true",
        help="Also insert each team's full Goals Achieved / In Progress slides into the deck "
             "(off by default -- the Key Highlights table is always refreshed regardless of this flag)",
    )
    parser.add_argument(
        "--report-date",
        type=lambda s: date.fromisoformat(s),
        default=None,
        help="Date to stamp on the title slide, YYYY-MM-DD (default: today, leave unset to skip updating it)",
    )
    parser.add_argument(
        "--no-date-update",
        action="store_true",
        help="Skip updating the title slide's date stamp entirely",
    )
    parser.add_argument(
        "--no-pdf",
        action="store_true",
        help="Skip exporting a PDF copy of the deck (PDF export is on by default; needs LibreOffice)",
    )
    args = parser.parse_args()

    # Auto-discover anything not explicitly overridden on the command line.
    if args.template is None or args.graph_dir is None or args.ppt_data_dir is None or args.month is None:
        try:
            discovered = discover_inputs(args.root)
        except DiscoveryError as exc:
            raise SystemExit(f"Auto-discovery failed: {exc}\nPass --template/--graph-dir/--ppt-data-dir/--month explicitly instead.")
        template_path = args.template or discovered.template_path
        graph_dir = args.graph_dir or discovered.graph_dir
        ppt_data_dir = args.ppt_data_dir or discovered.ppt_data_dir
        month = args.month or discovered.month
        print(f"Auto-detected: template={template_path}, month={month} (votes: {discovered.month_vote_detail})")
    else:
        template_path, graph_dir, ppt_data_dir, month = args.template, args.graph_dir, args.ppt_data_dir, args.month

    out_path = args.out or default_output_path(month)

    if args.no_date_update:
        report_date = None
    elif args.report_date:
        report_date = args.report_date
    else:
        workbook_path = any_graph_workbook(graph_dir)
        inferred = infer_report_date_from_graph_data(workbook_path, month) if workbook_path else None
        if inferred is not None:
            report_date = inferred
        else:
            print(
                f"Could not find {month} data in the Graph workbook to infer a date from; "
                f"leaving the title slide's date stamp unchanged. Pass --report-date YYYY-MM-DD to set it explicitly."
            )
            report_date = None

    logs, old_title_date = build_mbr(
        template_path=template_path,
        graph_dir=graph_dir,
        ppt_data_dir=ppt_data_dir,
        month=month,
        out_path=out_path,
        odp_cache_dir=args.odp_cache_dir,
        window_size=args.window_size,
        report_date=report_date,
        insert_goal_slides=args.insert_goal_slides,
    )
    print_summary(logs, out_path, old_title_date)

    checklist_path = write_review_checklist(logs, out_path, month, args.insert_goal_slides)
    print(f"Review checklist: {checklist_path}")

    if not args.no_pdf:
        try:
            pdf_path = export_pptx_to_pdf(out_path)
            print(f"PDF exported: {pdf_path}")
        except PdfExportError as exc:
            print(f"PDF export skipped: {exc}")


if __name__ == "__main__":
    main()
