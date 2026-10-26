"""
Refresh the "Key Highlights Across Workstreams" table (the two slides
right after the title slide) with this month's real data.

Replaces the "Completed" / "In-progress" cell content with extractive
summaries pulled from each project's actual Goals Achieved / In Progress
slide (key_highlights_extractor.py), and refreshes the "V2 Team" column
from the verified Graph data (project_registry.py /
graph_data_extractor.py) so manager-name changes show up automatically
rather than carrying over last month's names indefinitely.

The "LT Manager" column is left untouched: that's V2 Solutions' own
account-management assignment, which has no representation anywhere in
the Graph workbook or team decks -- there's no source of truth to refresh
it from. If that ever changes, edit it by hand (same as it presumably was
before this pipeline existed).
"""
from __future__ import annotations

from dataclasses import dataclass

from pptx.util import Pt

from .key_highlights_extractor import ExtractiveSummary


@dataclass
class TableRowUpdate:
    project_label: str
    manager_text: "str | None"  # new V2 Team text, or None to leave unchanged
    completed: "ExtractiveSummary | None"
    in_progress: "ExtractiveSummary | None"
    note: "str | None" = None  # e.g. "no team deck this month"
    match_keys: "list[str] | None" = None  # extra name candidates for row matching, beyond project_label
    lt_manager_text: "str | None" = None  # new LT Manager text, or None to leave unchanged
    project_name_text: "str | None" = None  # new Project column text, or None to leave unchanged


def find_key_highlights_tables(prs) -> list:
    """Return every table shape that looks like the Key Highlights table
    (header row containing "Project" and "What We Delivered" -- checked
    rather than assuming a fixed slide number, since this table can span
    more than one slide and earlier rebuilds of this deck might add/
    remove project rows)."""
    tables = []
    for slide in prs.slides:
        for shape in slide.shapes:
            if not shape.has_table:
                continue
            t = shape.table
            header_text = " ".join(c.text for c in t.rows[0].cells).lower()
            if "project" in header_text and "delivered" in header_text:
                tables.append(t)
    return tables


def _set_cell_bulleted_text(cell, lines: list[str], font_size_pt: float = 10.0) -> None:
    """Replace a table cell's content with one bulleted paragraph per
    line in `lines`, matching the existing 10pt theme-font style and
    bullet-marker formatting. If `lines` is empty, leaves a single
    "(none reported this month)" placeholder rather than an empty-
    looking cell, so it's visually obvious this isn't a missed update.
    """
    import copy

    tf = cell.text_frame
    first = tf.paragraphs[0]

    # Capture the first paragraph's <a:pPr> (bullet character/font/indent
    # settings) before clearing anything, so every new paragraph we add
    # gets the same bullet marker -- python-pptx's add_paragraph() starts
    # from a bare <a:p> with no pPr at all otherwise.
    template_pPr = first._p.find("{http://schemas.openxmlformats.org/drawingml/2006/main}pPr")

    # Clear all but the first paragraph (text frames always have >= 1),
    # then reuse that first paragraph for the first line.
    for p in list(tf.paragraphs[1:]):
        p._p.getparent().remove(p._p)

    for r in list(first.runs):
        r._r.getparent().remove(r._r)

    display_lines = lines or ["(none reported this month)"]
    for i, line in enumerate(display_lines):
        p = first if i == 0 else tf.add_paragraph()
        if i > 0 and template_pPr is not None:
            # add_paragraph()'s new <a:p> has no <a:pPr> yet; insert a
            # copy of the template one as the first child (pPr must
            # precede any <a:r> run elements per the schema).
            p._p.insert(0, copy.deepcopy(template_pPr))
        run = p.add_run()
        run.text = line
        run.font.size = Pt(font_size_pt)


def apply_row_update(table, row_index: int, update: TableRowUpdate, completed_col: int = 3, in_progress_col: int = 4, manager_col: int = 2, lt_manager_col: int = 1, project_col: int = 0) -> None:
    """Apply one project's refreshed data to a specific row of the Key
    Highlights table."""
    if update.lt_manager_text:
        _set_cell_bulleted_text(table.cell(row_index, lt_manager_col), [update.lt_manager_text])
    if update.manager_text:
        _set_cell_bulleted_text(table.cell(row_index, manager_col), [update.manager_text])

    completed_lines = update.completed.bullets if update.completed else []
    in_progress_lines = update.in_progress.bullets if update.in_progress else []

    if update.note and not completed_lines and not in_progress_lines:
        completed_lines = [update.note]

    _set_cell_bulleted_text(table.cell(row_index, completed_col), completed_lines)
    _set_cell_bulleted_text(table.cell(row_index, in_progress_col), in_progress_lines)

    # Renamed last -- row matching (which happens before this function is
    # called) relies on the *original* "Project" cell text, so the rename
    # must never affect anything matching still needs to do this run.
    if update.project_name_text:
        _set_cell_bulleted_text(table.cell(row_index, project_col), [update.project_name_text])


def refresh_key_highlights_tables(prs, updates_by_project_label: dict[str, TableRowUpdate]) -> list[str]:
    """Find every Key Highlights table in `prs` and refresh each row
    whose "Project" cell (column 0) matches a project in
    `updates_by_project_label` (case-insensitive, substring match in
    either direction, tried against every name in that project's
    `match_keys` as well as its main label -- the table's project names
    are short, inconsistent labels like "WPM Scrum" or "SLE" that don't
    reliably substring-match any single one of {registry display name,
    slide title hint, Graph sheet name} for every project, so all three
    are tried). Returns a log of what was/wasn't matched."""
    log = []
    matched_labels = set()

    for table in find_key_highlights_tables(prs):
        # Data rows start after the two header rows (row 0: "Project" /
        # "What We Delivered..." spanning header; row 1: "Completed" /
        # "In-progress" sub-header).
        for row_idx in range(2, len(table.rows)):
            project_cell_text = table.cell(row_idx, 0).text.strip()
            if not project_cell_text:
                continue
            match_label = None
            for label, update in updates_by_project_label.items():
                candidates = [label] + (update.match_keys or [])
                for candidate in candidates:
                    if candidate.lower() in project_cell_text.lower() or project_cell_text.lower() in candidate.lower():
                        match_label = label
                        break
                if match_label:
                    break
            if match_label is None:
                log.append(f"  No data match for table row {project_cell_text!r}; left unchanged")
                continue
            apply_row_update(table, row_idx, updates_by_project_label[match_label])
            matched_labels.add(match_label)
            log.append(f"  Updated row {project_cell_text!r} from {match_label!r}")

    for label in updates_by_project_label:
        if label not in matched_labels:
            log.append(f"  WARNING: {label!r} had no matching table row (table structure may have changed)")

    return log
