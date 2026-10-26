"""
Canonical project registry for MBR report building.

WHY THIS EXISTS
----------------
Earlier scripts in this project tried to match projects across three
different naming systems by fuzzy substring matching:
  - Template slide titles ("Project Metrics - SLE", "Project Metrics –
    Consumer Experience Team (CET)")
  - Graph workbook sheet names ("CanopyNext", "Consumer Web Kanban",
    "Consumer Experience Team", "Phoenix - Forms", "WPM Scrum ")
  - Team presentation filenames ("V2Solutions_LT_CanopyNext_March_Report
    .pptx", "V2Solutions_LT_Core_SVC_March_2026_Report.pptx")

These three naming systems do not line up algorithmically:
  - The CanopyNext Graph sheet's project is called "SLE" on the slide.
  - "Consumer Web" the slide / "Consumer Web Kanban" the sheet are
    produced by a deck whose filename says "Core_SVC" and whose title
    slide says "LT Consumer Website (LCW)".
  - "Consumer Experience Team (CET)" has its own Graph sheet but no
    dedicated Graph *filename* and no team presentation deck at all.
  - "WPM" only has a presentation in legacy .odp format, not .pptx.

Rather than re-attempt fragile fuzzy matching, this module is the single,
explicit, human-verified source of truth for "which slide goes with which
sheet goes with which deck." When a new project is added, add one entry
here (verifying the sheet/file names against the real workbook/decks)
instead of hoping a heuristic guesses correctly.
"""
from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class ProjectEntry:
    # Must exactly match (case-sensitive) the project's sheet name inside
    # the LT-SimplifiedMetrics-*.xlsx workbook.
    sheet_name: str
    # Substring (case-insensitive) expected somewhere in the template's
    # "Project Metrics - ..." slide title for this project.
    slide_title_hint: str
    # Filename glob pattern (relative to the ppt-data folder) for this
    # project's monthly team deck, if one exists. None if no deck is
    # produced for this project.
    deck_glob: "str | None" = None
    # Display label used in generated reports/logs.
    display_name: "str | None" = None
    # The LendingTree-side account/engagement manager for this project.
    # IMPORTANT: this name does not exist in the Graph workbook or in any
    # team deck -- it was only ever present as static text someone typed
    # into the "LT Manager" column of the Key Highlights table itself.
    # There is no per-project signal in sample_data/ to detect this from
    # automatically. Update this field by hand when it changes; it's the
    # one piece of this pipeline's output that has no source-of-truth
    # file behind it.
    lt_manager: "str | None" = None

    def label(self) -> str:
        return self.display_name or self.sheet_name.strip()


# Verified against sample_data/Graph/LT-SimplifiedMetrics-*.xlsx (every
# file is an identical copy of the same master workbook) and
# sample_data/final-data/06-01-2026.pptx as of this writing. If your
# template or workbook structure differs, update this table to match --
# it intentionally does NOT try to auto-detect these mappings, since the
# naming is inconsistent by nature (see module docstring).
PROJECT_REGISTRY: list[ProjectEntry] = [
    ProjectEntry(
        sheet_name="Consumer Web Kanban",
        slide_title_hint="Consumer Web",
        deck_glob="*Core*SVC*.pptx",
        display_name="Consumer Web",
        lt_manager="Robert Ploeger",
    ),
    ProjectEntry(
        sheet_name="Consumer Experience Team",
        slide_title_hint="Consumer Experience Team",
        deck_glob=None,  # no team deck produced for this project
        display_name="Consumer Experience Team (CET)",
        lt_manager="Robert Ploeger",
    ),
    ProjectEntry(
        sheet_name="WPM Scrum ",  # NOTE: trailing space is part of the real sheet name
        slide_title_hint="WordPress Modernization",
        deck_glob="*WPM*.odp",  # legacy OpenDocument format; convert before reading
        display_name="WordPress Modernization (WPM)",
        lt_manager="Robert Ploeger",
    ),
    ProjectEntry(
        sheet_name="SEOD",
        slide_title_hint="SEOD",
        deck_glob="*SEOD*.pptx",
        display_name="SEOD",
        lt_manager="Robert Ploeger",
    ),
    ProjectEntry(
        sheet_name="Phoenix - Forms",
        slide_title_hint="Phoenix",
        deck_glob=None,  # no team deck produced for this project
        display_name="Phoenix - Forms",
        lt_manager="Douglas Brown",
    ),
    ProjectEntry(
        sheet_name="CanopyNext",
        slide_title_hint="SLE",  # the slide calls this project "SLE", not "CanopyNext"
        deck_glob="*CanopyNext*.pptx",
        display_name="SLE (CanopyNext)",
        lt_manager="Sushanth Kancharla",
    ),
    ProjectEntry(
        sheet_name="Flex",
        slide_title_hint="FLEX",
        deck_glob="*FLEX*.pptx",
        display_name="FLEX",
        lt_manager="David Wilson",
    ),
]

# Projects with a Graph sheet that are intentionally excluded from the
# registry because they don't appear in the template's Project Metrics
# slides (kept here for visibility/documentation, not used at runtime):
#   - "UALS (InActive)" -- has a team deck (V2Solutions_LT_UALS_*.pptx)
#     and is referenced for "Goals Achieved" consolidation, but has no
#     "Project Metrics" slide in the template.
#   - "Wordpress Modernization (InActi..." -- looks like a legacy/inactive
#     duplicate of the WPM Scrum sheet; not used.


def find_entry_by_sheet(sheet_name: str) -> "ProjectEntry | None":
    for entry in PROJECT_REGISTRY:
        if entry.sheet_name.strip().lower() == sheet_name.strip().lower():
            return entry
    return None


def find_entry_by_slide_title(slide_title: str) -> "ProjectEntry | None":
    title_lower = slide_title.lower()
    for entry in PROJECT_REGISTRY:
        if entry.slide_title_hint.lower() in title_lower:
            return entry
    return None
