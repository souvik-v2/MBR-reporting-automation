"""
Consolidate each team's "Goals Achieved" / "In Progress" slides into the
master MBR deck.

DESIGN CHOICE: why this only copies slides, and never rewrites them
---------------------------------------------------------------------
Earlier scripts in this project (create_mbr_complete.py) tried to
*extract bullet text* from these slides via a heuristic ("looks like a
heading if capitalized, not a colon, 10-150 chars") and rebuild the
content as new text. Testing that heuristic against the real decks in
this project showed it isn't reliable: some teams bold their heading
bullets, some don't, and the heuristic would misclassify detail bullets
as headings (or vice versa) inconsistently across teams.

Rather than risk silently garbling a team's reported progress, this
module copies the *original slide* byte-for-byte (preserving exact
wording, formatting, and structure -- see slide_copy.py) instead of
re-deriving its content. This is both more reliable and more honest:
the consolidated deck shows exactly what each team reported.

EXTRACTION RULE
----------------
A team deck's slide is pulled into the consolidated deck if its title
matches STRICT_PATTERN: "Goals Achieved", "In Progress Ticket", or a
"Continue…." continuation slide (used by some teams to spill a long
Goals Achieved list onto a second slide). This pattern was checked
against all 6 real team decks in this project and correctly identifies
the relevant slide(s) in every one.

Slide titles that *look* related but don't match the strict pattern
(e.g. "Carryover and Reopen Details", "Completed Epic", "Carry Over") are
reported separately as "near misses" rather than silently skipped, so a
human reviewing the output knows to check whether they should be added
by hand. This project's 6 real decks have 4 such near-miss slides across
3 teams.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path

from pptx import Presentation

from .odp_converter import ConversionError, ensure_pptx
from .slide_copy import copy_slide, move_slide_to
from .key_achievements_finder import find_key_achievements_slide

STRICT_PATTERN = re.compile(r"goals achieved|^continue|in progress ticket", re.IGNORECASE)
NEARMISS_PATTERN = re.compile(
    r"carryover|carry over|reopen|completed epic|completed.*ticket", re.IGNORECASE
)


@dataclass
class ConsolidationResult:
    project_label: str
    deck_path: "Path | None" = None
    copied_slide_titles: list[str] = field(default_factory=list)
    near_miss_titles: list[str] = field(default_factory=list)
    error: "str | None" = None
    key_achievements_bullets: "list[str] | None" = None


def find_deck_for_entry(ppt_data_dir: Path, deck_glob: "str | None") -> "Path | None":
    if not deck_glob:
        return None
    matches = sorted(ppt_data_dir.glob(deck_glob))
    return matches[0] if matches else None


def consolidate_project_slides(
    dst_prs,
    ppt_data_dir: "str | Path",
    project_entry,
    insert_after_position: int,
    odp_cache_dir: "str | Path" = "output/_odp_cache",
    include_key_achievements: bool = True,
) -> ConsolidationResult:
    """Find this project's team deck (if any), pull its Goals
    Achieved / In Progress slides, and copy them into `dst_prs`
    immediately after `insert_after_position` (0-based). Returns a
    ConsolidationResult describing what happened; does not raise on a
    missing deck (that's an expected, logged case for some projects).

    If `include_key_achievements` is True (default) and the source deck
    has a real "Key Achievements" / "Key Contributions Beyond Ticket
    Goals" slide, that slide is also copied in -- placed right after the
    "In Progress Ticket" slide (or after "Goals Achieved" if this deck
    has no "In Progress Ticket" slide; see
    key_achievements_finder.find_insertion_anchor_index). Its bullet
    text is returned on the result so it can also refresh the matching
    "Project Metrics" slide's summary text box.
    """
    ppt_data_dir = Path(ppt_data_dir)
    result = ConsolidationResult(project_label=project_entry.label())

    deck_path = find_deck_for_entry(ppt_data_dir, project_entry.deck_glob)
    if deck_path is None:
        result.error = "No team deck configured/found for this project"
        return result
    result.deck_path = deck_path

    try:
        pptx_path = ensure_pptx(deck_path, odp_cache_dir)
    except ConversionError as exc:
        result.error = f"Could not read deck ({deck_path.name}): {exc}"
        return result

    try:
        src_prs = Presentation(pptx_path)
    except Exception as exc:
        result.error = f"Could not open converted deck ({pptx_path.name}): {exc}"
        return result

    insert_at = insert_after_position + 1
    for slide in src_prs.slides:
        title = (slide.shapes.title.text if slide.shapes.title else "").strip()
        if STRICT_PATTERN.search(title):
            new_slide = copy_slide(dst_prs, slide)
            move_slide_to(dst_prs, new_slide, insert_at)
            result.copied_slide_titles.append(title or "(untitled)")
            insert_at += 1
        elif NEARMISS_PATTERN.search(title):
            result.near_miss_titles.append(title)

    if include_key_achievements:
        ka_info = find_key_achievements_slide(src_prs)
        if ka_info is not None:
            ka_slide = src_prs.slides[ka_info.slide_index]
            new_slide = copy_slide(dst_prs, ka_slide)
            move_slide_to(dst_prs, new_slide, insert_at)
            result.copied_slide_titles.append(ka_info.title)
            result.key_achievements_bullets = ka_info.bullets

    return result


def consolidate_key_achievements_only(
    dst_prs,
    ppt_data_dir: "str | Path",
    project_entry,
    insert_after_position: int,
    odp_cache_dir: "str | Path" = "output/_odp_cache",
) -> ConsolidationResult:
    """Like `consolidate_project_slides`, but copies *only* the real
    "Key Achievements" / "Key Contributions Beyond Ticket Goals" slide
    (if the deck has one) -- none of the Goals Achieved / In Progress /
    Continue.... slides. Used when bulk slide insertion is turned off
    but the Key Achievements content should still be surfaced (both as
    its own slide and as the source for the metrics slide's summary text
    box -- see build_mbr.py)."""
    ppt_data_dir = Path(ppt_data_dir)
    result = ConsolidationResult(project_label=project_entry.label())

    deck_path = find_deck_for_entry(ppt_data_dir, project_entry.deck_glob)
    if deck_path is None:
        result.error = "No team deck configured/found for this project"
        return result
    result.deck_path = deck_path

    try:
        pptx_path = ensure_pptx(deck_path, odp_cache_dir)
        src_prs = Presentation(pptx_path)
    except (ConversionError, Exception) as exc:
        result.error = f"Could not read deck ({deck_path.name}): {exc}"
        return result

    ka_info = find_key_achievements_slide(src_prs)
    if ka_info is None:
        return result  # no error -- this is an expected "no such slide" case

    ka_slide = src_prs.slides[ka_info.slide_index]
    new_slide = copy_slide(dst_prs, ka_slide)
    move_slide_to(dst_prs, new_slide, insert_after_position + 1)
    result.copied_slide_titles.append(ka_info.title)
    result.key_achievements_bullets = ka_info.bullets
    return result


if __name__ == "__main__":
    import sys

    sys.path.insert(0, str(Path(__file__).resolve().parent))
    from project_registry import PROJECT_REGISTRY

    template_path = sys.argv[1] if len(sys.argv) > 1 else "sample_data/final-data/06-01-2026.pptx"
    ppt_data_dir = sys.argv[2] if len(sys.argv) > 2 else "sample_data/ppt-data"
    out_path = sys.argv[3] if len(sys.argv) > 3 else "output/_goals_consolidation_test.pptx"

    prs = Presentation(template_path)

    # Find "Project Details" -- the section divider we insert content
    # after -- by title rather than assuming a fixed slide number.
    project_details_idx = None
    for i, slide in enumerate(prs.slides):
        if slide.shapes.title and slide.shapes.title.text.strip().lower() == "project details":
            project_details_idx = i
            break
    if project_details_idx is None:
        raise SystemExit("Could not find a 'Project Details' slide in the template")

    insert_position = project_details_idx
    print(f"Inserting after slide {insert_position + 1} ('Project Details')\n")

    for entry in PROJECT_REGISTRY:
        result = consolidate_project_slides(prs, ppt_data_dir, entry, insert_position)
        print(f"=== {result.project_label} ===")
        if result.error:
            print(f"  SKIPPED: {result.error}")
        else:
            print(f"  Deck: {result.deck_path.name}")
            print(f"  Copied slides: {result.copied_slide_titles}")
            if result.near_miss_titles:
                print(f"  ** REVIEW NEEDED ** near-miss titles not auto-copied: {result.near_miss_titles}")
            insert_position += len(result.copied_slide_titles)
        print()

    Path(out_path).parent.mkdir(parents=True, exist_ok=True)
    prs.save(out_path)
    print(f"Saved: {out_path}")

    print("\nFinal slide order:")
    final_prs = Presentation(out_path)
    for i, slide in enumerate(final_prs.slides):
        title = slide.shapes.title.text if slide.shapes.title else "(no title)"
        print(f"  {i + 1}: {title!r}")
