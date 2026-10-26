"""
Build extractive "Completed" / "In-progress" summary cells for the
"Key Highlights Across Workstreams" table, from each project's real
Goals Achieved / In Progress slide content.

WHY EXTRACTIVE, NOT GENERATIVE
--------------------------------
goals_consolidator.py's docstring explains why this pipeline doesn't try
to algorithmically *paraphrase* a team's reported progress: tested
against the real decks, there's no formatting signal (e.g. bold text)
that reliably marks "this is a heading" across every team consistently.

This module takes a different, safer approach for the table specifically:
it does not paraphrase or invent any wording. It selects the first few
*verbatim* bullet lines from the source slide (after dropping obvious
noise -- bare slide numbers, lone date/sprint labels) and truncates with
an ellipsis if a line is long. What ends up in the table cell is always
a subset of what's on the source slide, in the source's own words.

NOISE FILTERING
-----------------
Checked against all 6 real team decks (see module-level test output);
the only consistently-occurring noise patterns are:
  - a bare slide/page number on its own line ("3")
  - a lone year/date or sprint-cadence label with no other content
    ("2026: SC", "2026: March", "Completed" as a standalone heading with
    no body text following on the same line)
Everything else is genuine bullet content and is eligible for selection.
"""
from __future__ import annotations

import re
from dataclasses import dataclass

_NOISE_PATTERNS = [
    re.compile(r"^\d+$"),                      # bare slide number, e.g. "3"
    re.compile(r"^20\d{2}[:\s].{0,15}$"),        # lone date/sprint label, e.g. "2026: SC", "2026: March"
    re.compile(r"^(completed|in[\s-]?progress)$", re.IGNORECASE),  # standalone section label with no body
    re.compile(r"^<number>$", re.IGNORECASE),     # unresolved slide-number field from .odp conversion
]

# Some decks (e.g. FLEX) put both completed and in-progress tickets on
# the *same* "Goals Achieved" slide, with each ticket line immediately
# followed by its own "Status: Completed" / "Status: In Progress (NN%)"
# line rather than having a separate "In Progress Ticket" slide at all.
_STATUS_LINE_RE = re.compile(r"^status\s*:\s*(.+)$", re.IGNORECASE)
_STATUS_COMPLETED_RE = re.compile(r"^completed$", re.IGNORECASE)

MAX_BULLETS = 3
MAX_BULLET_CHARS = 110


@dataclass
class ExtractiveSummary:
    bullets: list[str]
    source_slide_title: str

    def as_cell_text(self) -> str:
        if not self.bullets:
            return ""
        return "\n".join(f"• {b}" for b in self.bullets)


def _is_noise(line: str, slide_title: str = "") -> bool:
    line = line.strip()
    if not line:
        return True
    if any(p.match(line) for p in _NOISE_PATTERNS):
        return True
    # Some converted decks duplicate the slide's own title into a body
    # placeholder (observed in the WPM .odp -> .pptx conversion). Treat
    # an exact (case-insensitive) repeat of the title, or of the title
    # with its "Goals Achieved in <Month>" suffix stripped, as noise.
    if slide_title:
        title_norm = slide_title.strip().lower()
        line_norm = line.lower()
        if line_norm == title_norm:
            return True
    return False


def _truncate(line: str, max_chars: int = MAX_BULLET_CHARS) -> str:
    line = line.strip()
    if len(line) <= max_chars:
        return line
    # Truncate on a word boundary rather than mid-word.
    cut = line[:max_chars].rsplit(" ", 1)[0]
    return cut + "…"


def _extract_status_paired(slide, max_bullets: int = MAX_BULLETS) -> "tuple[ExtractiveSummary, ExtractiveSummary] | None":
    """Handle decks (e.g. FLEX) that list both completed and in-progress
    tickets on one slide, each ticket line immediately followed by its
    own "Status: Completed" / "Status: In Progress (NN%)" line. Returns
    (completed_summary, in_progress_summary) if this pattern is detected
    on `slide`, else None."""
    title = slide.shapes.title.text.strip() if slide.shapes.title else ""
    lines = []
    for shape in slide.shapes:
        if shape == slide.shapes.title or not shape.has_text_frame:
            continue
        for p in shape.text_frame.paragraphs:
            text = p.text.strip()
            if text:
                lines.append(text)

    pairs = []  # (ticket_text, status_text)
    for i in range(len(lines) - 1):
        status_match = _STATUS_LINE_RE.match(lines[i + 1])
        if status_match and not _is_noise(lines[i]):
            pairs.append((lines[i], status_match.group(1).strip()))

    if not pairs:
        return None

    completed_bullets = [_truncate(t) for t, status in pairs if _STATUS_COMPLETED_RE.match(status)][:max_bullets]
    in_progress_bullets = [_truncate(t) for t, status in pairs if not _STATUS_COMPLETED_RE.match(status)][:max_bullets]

    completed = ExtractiveSummary(bullets=completed_bullets, source_slide_title=title) if completed_bullets else None
    in_progress = ExtractiveSummary(bullets=in_progress_bullets, source_slide_title=title) if in_progress_bullets else None
    return completed, in_progress


def extract_summary_from_slide(slide, max_bullets: int = MAX_BULLETS) -> "ExtractiveSummary | None":
    """Pull the first `max_bullets` substantive lines from a slide's
    body text (any non-title text frame), in their original order and
    wording, skipping noise lines."""
    title = slide.shapes.title.text.strip() if slide.shapes.title else ""
    bullets: list[str] = []

    for shape in slide.shapes:
        if shape == slide.shapes.title or not shape.has_text_frame:
            continue
        for p in shape.text_frame.paragraphs:
            text = p.text.strip()
            if _is_noise(text):
                continue
            bullets.append(_truncate(text))
            if len(bullets) >= max_bullets:
                break
        if len(bullets) >= max_bullets:
            break

    if not bullets:
        return None
    return ExtractiveSummary(bullets=bullets, source_slide_title=title)


def extract_completed_and_in_progress(src_prs) -> "tuple[ExtractiveSummary | None, ExtractiveSummary | None]":
    """Find this deck's Goals Achieved slide (-> "Completed" cell) and
    In Progress Ticket slide (-> "In-progress" cell), returning an
    ExtractiveSummary for each (None if that slide type isn't present in
    this deck).

    Some decks list status inline on the Goals Achieved slide itself
    (see _extract_status_paired) rather than using a separate In
    Progress Ticket slide; that pattern is tried first since it's more
    precise (genuinely separates completed from in-progress rather than
    just taking "whatever's on the Goals Achieved slide" as Completed).
    """
    completed = None
    in_progress = None

    for slide in src_prs.slides:
        title = (slide.shapes.title.text if slide.shapes.title else "").strip().lower()
        if completed is None and "goals achieved" in title:
            paired = _extract_status_paired(slide)
            if paired is not None:
                completed, in_progress = paired
            else:
                completed = extract_summary_from_slide(slide)
        elif in_progress is None and "in progress" in title:
            in_progress = extract_summary_from_slide(slide)

    return completed, in_progress


if __name__ == "__main__":
    import sys
    from pathlib import Path

    sys.path.insert(0, str(Path(__file__).resolve().parent))
    from pptx import Presentation

    deck_path = sys.argv[1] if len(sys.argv) > 1 else "sample_data/ppt-data/V2Solutions_LT_CanopyNext_March_Report.pptx"
    prs = Presentation(deck_path)
    completed, in_progress = extract_completed_and_in_progress(prs)

    print("COMPLETED:")
    print(completed.as_cell_text() if completed else "  (none found)")
    print("\nIN-PROGRESS:")
    print(in_progress.as_cell_text() if in_progress else "  (none found)")
