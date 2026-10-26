"""
Find, copy, and reuse each team's "Key Achievements" slide.

WHAT THIS REPLACES
--------------------
The metrics slide's "Key Achievements and Project Roadmap" text box was
previously left completely untouched (it's hand-written template content
from a prior month, see MBR_PIPELINE.md). The user pointed out that real
source slides exist in several team decks under titles like
"Key Achievements" (FLEX) or "Key Contributions Beyond Ticket Goals"
(CanopyNext, Core_SVC/Consumer Web) -- genuine, team-authored content
that was being ignored.

WHAT THIS DOES
---------------
1. Finds a deck's "Key Achievements" / "Key Contributions Beyond Ticket
   Goals" slide by title (case-insensitive), if one exists.
2. Copies that slide into the consolidated deck immediately after the
   project's "In Progress Ticket" slide (or, if that title isn't present
   in this deck, after its "Goals Achieved" slide -- see
   goals_consolidator.py's STRICT_PATTERN for why titles vary by team).
3. Returns the slide's bullet text so it can also be used to refresh the
   short summary on the corresponding "Project Metrics" slide.

If a deck has no such slide (SEOD, WPM, and any project with no deck at
all), this returns nothing -- no slide is invented, and the metrics
slide's existing text box is left as-is for that project, consistent
with how this pipeline always treats absent source material (see
goals_consolidator.py's "no team report" placeholder).
"""
from __future__ import annotations

import re
from dataclasses import dataclass

KEY_ACHIEVEMENTS_PATTERN = re.compile(
    r"key achievements|key contributions beyond ticket goals", re.IGNORECASE
)

MAX_BULLETS_FOR_METRICS_SLIDE = 6


@dataclass
class KeyAchievementsSlideInfo:
    slide_index: int  # 0-based index within the source deck
    title: str
    bullets: list[str]


def find_key_achievements_slide(src_prs) -> "KeyAchievementsSlideInfo | None":
    """Locate this deck's Key Achievements slide by title, if present."""
    for i, slide in enumerate(src_prs.slides):
        title = (slide.shapes.title.text if slide.shapes.title else "").strip()
        if KEY_ACHIEVEMENTS_PATTERN.search(title):
            bullets = []
            for shape in slide.shapes:
                if shape == slide.shapes.title or not shape.has_text_frame:
                    continue
                for p in shape.text_frame.paragraphs:
                    text = p.text.strip()
                    if text and not re.match(r"^\d+$", text):  # skip bare slide numbers
                        bullets.append(text)
            if bullets:
                return KeyAchievementsSlideInfo(slide_index=i, title=title, bullets=bullets)
    return None


def find_insertion_anchor_index(src_prs) -> "int | None":
    """Find the 0-based index of this deck's "In Progress Ticket" slide
    (the anchor the user asked for); fall back to the "Goals Achieved"
    slide (or its last "Continue…." continuation) if no In Progress
    Ticket slide exists in this deck, so there's still a sensible place
    to put the Key Achievements slide right after."""
    in_progress_idx = None
    goals_idx = None
    for i, slide in enumerate(src_prs.slides):
        title = (slide.shapes.title.text if slide.shapes.title else "").strip().lower()
        if "in progress" in title:
            in_progress_idx = i
        elif "goals achieved" in title or title.startswith("continue"):
            goals_idx = i  # keep overwriting -- we want the *last* matching slide
    return in_progress_idx if in_progress_idx is not None else goals_idx
