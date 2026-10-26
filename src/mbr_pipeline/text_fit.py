"""
Fit text into a fixed-size PowerPoint text box.

WHY THIS EXISTS
----------------
The MBR template's "Key Achievements and Project Roadmap" text boxes are
sized to fit *whatever content existed when the template was last hand-
edited*. Every month's real content is a different length, and the box
cannot grow (the slide layout/master has a footer line at a fixed
vertical position -- see MBR_PIPELINE.md for the measured numbers), so a
long month's content silently overflows past the box and visually
collides with the chart row above it or the footer below it.

This module estimates, using real font metrics (PIL/Pillow with a
metric-compatible substitute for the deck's Calibri theme font), how
many wrapped lines a block of text will take at a given font size and
box width, then searches downward from the box's current font size to
the largest size that fits within the box's fixed height -- falling
back to truncating the content (with a "...see source slide" pointer)
only if even the smallest acceptable font size still overflows.
"""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from PIL import ImageFont

# Calibri (the deck's theme font) isn't installed in this environment;
# Liberation Sans is metrically close enough for *estimating* wrapped
# line counts (it's designed as an Arial substitute, similar x-height/
# advance-width family to Calibri for this purpose). This only affects
# the font-size/truncation decision, not the actual rendered output --
# PowerPoint/LibreOffice still render with the real theme font.
_FONT_CANDIDATES = [
    "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
    "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
]
_BOLD_FONT_CANDIDATES = [
    "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf",
    "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
]

EMU_PER_INCH = 914400
EMU_PER_PT = 12700


def _first_existing(paths: list[str]) -> "str | None":
    for p in paths:
        if Path(p).exists():
            return p
    return None


_REGULAR_FONT_PATH = _first_existing(_FONT_CANDIDATES)
_BOLD_FONT_PATH = _first_existing(_BOLD_FONT_CANDIDATES)

_font_cache: dict[tuple[str, int], "ImageFont.FreeTypeFont"] = {}


def _get_font(size_pt: float, bold: bool = False) -> "ImageFont.FreeTypeFont | None":
    path = _BOLD_FONT_PATH if bold else _REGULAR_FONT_PATH
    if path is None:
        return None
    key = (path, round(size_pt * 4))  # quarter-point cache resolution
    if key not in _font_cache:
        # PIL fonts are sized in pixels at 72dpi == points; round to int px.
        _font_cache[key] = ImageFont.truetype(path, size=max(1, round(size_pt)))
    return _font_cache[key]


@dataclass
class FitResult:
    font_size_pt: float
    truncated: bool
    paragraphs: list[str]  # possibly truncated, in final form to write


def _wrapped_line_count(text: str, font, max_width_px: float) -> int:
    """Estimate how many lines `text` wraps to at `font`'s size within
    `max_width_px`, using greedy word-wrap (matches how PowerPoint/
    LibreOffice wrap text closely enough for sizing decisions)."""
    if not text:
        return 1
    words = text.split(" ")
    lines = 1
    current = ""
    for word in words:
        candidate = f"{current} {word}".strip()
        width = font.getlength(candidate)
        if width <= max_width_px or not current:
            current = candidate
        else:
            lines += 1
            current = word
    return lines


def estimate_required_height_in(
    paragraphs: list[tuple[str, bool]],
    box_width_in: float,
    font_size_pt: float,
    line_spacing_factor: float = 1.2,
    para_spacing_pt: float = 4.0,
    margin_in: float = 0.05,
) -> float:
    """Estimate the height (inches) needed to fit `paragraphs` (list of
    (text, is_bold) tuples, one per bullet/line) into a box of
    `box_width_in` inches at `font_size_pt`. Falls back to a rough
    chars-per-line heuristic if no usable font file was found (keeps this
    module usable even on a system with neither Liberation Sans nor
    DejaVu Sans installed, just less precisely)."""
    usable_width_in = max(0.1, box_width_in - 2 * margin_in)
    line_height_in = (font_size_pt / 72.0) * line_spacing_factor

    total_lines = 0
    for text, is_bold in paragraphs:
        font = _get_font(font_size_pt, bold=is_bold)
        if font is not None:
            max_width_px = usable_width_in * 96  # approximate 96 DPI reference
            total_lines += _wrapped_line_count(text, font, max_width_px)
        else:
            chars_per_line = max(10, int(usable_width_in * 96 / (font_size_pt * 0.55)))
            total_lines += max(1, -(-len(text) // chars_per_line))  # ceil division

    height_in = total_lines * line_height_in
    height_in += len(paragraphs) * (para_spacing_pt / 72.0)
    return height_in


def fit_text_to_box(
    paragraphs: list[tuple[str, bool]],
    box_width_in: float,
    box_height_in: float,
    starting_font_size_pt: float = 12.0,
    min_font_size_pt: float = 8.0,
    font_step_pt: float = 0.5,
) -> FitResult:
    """Find the largest font size (>= min_font_size_pt) at which
    `paragraphs` fit within a box of the given width/height. If even
    `min_font_size_pt` overflows, truncate paragraphs (dropping from the
    end) and append a "...see source slide for full detail" note until
    the remainder fits, so the box never visually overflows."""
    size = starting_font_size_pt
    while size >= min_font_size_pt:
        height = estimate_required_height_in(paragraphs, box_width_in, size)
        if height <= box_height_in:
            return FitResult(font_size_pt=size, truncated=False, paragraphs=[t for t, _ in paragraphs])
        size -= font_step_pt
    size = min_font_size_pt

    # Even the smallest acceptable size overflows -- trim paragraphs from
    # the end until what remains (plus a pointer note) fits.
    working = list(paragraphs)
    note = "...see the inserted source slide for full detail"
    while working:
        trial = working + [(note, False)]
        height = estimate_required_height_in(trial, box_width_in, size)
        if height <= box_height_in:
            return FitResult(
                font_size_pt=size,
                truncated=True,
                paragraphs=[t for t, _ in working] + [note],
            )
        working = working[:-1]

    return FitResult(font_size_pt=size, truncated=True, paragraphs=[note])


if __name__ == "__main__":
    sample = [
        ("Key Achievements and Project Roadmap", True),
        ("Continued delivery of Small Lender Enablement (SLE) features and enhancements, "
         "with significant throughput growth in the latest reporting period.", False),
        ("Maintained stable delivery performance", False),
        ("The increase in reopen counts was primarily driven by process, dependency, "
         "and scope-related factors rather than implementation defects:", False),
        ("SLE-144 was reopened due to an invalid issue and not because of any implementation defect.", False),
        ("SLE-164 was reopened solely for a minor content/wording update and did not require any functional changes.", False),
        ("SLE-170 was reopened due to a scope change identified after initial completion.", False),
        ("SLE-178 was moved back to Open status because it was blocked by a dependent ticket and did not require any rework.", False),
        ("No team side blockers or operational impediments were reported", False),
        ("Risks / Constraints / Updates", True),
        ("The SLE board has transitioned from a sprint-based model to Kanban", False),
        ("Dependency related work may continue to impact ticket status changes and reopen metrics", False),
    ]
    result = fit_text_to_box(sample, box_width_in=12.33, box_height_in=2.73)
    print(f"Font size: {result.font_size_pt}pt, truncated: {result.truncated}")
    for p in result.paragraphs:
        print(" -", p[:80])
