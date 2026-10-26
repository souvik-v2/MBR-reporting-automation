"""
Apply text_fit's sizing decision to a real PowerPoint text box, in place,
without disturbing its paragraph/bullet structure or bold formatting.

This is intentionally narrow in scope: it only ever *shrinks* font size
(within a safe floor) or, as an absolute last resort, drops trailing
bullets and appends a pointer to the full source slide. It never
rewrites or summarizes the text itself -- that's a separate, explicit
step (see key_achievements_builder.py), not something done implicitly
as a side effect of fitting text into a box.
"""
from __future__ import annotations

from pptx.util import Emu, Pt

from .text_fit import fit_text_to_box

DEFAULT_STARTING_SIZE_PT = 12.0
DEFAULT_MIN_SIZE_PT = 8.0


def fit_textbox_in_place(shape, starting_size_pt: float = DEFAULT_STARTING_SIZE_PT, min_size_pt: float = DEFAULT_MIN_SIZE_PT) -> dict:
    """Resize the font of every run in `shape`'s text frame so the
    existing paragraphs fit within the shape's current (fixed) width and
    height. Returns a small dict describing what was done, for logging.

    If even `min_size_pt` overflows, trailing paragraphs are removed
    (last-to-first, skipping the always-kept section heading paragraphs)
    and a "...see source slide" note is appended, so the box never
    visually overflows past the slide's footer line.
    """
    tf = shape.text_frame
    box_width_in = Emu(shape.width).inches
    box_height_in = Emu(shape.height).inches

    paragraphs_info = []
    for p in tf.paragraphs:
        text = "".join(r.text for r in p.runs)
        is_bold = any(r.font.bold for r in p.runs) if p.runs else False
        paragraphs_info.append((text, is_bold))

    result = fit_text_to_box(
        paragraphs_info,
        box_width_in=box_width_in,
        box_height_in=box_height_in,
        starting_font_size_pt=starting_size_pt,
        min_font_size_pt=min_size_pt,
    )

    if result.truncated:
        _truncate_paragraphs(tf, len(result.paragraphs))

    for p in tf.paragraphs:
        for r in p.runs:
            r.font.size = Pt(result.font_size_pt)

    return {
        "font_size_pt": result.font_size_pt,
        "truncated": result.truncated,
        "original_paragraph_count": len(paragraphs_info),
        "final_paragraph_count": len(tf.paragraphs),
    }


def _truncate_paragraphs(text_frame, keep_count: int) -> None:
    """Remove paragraphs from the end of `text_frame` until only
    `keep_count` remain, then append a plain note paragraph. Used only
    when even the minimum font size can't make everything fit."""
    paragraphs = list(text_frame.paragraphs)
    if keep_count >= len(paragraphs):
        return

    # Drop the last (len - keep_count + 1) paragraphs to make room for
    # the note that fit_text_to_box already counted as the final entry.
    to_drop = paragraphs[keep_count - 1:]
    note_text = None
    for p in to_drop:
        note_text = "".join(r.text for r in p.runs)
    for p in to_drop:
        p._p.getparent().remove(p._p)

    # Re-add a single note paragraph at the same indent level as a normal
    # bullet, using the text fit_text_to_box decided on.
    if note_text:
        new_p = text_frame.add_paragraph()
        new_p.level = 1
        run = new_p.add_run()
        run.text = note_text
        run.font.italic = True


if __name__ == "__main__":
    import sys
    from pathlib import Path

    sys.path.insert(0, str(Path(__file__).resolve().parent))
    from pptx import Presentation

    template_path = sys.argv[1] if len(sys.argv) > 1 else "sample_data/final-data/06-01-2026.pptx"
    out_path = sys.argv[2] if len(sys.argv) > 2 else "output/_text_fit_test.pptx"

    prs = Presentation(template_path)
    for slide in prs.slides:
        if not slide.shapes.title or "project metrics" not in slide.shapes.title.text.lower():
            continue
        for shape in slide.shapes:
            if shape.has_text_frame and "Key Achievements" in shape.text_frame.text:
                info = fit_textbox_in_place(shape)
                print(f"{slide.shapes.title.text}: {info}")

    Path(out_path).parent.mkdir(parents=True, exist_ok=True)
    prs.save(out_path)
    print(f"Saved: {out_path}")
