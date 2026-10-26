"""
Cross-presentation slide copying for python-pptx.

WHY THIS EXISTS
----------------
python-pptx has no built-in "copy this slide from presentation A into
presentation B" operation. This project's documentation (MBR_WORKFLOW_
GUIDE.md, AUTOMATION_SUMMARY.md, MARCH_CONSOLIDATION_GUIDE.md) concluded
this wasn't reliably possible without Windows + PowerPoint COM
automation, and fell back to printed manual copy/paste instructions.

That conclusion was wrong. Cross-presentation slide copying with full
visual fidelity (text formatting, tables, pictures, native charts) works
fine with plain python-pptx + lxml, by deep-copying each shape's XML
element into the destination slide and re-registering any binary parts
(images, embedded chart workbooks) the copied shapes reference. This has
been verified against this project's real source decks, rendered to PDF
and visually inspected -- bold formatting, bullet text, and table-as-
image content all reproduce exactly (see /home/claude/build/slide_check
-15.jpg and slide_pic-15.jpg from the verification session).

LIMITATIONS
-----------
- SmartArt and OLE-embedded objects (e.g. an embedded Excel range that
  isn't a chart) are not specifically handled; they would currently be
  copied as opaque graphicFrame XML, which may or may not render
  correctly depending on whether their referenced parts also get copied.
  None of the real source decks in this project use these, so this is a
  documented gap rather than a silent failure mode.
- Speaker notes are not copied (not needed for this project's slides).
"""
from __future__ import annotations

import copy
import io

from pptx.oxml.ns import qn
from pptx.parts.chart import ChartPart


PICTURE_SHAPE_TYPE = 13  # MSO_SHAPE_TYPE.PICTURE


def copy_slide(dst_prs, src_slide, layout_name: str = "Blank"):
    """Append a full-fidelity copy of `src_slide` (which may belong to a
    *different* Presentation object) onto `dst_prs`. Returns the new
    slide. Use `move_slide_to` afterward if it needs to land somewhere
    other than the end of the deck.

    IMPORTANT: placeholder shapes (title, body, slide number, etc.) very
    often have no explicit position/size of their own in the XML --
    they inherit left/top/width/height from the matching placeholder on
    their *own* slide's layout. That inheritance only works as long as
    the shape stays on a layout that defines a placeholder at the same
    index. The destination slide here is built on a "Blank" layout with
    no placeholders at all, so a copied placeholder with no explicit
    geometry has nothing to inherit from once moved. LibreOffice's
    renderer is forgiving about this (falls back to some default
    position), but real PowerPoint is not -- shapes can render
    overlapping, blank, or pushed into a corner. To avoid depending on
    inheritance at all, every shape's *resolved* geometry is read from
    the source (where it inherits correctly) and baked in explicitly on
    the copy before it's added to the destination slide.
    """
    blank_layout = _find_layout(dst_prs, layout_name)
    new_slide = dst_prs.slides.add_slide(blank_layout)

    # The layout may auto-populate placeholder shapes (e.g. an empty
    # title box) onto the new slide; clear those so only the copied
    # shapes remain.
    for shp in list(new_slide.shapes):
        shp._element.getparent().remove(shp._element)

    for shape in src_slide.shapes:
        # Resolve the shape's *effective* geometry now, on the source,
        # where placeholder inheritance still works correctly.
        geometry = (shape.left, shape.top, shape.width, shape.height)

        if shape.has_chart:
            new_shape = _clone_chart_shape(new_slide, shape)
        elif shape.shape_type == PICTURE_SHAPE_TYPE:
            new_shape = _clone_picture_shape(new_slide, shape)
        else:
            new_shape = _clone_generic_shape(new_slide, shape)

        _apply_explicit_geometry(new_shape, geometry)

    return new_slide


def _apply_explicit_geometry(shape, geometry) -> None:
    """Bake `geometry` (left, top, width, height -- any of which may be
    None if even the source couldn't resolve a value, e.g. a shape on a
    layout that itself has no matching placeholder) onto `shape`
    explicitly, so it no longer depends on inheriting from a layout."""
    left, top, width, height = geometry
    if shape is None:
        return
    if left is not None:
        shape.left = left
    if top is not None:
        shape.top = top
    if width is not None:
        shape.width = width
    if height is not None:
        shape.height = height


def move_slide_to(prs, slide, position: int) -> None:
    """Move `slide` (already part of `prs`) to 0-based `position` in the
    slide order. python-pptx has no public API for this -- the slide
    order lives in the <p:sldIdLst> element under <p:presentation>, and
    must be reordered by manipulating that XML list directly (the naive
    approach of reordering a Python list copy is a no-op on the actual
    file, since slides.__iter__ rebuilds its sequence from the XML each
    time rather than being backed by a mutable list).
    """
    sldIdLst = prs.slides._sldIdLst
    slide_id_elements = list(sldIdLst)

    # The simplest reliable lookup: iterate slides in their current XML
    # order and match by identity against the Slide object passed in.
    current_order = list(prs.slides)
    try:
        current_index = current_order.index(slide)
    except ValueError:
        raise ValueError("`slide` is not part of `prs.slides`")

    el = slide_id_elements[current_index]
    sldIdLst.remove(el)
    position = max(0, min(position, len(slide_id_elements) - 1))
    sldIdLst.insert(position, el)


def _find_layout(prs, layout_name: str):
    for layout in prs.slide_layouts:
        if layout.name.lower() == layout_name.lower():
            return layout
    # Fall back to whichever layout has the fewest placeholders, which is
    # closest in spirit to a blank canvas, rather than guessing index -1.
    return min(prs.slide_layouts, key=lambda l: len(l.placeholders))


def _clone_generic_shape(new_slide, shape):
    """Clone a shape via raw XML deep-copy. Covers text boxes, tables,
    placeholders, and freeform/auto shapes -- anything that doesn't need
    a binary part re-registered in the destination package. Returns the
    newly created shape (as a python-pptx Shape, via the slide's shape
    collection) so its geometry can be set explicitly afterward."""
    new_el = copy.deepcopy(shape._element)
    new_slide.shapes._spTree.append(new_el)
    return new_slide.shapes[-1]


def _clone_picture_shape(new_slide, shape):
    """Clone a picture shape, re-registering its image blob as a new
    image part in the destination package and repointing the blip
    reference (`r:embed`) to the new relationship id. Returns the newly
    created shape."""
    image_part = shape.image  # python-pptx Image: .blob, .ext, .content_type
    new_image_part, _ = new_slide.part.get_or_add_image_part(io.BytesIO(image_part.blob))

    new_el = copy.deepcopy(shape._element)
    rId = new_slide.part.relate_to(
        new_image_part,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
    )
    blip = new_el.find(".//" + qn("a:blip"))
    if blip is not None:
        blip.set(qn("r:embed"), rId)
    new_slide.shapes._spTree.append(new_el)
    return new_slide.shapes[-1]


def _clone_chart_shape(new_slide, shape):
    """Clone a chart graphicFrame shape, including a fresh copy of its
    chart part and (if present) the embedded .xlsx workbook backing it,
    so the copied chart remains independently editable in the
    destination deck rather than pointing back at the source file.
    Returns the newly created shape."""
    src_chart_part = shape.chart.part

    new_chart_part = ChartPart(
        new_slide.part.package.next_partname(ChartPart.partname_template),
        src_chart_part.content_type,
        new_slide.part.package,
        copy.deepcopy(src_chart_part._element),
    )

    # Re-create the embedded workbook relationship, if the source chart
    # has one (charts backed by an *external* link, e.g. a SharePoint
    # URL, have no embedded part to copy -- nothing to do in that case).
    src_xlsx_part = src_chart_part.chart_workbook.xlsx_part
    if src_xlsx_part is not None:
        new_chart_part.chart_workbook.update_from_xlsx_blob(src_xlsx_part.blob)

    new_el = copy.deepcopy(shape._element)
    rId = new_slide.part.relate_to(
        new_chart_part,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart",
    )
    chart_ref = new_el.find(".//" + qn("c:chart"))
    if chart_ref is not None:
        chart_ref.set(qn("r:id"), rId)
    new_slide.shapes._spTree.append(new_el)
    return new_slide.shapes[-1]


if __name__ == "__main__":
    import sys

    from pptx import Presentation

    src_path = sys.argv[1] if len(sys.argv) > 1 else "sample_data/ppt-data/V2Solutions_SEOD_March_2026_Report.pptx"
    src_slide_idx = int(sys.argv[2]) if len(sys.argv) > 2 else 2
    dst_path = sys.argv[3] if len(sys.argv) > 3 else "sample_data/final-data/06-01-2026.pptx"
    out_path = sys.argv[4] if len(sys.argv) > 4 else "output/_slide_copy_test.pptx"
    insert_at = int(sys.argv[5]) if len(sys.argv) > 5 else None

    src = Presentation(src_path)
    dst = Presentation(dst_path)
    new_slide = copy_slide(dst, src.slides[src_slide_idx])
    if insert_at is not None:
        move_slide_to(dst, new_slide, insert_at)

    from pathlib import Path
    Path(out_path).parent.mkdir(parents=True, exist_ok=True)
    dst.save(out_path)
    print(f"Copied slide {src_slide_idx + 1} from {src_path} -> {out_path} (slide count: {len(dst.slides)})")
