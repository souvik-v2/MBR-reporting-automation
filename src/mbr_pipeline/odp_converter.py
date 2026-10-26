"""
Convert .odp (OpenDocument Presentation) files to .pptx using headless
LibreOffice, so they can be read with python-pptx like any other team
deck.

WHY THIS EXISTS
----------------
python-pptx cannot open .odp files at all -- it raises a raw `KeyError`
deep in its OPC-package loader, since .odp is a completely different
container format (ODF), not OOXML. This project's WPM team has at least
one month (sample_data/ppt-data/V2Solutions_WPM_March_2025_Report.odp)
delivered in this format, which every earlier script in this project
either ignored or would have crashed on.

LibreOffice (already available in this environment at /usr/bin/soffice)
converts .odp -> .pptx headlessly and reliably, with no manual step and
no Windows dependency. This has been verified against the real WPM file
in this project: slide titles, bullet text, and structure all carry over
correctly.
"""
from __future__ import annotations

import subprocess
from pathlib import Path


class ConversionError(RuntimeError):
    pass


def convert_odp_to_pptx(odp_path: "str | Path", out_dir: "str | Path", timeout: int = 90) -> Path:
    """Convert `odp_path` to .pptx using headless LibreOffice, writing the
    result into `out_dir` (created if needed). Returns the path to the
    converted .pptx file.

    Raises ConversionError if the `soffice` binary is unavailable, the
    conversion process fails, or (rarely, e.g. on a slow/loaded machine)
    times out without producing output.
    """
    odp_path = Path(odp_path)
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    if not odp_path.exists():
        raise ConversionError(f"Source file not found: {odp_path}")

    soffice = _find_soffice()
    if soffice is None:
        raise ConversionError(
            "LibreOffice ('soffice'/'libreoffice') was not found on this "
            "system. Install LibreOffice to enable .odp -> .pptx "
            "conversion, or supply this team's report as .pptx directly."
        )

    try:
        result = subprocess.run(
            [soffice, "--headless", "--convert-to", "pptx", "--outdir", str(out_dir), str(odp_path)],
            capture_output=True,
            text=True,
            timeout=timeout,
        )
    except subprocess.TimeoutExpired as exc:
        raise ConversionError(f"LibreOffice conversion of {odp_path.name} timed out after {timeout}s") from exc

    expected_out = out_dir / (odp_path.stem + ".pptx")
    if result.returncode != 0 or not expected_out.exists():
        raise ConversionError(
            f"LibreOffice conversion of {odp_path.name} failed "
            f"(exit code {result.returncode}).\nstdout: {result.stdout}\nstderr: {result.stderr}"
        )

    return expected_out


def _find_soffice() -> "str | None":
    for candidate in ("soffice", "libreoffice"):
        from shutil import which

        path = which(candidate)
        if path:
            return path
    return None


def ensure_pptx(deck_path: "str | Path", convert_cache_dir: "str | Path") -> Path:
    """Return a .pptx path for `deck_path`, converting it first if it's
    an .odp file. .pptx inputs are returned unchanged. Conversions are
    cached in `convert_cache_dir` keyed by source filename, so repeated
    runs don't reconvert unnecessarily."""
    deck_path = Path(deck_path)
    if deck_path.suffix.lower() == ".pptx":
        return deck_path
    if deck_path.suffix.lower() == ".odp":
        return convert_odp_to_pptx(deck_path, convert_cache_dir)
    raise ConversionError(f"Unsupported deck format: {deck_path.suffix} ({deck_path.name})")


if __name__ == "__main__":
    import sys

    src = sys.argv[1] if len(sys.argv) > 1 else "sample_data/ppt-data/V2Solutions_WPM_March_2025_Report.odp"
    out_dir = sys.argv[2] if len(sys.argv) > 2 else "output/_odp_cache"

    out = ensure_pptx(src, out_dir)
    print(f"Converted/located: {out}")

    from pptx import Presentation

    prs = Presentation(out)
    print(f"Slides: {len(prs.slides)}")
    for i, slide in enumerate(prs.slides):
        title = slide.shapes.title.text if slide.shapes.title else "(no title)"
        print(f"  {i + 1}: {title!r}")
