"""
Export the finished MBR PowerPoint deck to PDF, using headless
LibreOffice (the same `soffice` binary already required for .odp
conversion -- see odp_converter.py).

This is optional: if LibreOffice isn't installed, PDF export is skipped
with a clear message rather than failing the whole build, since the
.pptx output is the primary deliverable and PDF is a convenience export.
"""
from __future__ import annotations

import subprocess
from pathlib import Path
from shutil import which


class PdfExportError(RuntimeError):
    pass


def _find_soffice() -> "str | None":
    for candidate in ("soffice", "libreoffice"):
        path = which(candidate)
        if path:
            return path
    return None


def export_pptx_to_pdf(pptx_path: "str | Path", out_dir: "str | Path | None" = None, timeout: int = 120) -> Path:
    """Convert `pptx_path` to a PDF in the same folder (or `out_dir` if
    given). Returns the path to the generated PDF.

    Raises PdfExportError if LibreOffice isn't installed or the
    conversion fails/times out -- callers that want PDF export to be a
    soft failure (warn, don't crash the whole build) should catch this.
    """
    pptx_path = Path(pptx_path)
    if not pptx_path.exists():
        raise PdfExportError(f"Source file not found: {pptx_path}")

    out_dir = Path(out_dir) if out_dir else pptx_path.parent
    out_dir.mkdir(parents=True, exist_ok=True)

    soffice = _find_soffice()
    if soffice is None:
        raise PdfExportError(
            "LibreOffice ('soffice'/'libreoffice') was not found on this system. "
            "Install LibreOffice to enable PDF export (see requirements.txt)."
        )

    try:
        result = subprocess.run(
            [soffice, "--headless", "--convert-to", "pdf", "--outdir", str(out_dir), str(pptx_path)],
            capture_output=True,
            text=True,
            timeout=timeout,
        )
    except subprocess.TimeoutExpired as exc:
        raise PdfExportError(f"LibreOffice PDF export of {pptx_path.name} timed out after {timeout}s") from exc

    expected_out = out_dir / (pptx_path.stem + ".pdf")
    if result.returncode != 0 or not expected_out.exists():
        raise PdfExportError(
            f"LibreOffice PDF export of {pptx_path.name} failed "
            f"(exit code {result.returncode}).\nstdout: {result.stdout}\nstderr: {result.stderr}"
        )

    return expected_out


if __name__ == "__main__":
    import sys

    src = sys.argv[1] if len(sys.argv) > 1 else "output/MBR_March_2026.pptx"
    try:
        out = export_pptx_to_pdf(src)
        print(f"PDF exported: {out}")
    except PdfExportError as exc:
        print(f"PDF export skipped: {exc}")
