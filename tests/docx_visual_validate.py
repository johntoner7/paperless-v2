"""Phase 5.1 — Visual diff pipeline (optional, offline).

Renders two PDF versions of a DOCX and compares them page by page using
pixel-level image diffing. NOT part of the main conversion pipeline.

Reference options (in priority order):
  --libreoffice             LibreOffice headless DOCX→PDF  ← ground truth for CI/serverless
  --reference-pdf  <path>   pre-rendered PDF
  --reference-html <path>   pre-rendered HTML → WeasyPrint PDF
  (default)                 legacy docx_to_html → WeasyPrint PDF

Output:
  <out_dir>/candidate.pdf
  <out_dir>/reference.pdf
  <out_dir>/candidate_page_N.png
  <out_dir>/reference_page_N.png
  <out_dir>/diff_page_N.png       (amplified pixel diff, red channel)
  <out_dir>/report.json
"""

from __future__ import annotations

import argparse
import json
import os
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Optional

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

DPI = 150                   # render resolution for PDF → PNG
DIFF_THRESHOLD = 5          # pixel value (0–255) below which change is ignored
DIFF_AMPLIFY = 8            # multiply diff values so small changes are visible


# ---------------------------------------------------------------------------
# Public entry point
# ---------------------------------------------------------------------------

def find_libreoffice() -> Optional[str]:
    """Return the path to the soffice binary, or None if not found."""
    import shutil
    candidates = [
        "soffice",
        "/opt/homebrew/bin/soffice",
        "/usr/local/bin/soffice",
        "/usr/bin/soffice",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    ]
    for c in candidates:
        if shutil.which(c) or Path(c).is_file():
            return c
    return None


def run_visual_validation(
    docx_path: str | Path,
    out_dir: str | Path,
    reference_html: Optional[str | Path] = None,
    reference_pdf: Optional[str | Path] = None,
    use_libreoffice: bool = False,
    with_annotations: bool = False,
) -> dict:
    """Run the full visual diff pipeline and return the report dict."""
    docx_path = Path(docx_path)
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    # --- candidate: new AST pipeline ---
    candidate_pdf = out_dir / "candidate.pdf"
    _render_candidate(docx_path, candidate_pdf, with_annotations=with_annotations)

    # --- reference ---
    reference_pdf_path = out_dir / "reference.pdf"
    reference_method: str
    if reference_pdf:
        import shutil as _shutil
        _shutil.copy2(reference_pdf, reference_pdf_path)
        reference_method = f"pre-rendered PDF ({reference_pdf})"
    elif reference_html:
        _html_to_pdf(Path(reference_html).read_text(encoding="utf-8"), reference_pdf_path)
        reference_method = f"pre-rendered HTML ({reference_html})"
    elif use_libreoffice:
        lo = find_libreoffice()
        if not lo:
            raise RuntimeError(
                "LibreOffice not found. Install with: brew install --cask libreoffice"
            )
        _render_reference_libreoffice(docx_path, reference_pdf_path, lo)
        reference_method = f"LibreOffice headless ({lo})"
    else:
        _render_reference_legacy(docx_path, reference_pdf_path)
        reference_method = "legacy docx_to_html pipeline"

    # --- rasterise both PDFs to PNG pages ---
    cand_pages = _pdf_to_pngs(candidate_pdf, out_dir, prefix="candidate_page")
    ref_pages = _pdf_to_pngs(reference_pdf_path, out_dir, prefix="reference_page")

    # --- compare page by page ---
    page_reports = _compare_pages(cand_pages, ref_pages, out_dir)

    # --- build summary ---
    total = len(page_reports)
    changed = sum(1 for p in page_reports if p["changed_pixels_pct"] > 0)
    overall_mean = (
        round(sum(p["mean_diff"] for p in page_reports) / total, 4) if total else 0
    )

    report = {
        "docx": str(docx_path),
        "reference_method": reference_method,
        "candidate_pdf": str(candidate_pdf),
        "reference_pdf": str(reference_pdf_path),
        "dpi": DPI,
        "diff_threshold": DIFF_THRESHOLD,
        "pages": page_reports,
        "summary": {
            "total_pages_candidate": len(cand_pages),
            "total_pages_reference": len(ref_pages),
            "pages_compared": total,
            "pages_with_changes": changed,
            "overall_mean_diff": overall_mean,
            "visually_identical": changed == 0,
        },
    }

    report_path = out_dir / "report.json"
    report_path.write_text(json.dumps(report, indent=2, ensure_ascii=False), encoding="utf-8")
    return report


# ---------------------------------------------------------------------------
# Rendering helpers
# ---------------------------------------------------------------------------

def _render_candidate(docx_path: Path, out_pdf: Path, with_annotations: bool) -> None:
    from docx_ast import build_docx_ast
    from docx_renderer import render_ast_to_html
    from docx_to_html import inject_page_css

    ast = build_docx_ast(docx_path)

    annotations = None
    if with_annotations:
        try:
            from docx_annotator import generate_annotations
            annotations = generate_annotations(ast)
        except Exception as exc:
            print(f"[visual-validate] annotation generation skipped: {exc}", file=sys.stderr)

    raw_html = render_ast_to_html(ast, annotations)
    html_doc = inject_page_css(raw_html, ast.get("page_setup"))
    _html_to_pdf(html_doc, out_pdf)


def _render_reference_legacy(docx_path: Path, out_pdf: Path) -> None:
    from docx_to_html import convert_docx_to_a4_html
    html_doc = convert_docx_to_a4_html(str(docx_path))
    _html_to_pdf(html_doc, out_pdf)


def _render_reference_libreoffice(docx_path: Path, out_pdf: Path, soffice: str) -> None:
    """Convert DOCX → PDF using LibreOffice headless — the ground-truth reference.

    LibreOffice writes the PDF alongside the source file, so we use a temp
    directory to avoid polluting the source tree.
    """
    import shutil
    import tempfile

    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)
        # Copy DOCX into temp dir so LO writes the PDF there
        tmp_docx = tmp_path / docx_path.name
        shutil.copy2(docx_path, tmp_docx)

        result = subprocess.run(
            [
                soffice,
                "--headless",
                "--convert-to", "pdf",
                "--outdir", str(tmp_path),
                str(tmp_docx),
            ],
            capture_output=True,
            text=True,
            timeout=60,
        )
        if result.returncode != 0:
            raise RuntimeError(f"LibreOffice conversion failed: {result.stderr.strip()}")

        generated = tmp_path / (docx_path.stem + ".pdf")
        if not generated.exists():
            raise RuntimeError(f"LibreOffice did not produce {generated}")

        shutil.copy2(generated, out_pdf)


def _html_to_pdf(html_doc: str, out_pdf: Path) -> None:
    from weasyprint import HTML
    HTML(string=html_doc).write_pdf(str(out_pdf), presentational_hints=True)


# ---------------------------------------------------------------------------
# PDF → PNG rasterisation
# ---------------------------------------------------------------------------

def _pdf_to_pngs(pdf_path: Path, out_dir: Path, prefix: str) -> list[Path]:
    """Rasterise every page of *pdf_path* to PNG files using pdftoppm.

    Returns a sorted list of output PNG paths.
    """
    out_prefix = str(out_dir / prefix)
    result = subprocess.run(
        ["pdftoppm", "-r", str(DPI), "-png", str(pdf_path), out_prefix],
        capture_output=True,
        text=True,
    )
    if result.returncode != 0:
        raise RuntimeError(f"pdftoppm failed: {result.stderr.strip()}")

    pages = sorted(out_dir.glob(f"{prefix}-*.png"))
    return pages


# ---------------------------------------------------------------------------
# Pixel-level image comparison
# ---------------------------------------------------------------------------

def _compare_pages(
    cand_pages: list[Path],
    ref_pages: list[Path],
    out_dir: Path,
) -> list[dict]:
    from PIL import Image, ImageChops

    n = max(len(cand_pages), len(ref_pages))
    reports: list[dict] = []

    for i in range(n):
        page_num = i + 1

        if i >= len(cand_pages):
            reports.append(_missing_page_report(page_num, side="candidate"))
            continue
        if i >= len(ref_pages):
            reports.append(_missing_page_report(page_num, side="reference"))
            continue

        cand_img = Image.open(cand_pages[i]).convert("RGB")
        ref_img = Image.open(ref_pages[i]).convert("RGB")

        # Resize candidate to reference dimensions if they differ
        if cand_img.size != ref_img.size:
            cand_img = cand_img.resize(ref_img.size, Image.LANCZOS)

        diff_img = ImageChops.difference(cand_img, ref_img)

        # Save amplified diff (red channel only, for clarity)
        diff_path = out_dir / f"diff_page_{page_num}.png"
        _save_amplified_diff(diff_img, diff_path)

        # Compute metrics
        mean_diff, max_diff, changed_pct = _diff_metrics(diff_img)

        reports.append({
            "page": page_num,
            "width_px": ref_img.width,
            "height_px": ref_img.height,
            "mean_diff": round(mean_diff, 4),
            "max_diff": int(max_diff),
            "changed_pixels_pct": round(changed_pct, 2),
            "candidate_png": str(cand_pages[i]),
            "reference_png": str(ref_pages[i]),
            "diff_png": str(diff_path),
        })

    return reports


def _diff_metrics(diff_img) -> tuple[float, float, float]:
    """Return (mean_diff, max_diff, changed_pixels_pct) for a diff image."""
    import struct

    w, h = diff_img.size
    total_pixels = w * h
    if total_pixels == 0:
        return 0.0, 0.0, 0.0

    # tobytes() returns raw unsigned bytes — stable across Pillow versions
    raw = diff_img.tobytes()                  # length = w * h * 3 (RGB)
    channel_values = raw                      # bytes object supports iteration
    mean_diff = sum(channel_values) / len(channel_values)
    max_diff = max(channel_values)

    # Pixels where ANY channel exceeds the threshold (stride = 3 for RGB)
    changed = sum(
        1 for i in range(0, len(raw), 3)
        if raw[i] > DIFF_THRESHOLD or raw[i+1] > DIFF_THRESHOLD or raw[i+2] > DIFF_THRESHOLD
    )
    changed_pct = changed / total_pixels * 100

    return mean_diff, max_diff, changed_pct


def _save_amplified_diff(diff_img, path: Path) -> None:
    """Save diff image with differences amplified and false-coloured red."""
    from PIL import Image

    r, g, b = diff_img.split()
    # Combine channels into a single intensity, then amplify in the red channel
    intensity = Image.merge("RGB", (r, g, b)).convert("L")
    amplified = intensity.point(lambda v: min(255, v * DIFF_AMPLIFY))
    black = Image.new("L", intensity.size, 0)
    visual = Image.merge("RGB", (amplified, black, black))
    visual.save(str(path))


def _missing_page_report(page_num: int, side: str) -> dict:
    return {
        "page": page_num,
        "error": f"Page missing from {side}",
        "mean_diff": 1.0,
        "max_diff": 255,
        "changed_pixels_pct": 100.0,
    }


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def _cli(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Visual diff: new AST pipeline vs reference rendering"
    )
    parser.add_argument("docx", help="Input DOCX file")
    parser.add_argument(
        "--out-dir", "-o", default="./visual_diff_out",
        help="Directory for PDFs, PNGs, and report (default: ./visual_diff_out)"
    )
    parser.add_argument(
        "--reference-html", metavar="PATH",
        help="Use this HTML file as the reference instead of the legacy pipeline"
    )
    parser.add_argument(
        "--reference-pdf", metavar="PATH",
        help="Use this PDF file as the reference"
    )
    parser.add_argument(
        "--libreoffice", action="store_true",
        help="Use LibreOffice headless as the ground-truth reference (recommended for CI)"
    )
    parser.add_argument(
        "--with-annotations", action="store_true",
        help="Run AI annotation generation for the candidate render (requires ANTHROPIC_API_KEY)"
    )
    args = parser.parse_args(argv)

    try:
        report = run_visual_validation(
            docx_path=args.docx,
            out_dir=args.out_dir,
            reference_html=args.reference_html,
            reference_pdf=args.reference_pdf,
            use_libreoffice=args.libreoffice,
            with_annotations=args.with_annotations,
        )
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1

    # Print human-readable summary
    s = report["summary"]
    print(f"\nVisual diff report — {args.docx}")
    print(f"  Reference:           {report['reference_method']}")
    print(f"  Pages compared:      {s['pages_compared']}")
    print(f"  Pages with changes:  {s['pages_with_changes']}")
    print(f"  Overall mean diff:   {s['overall_mean_diff']:.4f} / 255")
    print(f"  Visually identical:  {s['visually_identical']}")
    print()
    for p in report["pages"]:
        if p.get("error"):
            print(f"  Page {p['page']}: {p['error']}")
        else:
            print(
                f"  Page {p['page']}: mean={p['mean_diff']:.4f}  "
                f"max={p['max_diff']}  changed={p['changed_pixels_pct']:.2f}%  "
                f"→ {p['diff_png']}"
            )
    print(f"\n  Full report: {args.out_dir}/report.json")
    return 0


if __name__ == "__main__":
    raise SystemExit(_cli())
