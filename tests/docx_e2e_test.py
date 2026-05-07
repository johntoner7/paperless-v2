"""Phase 6.1 — End-to-end pipeline test.

Runs the full pipeline:
    DOCX → AST → HTML → (annotations) → final HTML

Checks:
  1. Content loss     — structural counts match between AST and rendered HTML
  2. Stability        — two runs produce byte-identical HTML (determinism)
  3. Visual fidelity  — pixel diff vs legacy pipeline
  4. Annotation round-trip — annotations apply without breaking validation

Exits 0 if all checks pass, 1 if any fail.
"""

from __future__ import annotations

import hashlib
import json
import sys
import time
from pathlib import Path


def run(docx_path: str, out_dir: str = "./e2e_out", api_key_available: bool = True) -> dict:
    docx = Path(docx_path)
    out = Path(out_dir)
    out.mkdir(parents=True, exist_ok=True)

    report: dict = {
        "docx": str(docx),
        "checks": {},
        "passed": [],
        "failed": [],
    }

    # ------------------------------------------------------------------
    # Step 1 — Build AST
    # ------------------------------------------------------------------
    t0 = time.perf_counter()
    from docx_ast import build_docx_ast
    ast = build_docx_ast(docx)
    ast_ms = round((time.perf_counter() - t0) * 1000)

    sections = [s["kind"] for s in ast["sections"]]
    body = next(s for s in ast["sections"] if s["kind"] == "body")
    top_blocks = len(body["blocks"])

    report["ast"] = {
        "sections": sections,
        "body_blocks": top_blocks,
        "relationships": len(ast["relationships"]),
        "media_files": len(ast["images"]),
        "build_ms": ast_ms,
    }
    _pass(report, "ast_build", f"AST built in {ast_ms}ms — {top_blocks} body blocks, "
          f"{len(ast['images'])} media file(s)")

    # ------------------------------------------------------------------
    # Step 2 — Render without annotations (baseline)
    # ------------------------------------------------------------------
    from docx_renderer import render_ast_to_html, validate_ast_html
    from docx_to_html import inject_page_css

    t0 = time.perf_counter()
    html_bare = render_ast_to_html(ast)
    render_ms = round((time.perf_counter() - t0) * 1000)

    (out / "output_bare.html").write_text(
        inject_page_css(html_bare, ast.get("page_setup")), encoding="utf-8"
    )

    # ------------------------------------------------------------------
    # Step 3 — Content loss check (structural validation)
    # ------------------------------------------------------------------
    val = validate_ast_html(ast, html_bare)
    report["checks"]["content_loss"] = val

    if val["valid"]:
        counts = val["matches"]
        _pass(report, "content_loss",
              f"No content loss — paragraphs:{counts['paragraphs']}  "
              f"tables:{counts['tables']}  images:{counts['images']}")
    else:
        _fail(report, "content_loss",
              f"Content loss detected: {val['mismatches']}")

    # ------------------------------------------------------------------
    # Step 4 — Stability (determinism): render twice, compare hashes
    # ------------------------------------------------------------------
    html_run2 = render_ast_to_html(ast)
    h1 = hashlib.sha256(html_bare.encode()).hexdigest()
    h2 = hashlib.sha256(html_run2.encode()).hexdigest()

    if h1 == h2:
        _pass(report, "stability", f"Output is deterministic (sha256={h1[:16]}…)")
    else:
        _fail(report, "stability", "Two renders produced different HTML — non-deterministic output")

    # ------------------------------------------------------------------
    # Step 5 — Annotation round-trip
    # ------------------------------------------------------------------
    annotations: dict | None = None

    if api_key_available:
        try:
            from docx_annotator import generate_annotations
            t0 = time.perf_counter()
            annotations = generate_annotations(ast)
            ann_ms = round((time.perf_counter() - t0) * 1000)

            ann_count = len(annotations.get("annotations", {}))
            doc_type = annotations.get("document_type", "unknown")
            report["annotations"] = {
                "document_type": doc_type,
                "annotation_count": ann_count,
                "generate_ms": ann_ms,
            }

            html_annotated = render_ast_to_html(ast, annotations)
            val_ann = validate_ast_html(ast, html_annotated)

            if val_ann["valid"]:
                _pass(report, "annotation_roundtrip",
                      f"Annotations applied cleanly — {ann_count} blocks annotated, "
                      f"document_type={doc_type!r}, render still valid")
            else:
                _fail(report, "annotation_roundtrip",
                      f"Structural mismatch after annotations: {val_ann['mismatches']}")

            (out / "output_annotated.html").write_text(
                inject_page_css(html_annotated, ast.get("page_setup")), encoding="utf-8"
            )
        except Exception as exc:
            _fail(report, "annotation_roundtrip", f"Annotation generation failed: {exc}")
    else:
        report["checks"]["annotation_roundtrip"] = "skipped (no API key)"
        report["passed"].append("annotation_roundtrip (skipped)")

    # ------------------------------------------------------------------
    # Step 6 — Visual fidelity (pixel diff vs legacy pipeline)
    # ------------------------------------------------------------------
    try:
        from docx_visual_validate import run_visual_validation
        vr = run_visual_validation(
            docx_path=docx,
            out_dir=out / "visual_diff",
            with_annotations=False,
        )
        s = vr["summary"]
        report["visual_diff"] = s

        mean = s["overall_mean_diff"]
        changed = s["pages_with_changes"]
        total = s["pages_compared"]

        # Fidelity is "good" if mean pixel diff is under 15/255 (~6% of scale)
        # This tolerates font hinting and minor layout shifts vs legacy renderer
        if mean < 15:
            _pass(report, "visual_fidelity",
                  f"Mean pixel diff {mean:.2f}/255 across {total} page(s) "
                  f"({changed} page(s) changed) — within acceptable tolerance")
        else:
            _fail(report, "visual_fidelity",
                  f"Mean pixel diff {mean:.2f}/255 exceeds threshold (15/255)")
    except Exception as exc:
        _fail(report, "visual_fidelity", f"Visual diff failed: {exc}")

    # ------------------------------------------------------------------
    # Summary
    # ------------------------------------------------------------------
    report["render_ms"] = render_ms
    report["result"] = "PASS" if not report["failed"] else "FAIL"

    report_path = out / "e2e_report.json"
    report_path.write_text(json.dumps(report, indent=2, ensure_ascii=False), encoding="utf-8")

    return report


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _pass(report: dict, name: str, msg: str) -> None:
    report["passed"].append(name)
    report["checks"][name] = {"result": "PASS", "detail": msg}


def _fail(report: dict, name: str, msg: str) -> None:
    report["failed"].append(name)
    report["checks"][name] = {"result": "FAIL", "detail": msg}


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    import os
    docx = sys.argv[1] if len(sys.argv) > 1 else "invoice.docx"
    out_dir = sys.argv[2] if len(sys.argv) > 2 else "./e2e_out"
    api_available = bool(os.environ.get("ANTHROPIC_API_KEY"))

    report = run(docx, out_dir, api_key_available=api_available)

    # Print human-readable summary
    print("\n" + "=" * 60)
    print(f"  End-to-End Test — {docx}")
    print("=" * 60)
    for name, check in report["checks"].items():
        if isinstance(check, dict):
            icon = "✓" if check["result"] == "PASS" else "✗"
            print(f"  {icon}  {name}")
            print(f"       {check['detail']}")
        else:
            print(f"  -  {name}: {check}")

    print()
    if report.get("ast"):
        a = report["ast"]
        print(f"  AST: {a['sections']} | {a['body_blocks']} blocks | "
              f"{a['media_files']} image(s) | built in {a['build_ms']}ms")
    if report.get("annotations"):
        an = report["annotations"]
        print(f"  Ann: {an['annotation_count']} annotations | "
              f"type={an['document_type']!r} | generated in {an['generate_ms']}ms")
    if report.get("visual_diff"):
        vd = report["visual_diff"]
        print(f"  Vis: mean_diff={vd['overall_mean_diff']:.2f}/255 | "
              f"{vd['pages_with_changes']}/{vd['pages_compared']} pages changed")

    print()
    print(f"  Result: {report['result']}  "
          f"({len(report['passed'])} passed, {len(report['failed'])} failed)")
    print(f"  Report: {out_dir}/e2e_report.json")
    print("=" * 60 + "\n")

    sys.exit(0 if report["result"] == "PASS" else 1)
