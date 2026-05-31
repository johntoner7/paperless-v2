import logging

LOG = logging.getLogger(__name__)


def _render_page_png(page, dpi: int = 120) -> bytes:
    import fitz
    mat = fitz.Matrix(dpi / 72, dpi / 72)
    pix = page.get_pixmap(matrix=mat, alpha=False)
    return pix.tobytes("png")


def _acroform_regions(page, page_index: int) -> list:
    import fitz
    regions = []
    for widget in page.widgets():
        if widget.field_type not in (
            fitz.PDF_WIDGET_TYPE_TEXT,
            fitz.PDF_WIDGET_TYPE_CHECKBOX,
            fitz.PDF_WIDGET_TYPE_COMBOBOX,
        ):
            continue
        rect = widget.rect  # PyMuPDF already uses top-left origin
        rtype = "checkbox" if widget.field_type == fitz.PDF_WIDGET_TYPE_CHECKBOX else "text"
        regions.append({
            "id": f"p{page_index}_w{len(regions)}",
            "page": page_index,
            "type": rtype,
            "box": {
                "x": float(rect.x0),
                "y": float(rect.y0),
                "w": float(rect.width),
                "h": float(rect.height),
            },
            "value": "",
            "acroFieldName": widget.field_name,
        })
    return regions


def _heuristic_regions(page, page_index: int) -> list:
    """Detect empty fill areas in a flat PDF.

    Looks for text lines ending with ':' that have significant whitespace to
    the right — typical of "Label: ________" form patterns.
    Also detects horizontal rule annotations as field boundaries.
    """
    page_w = float(page.rect.width)
    blocks = page.get_text("dict", flags=0)["blocks"]
    regions = []
    region_idx = 0

    for block in blocks:
        if block.get("type") != 0:
            continue
        for line in block.get("lines", []):
            spans = line.get("spans", [])
            if not spans:
                continue
            line_text = "".join(s.get("text", "") for s in spans).strip()
            if not line_text.endswith(":"):
                continue

            last_span = spans[-1]
            label_right = float(last_span["bbox"][2])
            line_y0 = float(line["bbox"][1])
            line_y1 = float(line["bbox"][3])
            line_h = max(line_y1 - line_y0, 12.0)

            remaining_w = page_w - label_right - 10
            if remaining_w < page_w * 0.15:
                continue

            regions.append({
                "id": f"p{page_index}_r{region_idx}",
                "page": page_index,
                "type": "text",
                "box": {
                    "x": label_right + 4.0,
                    "y": line_y0,
                    "w": remaining_w,
                    "h": line_h,
                },
                "value": "",
                "acroFieldName": None,
            })
            region_idx += 1

    return regions


def extract_pdf_pages(pdf_path: str, dpi: int = 120) -> dict:
    """Open a PDF, detect AcroForm fields or heuristic regions, render page PNGs.

    Coordinate system: points, top-left origin, y increasing downward.
    PyMuPDF uses this convention natively, so no flip is needed.

    Returns:
        {
            "route": "acroform" | "flat",
            "pages": [
                {
                    "index": int,
                    "width_pt": float,
                    "height_pt": float,
                    "png_bytes": bytes,   # caller stores this separately
                    "regions": [...],
                }
            ]
        }
    """
    import fitz

    doc = fitz.open(pdf_path)

    has_acroform = any(list(doc[i].widgets()) for i in range(len(doc)))
    route = "acroform" if has_acroform else "flat"

    pages = []
    for i in range(len(doc)):
        page = doc[i]
        rect = page.rect
        png_bytes = _render_page_png(page, dpi=dpi)

        regions = _acroform_regions(page, i) if has_acroform else _heuristic_regions(page, i)

        pages.append({
            "index": i,
            "width_pt": float(rect.width),
            "height_pt": float(rect.height),
            "png_bytes": png_bytes,
            "regions": regions,
        })

    doc.close()
    LOG.info("extract_pdf_pages: path=%s route=%s pages=%d", pdf_path, route, len(pages))
    return {"route": route, "pages": pages}
