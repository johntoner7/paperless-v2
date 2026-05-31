import io
import logging

LOG = logging.getLogger(__name__)


def _fit_fontsize(rect, text: str, initial: float = 11.0) -> float:
    """Find the largest font size where text fits within rect width."""
    import fitz
    size = min(float(rect.height) * 0.7, initial)
    min_size = 4.0
    while size > min_size:
        if fitz.get_text_length(text, fontname="helv", fontsize=size) <= float(rect.width):
            break
        size -= 0.5
    return max(size, min_size)


def export_pdf(original_pdf_path: str, pages: list) -> bytes:
    """Stamp filled region values onto the original PDF.

    Coordinate system: points, top-left origin (PyMuPDF native) — no flip needed.

    Args:
        original_pdf_path: path to the untouched original PDF.
        pages: list of page dicts from extract_pdf_pages, with region['value'] filled.

    Returns:
        PDF bytes of the stamped document.
    """
    import fitz

    doc = fitz.open(original_pdf_path)

    for page_model in pages:
        page_idx = page_model.get("index", 0)
        if page_idx >= len(doc):
            continue
        page = doc[page_idx]

        for region in page_model.get("regions", []):
            value = region.get("value", "")
            if not value:
                continue

            box = region["box"]
            rect = fitz.Rect(
                box["x"],
                box["y"],
                box["x"] + box["w"],
                box["y"] + box["h"],
            )
            rtype = region.get("type", "text")

            # Checkbox: stamp a check mark; text: stamp the entered value
            if rtype == "checkbox":
                stamp = "☑" if value else ""
            else:
                stamp = str(value)

            if not stamp:
                continue

            acro_name = region.get("acroFieldName")
            if acro_name:
                # Route 6a: set the AcroForm widget value
                for widget in page.widgets():
                    if widget.field_name == acro_name:
                        widget.field_value = stamp
                        widget.update()
                        break
            else:
                # Route 6b: stamp free text at the region's bounding box
                fontsize = _fit_fontsize(rect, stamp)
                align = fitz.TEXT_ALIGN_CENTER if rtype == "checkbox" else fitz.TEXT_ALIGN_LEFT
                rc = page.insert_textbox(
                    rect,
                    stamp,
                    fontname="helv",
                    fontsize=fontsize,
                    align=align,
                    color=(0, 0, 0),
                )
                if rc < 0:
                    LOG.warning(
                        "insert_textbox overflow on page %d region %s (rc=%d)",
                        page_idx, region.get("id", "?"), rc,
                    )

    buf = io.BytesIO()
    doc.save(buf, deflate=True, garbage=4)
    doc.close()
    return buf.getvalue()
