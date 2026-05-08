"""Pure AST → HTML renderer and structural validator.

Phase 2.1  render_ast_to_html(ast)   — structural correctness, no styling.
Phase 2.2  validate_ast_html(ast, html) — structural count validation.
Phase 3.1  Inline CSS for font, alignment, and spacing.
Phase 3.2  Table layout fidelity: cell alignment, border-collapse, widths.
Phase 3.3  Image rendering: max-width, dimensions, inline vs anchor positioning.
Phase 4.3  Optional annotations overlay: AST + annotations → enhanced HTML.
"""

from __future__ import annotations

import html
import re
from typing import Optional


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def render_ast_to_html(ast: dict, annotations: Optional[dict] = None) -> str:
    """Convert a DocxAST dict to a complete HTML document.

    Pure function — same input always produces same output.
    Pass an annotations dict (from docx_annotator.generate_annotations) to
    apply layout hints. Renderer works identically without annotations.
    """
    ann_blocks: dict = (annotations or {}).get("annotations", {})
    parts: list[str] = []
    for si, section in enumerate(ast.get("sections", [])):
        for bi, block in enumerate(section.get("blocks", [])):
            key = f"s{si}:b{bi}"
            block_ann = ann_blocks.get(key, {})
            rendered = _render_block(block, block_ann)
            if rendered:
                parts.append(rendered)
    body = "\n".join(parts)
    return _wrap_html(f'<div class="document">\n{body}\n</div>')


def validate_ast_html(ast: dict, rendered_html: str) -> dict:
    """Validate rendered HTML structural counts against the AST source of truth.

    Returns a report with 'valid', 'ast', 'html', 'matches', 'mismatches'.
    """
    ast_counts = _count_ast_nodes(ast)
    html_counts = _count_html_nodes(rendered_html)

    matches: dict = {}
    mismatches: dict = {}
    for key, ast_val in ast_counts.items():
        html_val = html_counts.get(key, 0)
        if ast_val == html_val:
            matches[key] = ast_val
        else:
            mismatches[key] = {"ast": ast_val, "html": html_val}

    return {
        "valid": len(mismatches) == 0,
        "ast": ast_counts,
        "html": html_counts,
        "matches": matches,
        "mismatches": mismatches,
    }


# ---------------------------------------------------------------------------
# Block rendering
# ---------------------------------------------------------------------------

def _render_block(block: dict, ann: Optional[dict] = None) -> str:
    ann = ann or {}
    kind = block.get("type")
    if kind == "paragraph":
        return _render_paragraph(block, ann)
    if kind == "table":
        return _render_table(block, ann)
    return ""


def _render_paragraph(node: dict, ann: Optional[dict] = None) -> str:
    ann = ann or {}
    style = node.get("style", {})

    # 4.3 — annotation may override the heading tag
    hw = ann.get("heading_weight")
    tag = f"h{hw}" if hw else _heading_tag(style)

    # When a right-anchored image shares a paragraph with text, float: right
    # places the image correctly but the text sits at the top of the float's
    # line box, causing vertical misalignment. Flexbox with row-reverse fixes
    # both axes: image (first in source) ends up at the right, text vertically
    # centred to its left. Left-anchored images (e.g. header logo) keep their
    # original float behaviour — switching them to flex makes the paragraph as
    # tall as the image, which can push content onto an extra page.
    use_flex = (
        _para_has_text(node)
        and _para_anchor_on_right(node)
    )

    inner = "".join(_render_inline(c, in_flex_para=use_flex) for c in node.get("children", []))

    ann_css = _annotation_block_css(ann)

    if use_flex:
        # Use order:2 on the image (set in _render_image) to push it to the
        # right without reversing text spans. justify-content:flex-end aligns
        # the whole group at the right edge.
        flex_css = "display: flex; align-items: center; justify-content: flex-end"
        css = "; ".join(c for c in [flex_css, ann_css] if c)
    else:
        base_css = _paragraph_css(style)
        css = "; ".join(c for c in [base_css, ann_css] if c)

    style_attr = f' style="{css}"' if css else ""

    if not inner.strip():
        return f'<{tag} class="docx-empty-paragraph"{style_attr}></{tag}>'
    return f"<{tag}{style_attr}>{inner}</{tag}>"


def _para_has_text(node: dict) -> bool:
    for run in node.get("children", []):
        for item in run.get("children", []):
            if item.get("type") == "text" and (item.get("text") or "").strip():
                return True
    return False


def _para_anchor_on_right(node: dict) -> bool:
    for run in node.get("children", []):
        for item in run.get("children", []):
            if item.get("type") == "image":
                pos = item.get("positioning") or {}
                if pos.get("kind") == "anchor" and pos.get("offset_h_emu", 0) >= 3_200_000:
                    return True
    return False


def _heading_tag(style: dict) -> str:
    name = style.get("style_name", "")
    if name:
        m = re.search(r"heading\s*([1-6])", name.lower())
        if m:
            return f"h{m.group(1)}"
    return "p"


# 3.1 — paragraph CSS builder
def _paragraph_css(style: dict) -> str:
    parts: list[str] = []

    align = style.get("alignment")
    if align:
        parts.append(f"text-align: {_map_alignment(align)}")

    spacing = style.get("spacing") or {}
    before = spacing.get("before")
    after = spacing.get("after")
    line = spacing.get("line")
    if before is not None:
        px = _twips_to_px(before)
        if px is not None:
            parts.append(f"margin-top: {px}px")
    if after is not None:
        px = _twips_to_px(after)
        if px is not None:
            parts.append(f"margin-bottom: {px}px")
    if line is not None:
        try:
            lh = round(int(line) / 240, 2)
            if lh > 0:
                parts.append(f"line-height: {lh}")
        except (ValueError, TypeError):
            pass

    indent = style.get("indent") or {}
    left = indent.get("left")
    right = indent.get("right")
    first_line = indent.get("firstLine")
    hanging = indent.get("hanging")

    if left is not None:
        px = _twips_to_px(left)
        if px is not None:
            parts.append(f"padding-left: {px}px")
    if right is not None:
        px = _twips_to_px(right)
        if px is not None:
            parts.append(f"padding-right: {px}px")

    if first_line is not None:
        px = _twips_to_px(first_line)
        if px is not None:
            parts.append(f"text-indent: {px}px")
    elif hanging is not None:
        # Hanging indent: negative first-line + matching padding compensation
        px = _twips_to_px(hanging)
        if px is not None:
            parts.append(f"text-indent: -{px}px")
            if left is None:
                parts.append(f"padding-left: {px}px")

    return "; ".join(parts)


# ---------------------------------------------------------------------------
# Inline rendering
# ---------------------------------------------------------------------------

def _render_inline(node: dict, in_flex_para: bool = False) -> str:
    kind = node.get("type")
    if kind == "run":
        return _render_run(node, in_flex_para=in_flex_para)
    if kind == "hyperlink":
        return _render_hyperlink(node)
    return ""


def _render_run(node: dict, in_flex_para: bool = False) -> str:
    inner = "".join(_render_run_content(c, in_flex_para=in_flex_para) for c in node.get("children", []))
    if not inner:
        return ""
    style = node.get("style", {})

    # 3.1 — superscript / subscript (innermost — closest to text)
    vert = style.get("vertical_align")
    if vert == "superscript":
        inner = f"<sup>{inner}</sup>"
    elif vert == "subscript":
        inner = f"<sub>{inner}</sub>"

    # Semantic text decorations
    if style.get("underline"):
        inner = f"<u>{inner}</u>"
    if style.get("italic"):
        inner = f"<em>{inner}</em>"
    if style.get("bold"):
        inner = f"<strong>{inner}</strong>"

    # 3.1 — font CSS span (outermost — applies to everything including semantic tags)
    css = _run_font_css(style)
    if css:
        inner = f'<span style="{css}">{inner}</span>'

    return inner


# 3.1 — run font CSS builder (font metrics + color + strike only; bold/italic/underline via tags)
def _run_font_css(style: dict) -> str:
    parts: list[str] = []
    fs = style.get("font_size_pt")
    if fs is not None:
        parts.append(f"font-size: {fs}pt")
    ff = style.get("font_family")
    if ff:
        parts.append(f"font-family: '{ff}'")
    color = style.get("color")
    if color:
        c = color if color.startswith("#") else f"#{color}"
        parts.append(f"color: {c}")
    if style.get("strike"):
        parts.append("text-decoration: line-through")
    return "; ".join(parts)


def _render_run_content(node: dict, in_flex_para: bool = False) -> str:
    kind = node.get("type")
    if kind == "text":
        text = html.escape(node.get("text", ""), quote=False)
        if node.get("preserve_spaces"):
            text = text.replace(" ", " ")
        return text
    if kind == "break":
        return "&emsp;" if node.get("break_type") == "tab" else "<br>"
    if kind == "image":
        return _render_image(node, in_flex_para=in_flex_para)
    return ""


def _render_hyperlink(node: dict) -> str:
    inner = "".join(_render_run(r) for r in node.get("children", []))
    if not inner:
        return ""
    href = node.get("href")
    if href:
        return f'<a href="{html.escape(href, quote=True)}">{inner}</a>'
    return inner


# ---------------------------------------------------------------------------
# 3.3 — Image rendering
# ---------------------------------------------------------------------------

def _render_image(node: dict, in_flex_para: bool = False) -> str:
    src = node.get("data_uri") or ""
    if not src:
        return ""

    positioning = node.get("positioning") or {}
    kind = positioning.get("kind", "inline")

    css_parts: list[str] = ["max-width: 100%", "height: auto"]

    if in_flex_para:
        # Vertical centering is handled by the flex container; just block-display
        # so the image doesn't collapse. flex-shrink:0 prevents it being squished.
        css_parts += ["display: block", "flex-shrink: 0", "order: 2", "margin-left: 8px"]
    elif kind == "anchor":
        offset_h = positioning.get("offset_h_emu")
        rel_from_h = positioning.get("relative_from_h", "")
        if offset_h is not None and rel_from_h in ("margin", "page", "column"):
            # Threshold ~3.5in (3,200,400 EMU): below = left side, above = right side.
            if offset_h >= 3_200_000:
                css_parts += ["display: block", "float: right", "margin-left: 8px"]
            else:
                css_parts += ["display: block", "float: left", "margin-right: 8px"]
        else:
            css_parts.append("display: block")
    # inline images render naturally in flow — no extra positioning

    css = "; ".join(css_parts)
    attrs = [
        f'src="{html.escape(src, quote=True)}"',
        'alt=""',
        'class="docx-image"',
        f'style="{css}"',
    ]
    if node.get("width_px"):
        attrs.append(f'width="{node["width_px"]}"')
    if node.get("height_px"):
        attrs.append(f'height="{node["height_px"]}"')

    return f'<img {" ".join(attrs)}>'


# ---------------------------------------------------------------------------
# 3.2 — Table rendering
# ---------------------------------------------------------------------------

def _render_table(node: dict, ann: Optional[dict] = None) -> str:
    ann = ann or {}
    ann_css = _annotation_block_css(ann)
    table_style = "width: 100%; border-collapse: collapse; table-layout: fixed;"
    if ann_css:
        table_style += " " + ann_css
    tbl_borders = node.get("tbl_borders") or {}
    colgroup = _render_colgroup(node.get("column_widths_twips", []))
    rows = "".join(_render_row(r, tbl_borders) for r in node.get("rows", []))
    return f'<table style="{table_style}">{colgroup}{rows}</table>'


def _render_colgroup(widths: list) -> str:
    if not widths:
        return ""
    total = sum(w for w in widths if w is not None)
    cols: list[str] = []
    if total:
        for w in widths:
            if w is None:
                cols.append("<col>")
            else:
                pct = round(w / total * 100, 2)
                cols.append(f'<col style="width: {pct}%;">')
    else:
        cols = ["<col>" for _ in widths]
    return f"<colgroup>{''.join(cols)}</colgroup>"


def _render_row(row: dict, tbl_borders: dict) -> str:
    return "<tr>" + "".join(_render_cell(c, tbl_borders) for c in row.get("cells", [])) + "</tr>"


def _render_cell(cell: dict, tbl_borders: Optional[dict] = None) -> str:
    attrs: list[str] = []
    if cell.get("colspan", 1) > 1:
        attrs.append(f'colspan="{cell["colspan"]}"')
    if cell.get("rowspan", 1) > 1:
        attrs.append(f'rowspan="{cell["rowspan"]}"')

    # 3.2 — cell alignment CSS
    css = _cell_css(cell.get("style", {}), tbl_borders or {})
    if css:
        attrs.append(f'style="{css}"')

    attr_str = (" " + " ".join(attrs)) if attrs else ""
    inner = "".join(_render_block(b) for b in cell.get("children", []))
    return f"<td{attr_str}>{inner}</td>"


def _cell_css(style: dict, tbl_borders: Optional[dict] = None) -> str:
    parts: list[str] = ["padding: 4pt", "vertical-align: top", "overflow-wrap: break-word"]

    va = style.get("vertical_align")
    if va:
        # Override default; DOCX "center" maps to CSS "middle" for table cells
        css_va = "middle" if va == "center" else va
        parts[-1] = f"vertical-align: {css_va}"

    align = style.get("alignment")
    if align:
        parts.append(f"text-align: {_map_alignment(align)}")

    bg = style.get("background")
    if bg:
        parts.append(f"background-color: #{bg}")

    # Cell-level borders take priority; fall back to table-level borders for each edge.
    # insideH covers top/bottom (inner horizontal borders), insideV covers left/right.
    cell_borders = style.get("borders") or {}
    tbl_borders = tbl_borders or {}
    _TBL_EDGE = {"top": "insideH", "bottom": "insideH", "left": "insideV", "right": "insideV"}
    for edge in ("top", "left", "bottom", "right"):
        b = cell_borders.get(edge) or tbl_borders.get(_TBL_EDGE[edge])
        if b:
            parts.append(
                f"border-{edge}: {b['width_pt']}pt {b['style']} #{b['color']}"
            )

    return "; ".join(parts)


# ---------------------------------------------------------------------------
# 4.3 — Annotation CSS helpers
# ---------------------------------------------------------------------------

_SPACING_PX = {"none": "0", "small": "8px", "medium": "16px", "large": "32px"}


def _annotation_block_css(ann: dict) -> str:
    """Translate annotation hints to inline CSS overrides."""
    parts: list[str] = []

    sp = ann.get("spacing_before")
    if sp in _SPACING_PX:
        parts.append(f"margin-top: {_SPACING_PX[sp]}")

    sp = ann.get("spacing_after")
    if sp in _SPACING_PX:
        parts.append(f"margin-bottom: {_SPACING_PX[sp]}")

    emphasis = ann.get("emphasis")
    if emphasis == "high":
        parts += ["border-left: 3px solid #3b82f6", "padding-left: 8px", "background: #eff6ff"]
    elif emphasis == "medium":
        parts += ["border-left: 2px solid #9ca3af", "padding-left: 6px"]
    elif emphasis == "low":
        parts.append("opacity: 0.75")

    if ann.get("suppress"):
        parts.append("opacity: 0.4")

    img_display = ann.get("image_display")
    if img_display == "full-width":
        parts += ["width: 100%", "display: block"]
    elif img_display == "float-right":
        parts += ["float: right", "margin-left: 16px"]

    return "; ".join(parts)


# ---------------------------------------------------------------------------
# Shared helpers
# ---------------------------------------------------------------------------

def _map_alignment(val: str) -> str:
    return {"both": "justify", "start": "left", "end": "right"}.get(val, val)


def _twips_to_px(value: str | int | None) -> Optional[int]:
    if value is None:
        return None
    try:
        return max(1, round(int(value) / 15))
    except (ValueError, TypeError):
        return None


def _wrap_html(body: str) -> str:
    return (
        "<!DOCTYPE html>\n"
        '<html lang="en">\n'
        "<head>\n"
        '<meta charset="utf-8">\n'
        '<meta name="viewport" content="width=device-width, initial-scale=1">\n'
        "<title>DOCX Preview</title>\n"
        "</head>\n"
        "<body>\n"
        f"{body}\n"
        "</body>\n"
        "</html>\n"
    )


# ---------------------------------------------------------------------------
# 2.2 — Structural validation (unchanged)
# ---------------------------------------------------------------------------

def _count_ast_nodes(ast: dict) -> dict:
    counts = {"paragraphs": 0, "tables": 0, "images": 0}

    def walk(obj: object) -> None:
        if isinstance(obj, dict):
            t = obj.get("type")
            if t == "paragraph":
                counts["paragraphs"] += 1
            elif t == "table":
                counts["tables"] += 1
            elif t == "image":
                counts["images"] += 1
            for v in obj.values():
                walk(v)
        elif isinstance(obj, list):
            for v in obj:
                walk(v)

    walk(ast)
    return counts


def _count_html_nodes(html_text: str) -> dict:
    return {
        "paragraphs": len(re.findall(r"<(?:p|h[1-6])\b", html_text, re.IGNORECASE)),
        "tables": len(re.findall(r"<table\b", html_text, re.IGNORECASE)),
        "images": len(re.findall(r"<img\b", html_text, re.IGNORECASE)),
    }


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    import json
    import sys
    from docx_ast import build_docx_ast

    if len(sys.argv) < 2:
        print("Usage: python docx_renderer.py input.docx [output.html] [annotations.json]", file=sys.stderr)
        sys.exit(2)

    try:
        ast = build_docx_ast(sys.argv[1])
    except Exception as exc:
        print(f"AST build failed: {exc}", file=sys.stderr)
        sys.exit(1)

    annotations: Optional[dict] = None
    ann_path = sys.argv[3] if len(sys.argv) > 3 else None
    if ann_path:
        with open(ann_path, encoding="utf-8") as fh:
            annotations = json.load(fh)

    rendered = render_ast_to_html(ast, annotations)
    report = validate_ast_html(ast, rendered)

    out_path = sys.argv[2] if len(sys.argv) > 2 else None
    if out_path:
        from docx_to_html import inject_page_css
        html_out = inject_page_css(rendered, ast.get("page_setup"))
        with open(out_path, "w", encoding="utf-8") as fh:
            fh.write(html_out)
        print(f"Wrote: {out_path}")

    print(json.dumps(report, indent=2))
