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
    page_setup = ast.get("page_setup", {})

    header_html = ""
    footer_html = ""
    body_blocks: list = []
    body_si = 0

    for si, section in enumerate(ast.get("sections", [])):
        kind = section.get("kind")
        blocks = section.get("blocks", [])
        ann_map = {f"b{bi}": ann_blocks.get(f"s{si}:b{bi}", {}) for bi in range(len(blocks))}
        if kind == "header":
            header_html = _render_blocks(blocks, ann_map)
        elif kind == "footer":
            footer_html = _render_blocks(blocks, ann_map)
        elif kind == "body":
            body_blocks = blocks
            body_si = si

    # Tag body blocks with their positional index BEFORE splitting pages so
    # each block carries _idx through into the page-split copies.
    for idx, block in enumerate(body_blocks):
        block["_idx"] = idx

    page_groups = _split_pages(body_blocks)
    page_divs: list[str] = []
    block_offset = 0
    for group in page_groups:
        ann_map = {
            f"b{li}": ann_blocks.get(f"s{body_si}:b{block_offset + li}", {})
            for li in range(len(group))
        }
        # tag_block_idx=False because _idx is already set on each block above
        content = _render_blocks(group, ann_map)
        hdr = f'<div class="docx-page-header">{header_html}</div>' if header_html else ""
        ftr = f'<div class="docx-page-footer">{footer_html}</div>' if footer_html else ""
        page_divs.append(f'<div class="docx-page">{hdr}<div class="docx-page-content">{content}</div>{ftr}</div>')
        block_offset += len(group)

    body = "\n".join(page_divs)
    page_css = _build_page_css(page_setup)
    return _wrap_html(f'<div class="document">\n{body}\n</div>', extra_css=page_css)


def _split_pages(blocks: list) -> list[list]:
    """Split block list into page groups.

    Handles two kinds of page-break markers on blocks:
    - page_break_before: True  — whole block starts a new page
    - page_break_info: {...}   — break is mid-block; the block is split into a
                                 head slice (current page) and tail slice (new page)
    """
    if not blocks:
        return []
    pages: list[list] = []
    current: list = []
    for block in blocks:
        if block.get("page_break_before") and current:
            pages.append(current)
            current = [block]
            continue

        pb = block.get("page_break_info")
        if pb and current:
            para_split = pb.get("para", 0)
            # Head slice: render this block up to the split paragraph
            head = dict(block)
            head["_para_split_head"] = para_split
            head.pop("page_break_info", None)
            current.append(head)
            pages.append(current)
            # Tail slice: render this block from the split paragraph onwards
            tail = dict(block)
            tail["_para_split_tail"] = para_split
            tail.pop("page_break_info", None)
            current = [tail]
            continue

        current.append(block)

    if current:
        pages.append(current)
    return pages


def _build_page_css(page_setup: dict) -> str:
    dpi = 96
    w_px = round((page_setup.get("width_in") or 8.5) * dpi)
    h_px = round((page_setup.get("height_in") or 11.0) * dpi)
    margins = page_setup.get("margins_in") or {}
    mt = round((margins.get("top") or 0.5) * dpi)
    mr = round((margins.get("right") or 0.5) * dpi)
    mb = round((margins.get("bottom") or 0.5) * dpi)
    ml = round((margins.get("left") or 0.5) * dpi)

    font_pt = page_setup.get("default_font_size_pt") or 11
    line_h = page_setup.get("default_line_height") or 1.08
    font = f"font-size:{font_pt}pt;line-height:{line_h};font-family:Calibri,'Segoe UI',Arial,sans-serif"

    after_twips = page_setup.get("default_para_after_twips")
    before_twips = page_setup.get("default_para_before_twips")
    para_mb = round(int(after_twips) * dpi / 1440) if after_twips is not None else 0
    para_mt = round(int(before_twips) * dpi / 1440) if before_twips is not None else 0

    # Build correct descendant selectors — comma list within `sel` would break descendant syntax
    zones = [".docx-page-header", ".docx-page-content", ".docx-page-footer"]
    block_tags = ["p", "h1", "h2", "h3", "h4", "h5", "h6"]
    list_tags = ["ul", "ol"]
    block_sel = ", ".join(f"{z} {t}" for z in zones for t in block_tags)
    list_sel  = ", ".join(f"{z} {t}" for z in zones for t in list_tags)

    return (
        f".document{{background:#e0e0e0;padding:24px;display:flex;flex-direction:column;align-items:center;gap:24px}}"
        f".docx-page{{width:{w_px}px;height:{h_px}px;background:white;box-shadow:0 2px 8px rgba(0,0,0,.25);box-sizing:border-box;display:flex;flex-direction:column;overflow:hidden}}"
        f".docx-page-header{{padding:{mt}px {mr}px 0 {ml}px;{font}}}"
        f".docx-page-content{{flex:1;padding:0 {mr}px 0 {ml}px;{font};overflow:hidden}}"
        f".docx-page-footer{{padding:0 {mr}px {mb}px {ml}px;{font}}}"
        # Reset browser block element margins to match Word paragraph spacing
        f"{block_sel}{{margin-top:{para_mt}px;margin-bottom:{para_mb}px}}"
        # Reset list indentation/margins to browser-neutral values
        f"{list_sel}{{margin-top:{para_mt}px;margin-bottom:{para_mb}px;padding-left:2em}}"
        # Word's Normal Table style suppresses paragraph spacing inside cells;
        # spacing there is controlled by row height, not paragraph margins.
        # Use !important to beat the zone-level `.docx-page-content p` rule (specificity 0,1,1).
        "td p,th p{margin-top:0!important;margin-bottom:0!important}"
        "td li,th li{margin-top:0!important;margin-bottom:0!important}"
        "td ul,td ol,th ul,th ol{margin-top:0!important;margin-bottom:0!important}"
    )


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
    block_idx = block.get("_idx")
    if kind == "paragraph":
        return _render_paragraph(block, ann, block_idx=block_idx)
    if kind == "table":
        head = block.get("_para_split_head")
        tail = block.get("_para_split_tail")
        return _render_table(block, ann, para_split_head=head, para_split_tail=tail, block_idx=block_idx)
    return ""


def _render_paragraph(node: dict, ann: Optional[dict] = None, block_idx: Optional[int] = None, para_idx: Optional[int] = None) -> str:
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

    children = node.get("children", [])
    inner = "".join(
        _render_inline(c, in_flex_para=use_flex, run_idx=ri)
        for ri, c in enumerate(children)
    )

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

    # Build data-* position attributes when block_idx is provided
    data_attrs = ""
    if block_idx is not None:
        p_idx = para_idx if para_idx is not None else 0
        data_attrs = f' data-block="{block_idx}" data-para="{p_idx}"'

    if not inner.strip():
        return f'<{tag} class="docx-empty-paragraph"{data_attrs}{style_attr}></{tag}>'
    return f"<{tag}{data_attrs}{style_attr}>{inner}</{tag}>"


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

def _render_inline(node: dict, in_flex_para: bool = False, run_idx: Optional[int] = None) -> str:
    kind = node.get("type")
    if kind == "run":
        return _render_run(node, in_flex_para=in_flex_para, run_idx=run_idx)
    if kind == "hyperlink":
        return _render_hyperlink(node, run_idx=run_idx)
    return ""


def _render_run(node: dict, in_flex_para: bool = False, run_idx: Optional[int] = None) -> str:
    inner = "".join(_render_run_content(c, in_flex_para=in_flex_para) for c in node.get("children", []))
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
    # Always emit a wrapper span when run_idx is set so data-run is always present.
    css = _run_font_css(style)
    if run_idx is not None:
        run_attr = f' data-run="{run_idx}"'
        style_attr = f' style="{css}"' if css else ""
        return f'<span{run_attr}{style_attr}>{inner}</span>'
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
        # Preserve runs of 2+ spaces that HTML would otherwise collapse
        text = re.sub(r" {2,}", lambda m: " " * len(m.group()), text)
        return text
    if kind == "break":
        return "&emsp;" if node.get("break_type") == "tab" else "<br>"
    if kind == "image":
        return _render_image(node, in_flex_para=in_flex_para)
    if kind == "checkbox":
        char = "☒" if node.get("checked") else "☐"
        return f'<span data-type="checkbox" style="font-family:serif;font-size:1.1em;vertical-align:middle">{char}</span>'
    return ""


def _render_hyperlink(node: dict, run_idx: Optional[int] = None) -> str:
    inner = "".join(_render_run(r) for r in node.get("children", []))
    if not inner:
        return ""
    href = node.get("href")
    run_attr = f' data-run="{run_idx}"' if run_idx is not None else ""
    if href:
        return f'<a{run_attr} href="{html.escape(href, quote=True)}">{inner}</a>'
    if run_attr:
        return f'<span{run_attr}>{inner}</span>'
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
# 3.1b — Block-list renderer (handles list grouping)
# ---------------------------------------------------------------------------

def _para_inner_html(node: dict) -> str:
    """Return the inline content of a paragraph node without the outer tag."""
    use_flex = _para_has_text(node) and _para_anchor_on_right(node)
    return "".join(_render_inline(c, in_flex_para=use_flex) for c in node.get("children", []))


def _render_blocks_in_cell(blocks: list, block_idx: Optional[int] = None) -> str:
    """Like _render_blocks but for cell children: threads block_idx/para_idx into
    paragraphs for data-* attributes while still grouping list items into <ul>/<ol>."""
    parts: list[str] = []
    i = 0
    while i < len(blocks):
        block = blocks[i]
        num_id = block.get("style", {}).get("num_id") if block.get("type") == "paragraph" else None
        if num_id:
            group: list[dict] = []
            pi_start = i
            while (
                i < len(blocks)
                and blocks[i].get("type") == "paragraph"
                and blocks[i].get("style", {}).get("num_id") == num_id
            ):
                group.append(blocks[i])
                i += 1
            style0 = group[0].get("style", {})
            list_tag = "ol" if style0.get("list_ordered") else "ul"
            marker = style0.get("list_marker", "disc")
            lis = "".join(
                f"<li>{_para_inner_html_with_idx(item, block_idx=block_idx, para_idx=pi_start + gi)}</li>"
                for gi, item in enumerate(group)
            )
            parts.append(f'<{list_tag} style="list-style-type: {marker}">{lis}</{list_tag}>')
        else:
            if block.get("type") == "paragraph":
                parts.append(_render_paragraph(block, para_idx=i, block_idx=block_idx))
            else:
                parts.append(_render_block(block))
            i += 1
    return "".join(parts)


def _para_inner_html_with_idx(node: dict, block_idx: Optional[int] = None, para_idx: Optional[int] = None) -> str:
    """Same as _para_inner_html but used inside list items — extracts just the inner HTML."""
    use_flex = (
        node.get("style", {}).get("alignment") == "right"
        and not _para_has_text(node)
        and _para_anchor_on_right(node)
    )
    children = node.get("children", [])
    return "".join(
        _render_inline(c, in_flex_para=use_flex, run_idx=ri)
        for ri, c in enumerate(children)
    )


def _render_blocks(blocks: list, ann_map: dict | None = None, tag_block_idx: bool = False) -> str:
    """Render a sequence of blocks, grouping consecutive list paragraphs into <ul>/<ol>.

    When tag_block_idx is True (body section only), each block gets its positional
    index stored as block["_idx"] before rendering so data-block attributes are emitted.
    """
    ann_map = ann_map or {}
    if tag_block_idx:
        for idx, block in enumerate(blocks):
            block["_idx"] = idx
    parts: list[str] = []
    i = 0
    while i < len(blocks):
        block = blocks[i]
        num_id = block.get("style", {}).get("num_id") if block.get("type") == "paragraph" else None
        if num_id:
            group: list[dict] = []
            while (
                i < len(blocks)
                and blocks[i].get("type") == "paragraph"
                and blocks[i].get("style", {}).get("num_id") == num_id
            ):
                group.append(blocks[i])
                i += 1
            style0 = group[0].get("style", {})
            list_tag = "ol" if style0.get("list_ordered") else "ul"
            marker = style0.get("list_marker", "disc")
            lis = "".join(f"<li>{_para_inner_html(item)}</li>" for item in group)
            parts.append(f'<{list_tag} style="list-style-type: {marker}">{lis}</{list_tag}>')
        else:
            parts.append(_render_block(block, ann_map.get(f"b{i}", {})))
            i += 1
    return "".join(parts)


# ---------------------------------------------------------------------------
# 3.2 — Table rendering
# ---------------------------------------------------------------------------

def _render_table(
    node: dict,
    ann: Optional[dict] = None,
    para_split_head: Optional[int] = None,
    para_split_tail: Optional[int] = None,
    block_idx: Optional[int] = None,
) -> str:
    ann = ann or {}
    ann_css = _annotation_block_css(ann)
    table_style = "width: 100%; border-collapse: collapse; table-layout: fixed;"
    if ann_css:
        table_style += " " + ann_css
    tbl_borders = node.get("tbl_borders") or {}
    tbl_cell_margins = node.get("tbl_cell_margins") or {}
    colgroup = _render_colgroup(node.get("column_widths_twips", []))
    rows = "".join(
        _render_row(r, tbl_borders, tbl_cell_margins, row_idx=ri, block_idx=block_idx, para_split_head=para_split_head, para_split_tail=para_split_tail)
        for ri, r in enumerate(node.get("rows", []))
    )
    data_attrs = f' data-block="{block_idx}"' if block_idx is not None else ""
    return f'<table{data_attrs} style="{table_style}">{colgroup}{rows}</table>'


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


def _render_row(
    row: dict,
    tbl_borders: dict,
    tbl_cell_margins: dict | None = None,
    row_idx: Optional[int] = None,
    block_idx: Optional[int] = None,
    para_split_head: Optional[int] = None,
    para_split_tail: Optional[int] = None,
) -> str:
    h = row.get("height_twips")
    row_style = f' style="height: {_twips_to_px(h)}px"' if h else ""
    data_attrs = f' data-row="{row_idx}"' if row_idx is not None else ""
    cells = "".join(
        _render_cell(c, tbl_borders, tbl_cell_margins, cell_idx=ci, block_idx=block_idx, row_idx=row_idx, para_split_head=para_split_head, para_split_tail=para_split_tail)
        for ci, c in enumerate(row.get("cells", []))
    )
    return f"<tr{data_attrs}{row_style}>{cells}</tr>"


def _render_cell(
    cell: dict,
    tbl_borders: Optional[dict] = None,
    tbl_cell_margins: Optional[dict] = None,
    cell_idx: Optional[int] = None,
    block_idx: Optional[int] = None,
    row_idx: Optional[int] = None,
    para_split_head: Optional[int] = None,
    para_split_tail: Optional[int] = None,
) -> str:
    attrs: list[str] = []
    if cell_idx is not None:
        attrs.append(f'data-cell="{cell_idx}"')
    if cell.get("colspan", 1) > 1:
        attrs.append(f'colspan="{cell["colspan"]}"')
    if cell.get("rowspan", 1) > 1:
        attrs.append(f'rowspan="{cell["rowspan"]}"')

    css = _cell_css(cell.get("style", {}), tbl_borders or {}, tbl_cell_margins or {})
    if css:
        attrs.append(f'style="{css}"')

    attr_str = (" " + " ".join(attrs)) if attrs else ""
    children = cell.get("children", [])
    if para_split_head is not None:
        children = children[:para_split_head]
    elif para_split_tail is not None:
        children = children[para_split_tail:]

    inner = _render_blocks_in_cell(children, block_idx=block_idx)
    return f"<td{attr_str}>{inner}</td>"


def _cell_css(style: dict, tbl_borders: Optional[dict] = None, tbl_cell_margins: Optional[dict] = None) -> str:
    parts: list[str] = []

    # Padding: prefer cell-level tcMar, then table-style default, then fallback
    tbl_cell_margins = tbl_cell_margins or {}
    if tbl_cell_margins:
        top_px    = _twips_to_px(tbl_cell_margins.get("top",    0)) or 0
        right_px  = _twips_to_px(tbl_cell_margins.get("right",  0)) or 0
        bottom_px = _twips_to_px(tbl_cell_margins.get("bottom", 0)) or 0
        left_px   = _twips_to_px(tbl_cell_margins.get("left",   0)) or 0
        parts.append(f"padding: {top_px}px {right_px}px {bottom_px}px {left_px}px")
    else:
        # Word's Normal Table default: 0px top/bottom, 108 twips (7px) left/right
        parts.append("padding: 0px 7px")

    va = style.get("vertical_align")
    css_va = "middle" if va == "center" else (va or "top")
    parts.append(f"vertical-align: {css_va}")
    parts.append("overflow-wrap: break-word")

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
        b = cell_borders.get(edge) or tbl_borders.get(_TBL_EDGE[edge]) or tbl_borders.get(edge)
        if b:
            parts.append(f"border-{edge}: {b['width_pt']}pt {b['style']} #{b['color']}")

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
    # Map "both" (Word's justify) to "left" — browser justify stretches word-spacing on
    # short wrapped lines producing unusable output vs Word's sophisticated layout engine.
    return {"both": "left", "start": "left", "end": "right"}.get(val, val)


def _twips_to_px(value: str | int | None) -> Optional[int]:
    if value is None:
        return None
    try:
        return max(1, round(int(value) / 15))
    except (ValueError, TypeError):
        return None


def _wrap_html(body: str, extra_css: str = "") -> str:
    style_block = f"<style>\n{extra_css}\n</style>\n" if extra_css else ""
    return (
        "<!DOCTYPE html>\n"
        '<html lang="en">\n'
        "<head>\n"
        '<meta charset="utf-8">\n'
        '<meta name="viewport" content="width=device-width, initial-scale=1">\n'
        "<title>DOCX Preview</title>\n"
        f"{style_block}"
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

    use_ai = "--use-ai" in sys.argv
    if use_ai:
        sys.argv.remove("--use-ai")

    if len(sys.argv) < 2:
        print("Usage: python docx_renderer.py input.docx [output.html] [--use-ai]", file=sys.stderr)
        sys.exit(2)

    try:
        ast = build_docx_ast(sys.argv[1])
    except Exception as exc:
        print(f"AST build failed: {exc}", file=sys.stderr)
        sys.exit(1)

    annotations: Optional[dict] = None
    
    # If --use-ai flag is set, generate annotations using Claude
    if use_ai:
        try:
            from docx_annotator import generate_annotations
            annotations = generate_annotations(ast)
            print(f"[docx_renderer] AI annotations generated", file=sys.stderr)
        except Exception as exc:
            print(f"[docx_renderer] Warning: AI annotation failed ({exc}), continuing without annotations", file=sys.stderr)
    else:
        # Otherwise, check for annotations file as third arg
        ann_path = sys.argv[3] if len(sys.argv) > 3 else None
        if ann_path:
            with open(ann_path, encoding="utf-8") as fh:
                annotations = json.load(fh)

    rendered = render_ast_to_html(ast, annotations)
    report = validate_ast_html(ast, rendered)

    out_path = sys.argv[2] if len(sys.argv) > 2 else None
    if out_path:
        with open(out_path, "w", encoding="utf-8") as fh:
            fh.write(rendered)
        print(f"Wrote: {out_path}")

    print(json.dumps(report, indent=2))
