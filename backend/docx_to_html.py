from __future__ import annotations

import base64
import html
import json
from html.parser import HTMLParser
import mimetypes
import os
import re
import sys
import zipfile
from pathlib import Path
from typing import Dict, Iterable, Optional
from xml.etree import ElementTree as ET


W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
R_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
A_NS = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"


class DocxConversionError(RuntimeError):
    pass


class PdfConversionError(RuntimeError):
    pass


VOID_HTML_TAGS = {"br", "img", "meta", "link", "hr", "input"}
PRESERVE_EMPTY_HTML_TAGS = {"p", "td", "tr", "table", "colgroup", "col"}

A4_STYLESHEET = """
<style>
@page {
    size: A4;
    margin: 12mm;
}

html, body {
    width: 100%;
    margin: 0;
    padding: 0;
}

body {
    color: #111827;
    font-family: Arial, Helvetica, sans-serif;
    font-size: 10pt;
    line-height: 1.25;
}

.document {
    width: 100%;
}

.document p,
.document h1,
.document h2,
.document h3,
.document h4,
.document h5,
.document h6,
.document table {
    margin: 0 0 4pt;
}

.docx-empty-paragraph {
    margin: 0;
    line-height: 0.2;
    min-height: 0.2em;
}

.document p,
.document h1,
.document h2,
.document h3,
.document h4,
.document h5,
.document h6,
.document img {
    break-inside: avoid;
    page-break-inside: avoid;
}

.document h1,
.document h2,
.document h3,
.document h4,
.document h5,
.document h6 {
    break-after: avoid-page;
    page-break-after: avoid;
}

.document h1,
.document h2,
.document h3,
.document h4,
.document h5,
.document h6 {
    line-height: 1.2;
    font-weight: 700;
}

.document h1 { font-size: 18pt; }
.document h2 { font-size: 14pt; }
.document h3 { font-size: 12pt; }
.document h4 { font-size: 11pt; }
.document h5 { font-size: 10pt; }
.document h6 { font-size: 9pt; }

.document table {
    width: 100%;
    border-collapse: collapse;
    table-layout: auto;
}

.doc-table td,
.doc-table th {
    border: 1px solid #d1d5db;
    padding: 6pt;
    vertical-align: top;
    word-wrap: break-word;
    overflow-wrap: anywhere;
}

.docx-image {
    max-width: 100%;
    height: auto;
}
</style>
""".strip()


def convert_docx_to_html(docx_path: str | Path) -> str:
    return _convert_docx_to_html(docx_path)


def convert_docx_to_html_with_assets(docx_path: str | Path, assets_dir: str | Path) -> str:
    return _convert_docx_to_html(docx_path, assets_dir=Path(assets_dir))


def _convert_docx_to_html(docx_path: str | Path, assets_dir: Optional[Path] = None) -> str:
    path = Path(docx_path)
    if not path.is_file():
        raise DocxConversionError(f"DOCX file not found: {path}")

    try:
        with zipfile.ZipFile(path) as archive:
            document_xml = _read_xml(archive, "word/document.xml")
            styles = _parse_styles(_read_xml(archive, "word/styles.xml"))
            relationships = _parse_relationships(_read_xml(archive, "word/_rels/document.xml.rels"))
            media = _load_media(archive)

            # load any headers/footers referenced by the document relationships
            header_parts: list[tuple[ET.Element, Dict[str, Dict[str, str]]]] = []
            footer_parts: list[tuple[ET.Element, Dict[str, Dict[str, str]]]] = []
            for rel in relationships.values():
                rtype = rel.get("type", "")
                target = rel.get("target", "")
                if not target or not rtype:
                    continue
                # Only process header/footer relationship types
                if not (rtype.endswith('/header') or rtype.endswith('/headers') or 
                        rtype.endswith('/footer') or rtype.endswith('/footers')):
                    continue
                
                # relationship targets are relative to the "word/" folder
                part_path = f"word/{target}" if not target.startswith("word/") else target
                try:
                    part_xml = _read_xml(archive, part_path)
                    # load relationships for this part (e.g., "word/footer1.xml" -> "word/_rels/footer1.xml.rels")
                    # extract filename from part_path
                    part_dir, part_filename = part_path.rsplit("/", 1)
                    part_rels_path = f"{part_dir}/_rels/{part_filename}.rels"
                    try:
                        part_rels_root = _read_xml(archive, part_rels_path)
                        part_relationships = _parse_relationships(part_rels_root)
                    except KeyError:
                        # part has no relationships
                        part_relationships = {}
                    
                    if rtype.endswith('/header') or rtype.endswith('/headers'):
                        header_parts.append((part_xml, part_relationships))
                    elif rtype.endswith('/footer') or rtype.endswith('/footers'):
                        footer_parts.append((part_xml, part_relationships))
                except KeyError:
                    # missing optional part; ignore
                    pass
    except zipfile.BadZipFile as exc:
        raise DocxConversionError(f"Invalid DOCX archive: {path}") from exc
    except KeyError as exc:
        raise DocxConversionError(f"Missing required DOCX part: {exc.args[0]}") from exc

    body = document_xml.find(f"{W_NS}body")
    if body is None:
        raise DocxConversionError("DOCX document body is missing")

    blocks: list[str] = []

    # render header parts first (so header images/branding appear at top)
    for hdr, hdr_rels in header_parts:
        for child in hdr:
            tag = _local_name(child.tag)
            if tag == "p":
                blocks.append(_render_paragraph(child, styles, hdr_rels, media, assets_dir))
            elif tag == "tbl":
                blocks.append(_render_table(child, styles, hdr_rels, media, assets_dir))

    # render the main document body
    for child in body:
        tag = _local_name(child.tag)
        if tag == "p":
            blocks.append(_render_paragraph(child, styles, relationships, media, assets_dir))
        elif tag == "tbl":
            blocks.append(_render_table(child, styles, relationships, media, assets_dir))

    # render footer parts after the body
    for ftr, ftr_rels in footer_parts:
        for child in ftr:
            tag = _local_name(child.tag)
            if tag == "p":
                blocks.append(_render_paragraph(child, styles, ftr_rels, media, assets_dir))
            elif tag == "tbl":
                blocks.append(_render_table(child, styles, ftr_rels, media, assets_dir))

    return _wrap_html("\n".join(blocks))


def normalize_html(raw_html: str) -> str:
    parser = _HtmlNormalizer()
    parser.feed(raw_html)
    parser.close()
    source_children = _extract_normalization_source(parser.root)
    normalized_children = _normalize_children(source_children)
    normalized_children = _trim_root_edge_spacers(normalized_children)
    body = _serialize_children([_Node("div", attrs={"class": "document"}, children=normalized_children)])
    return _wrap_html(body)


def convert_docx_to_normalized_html(docx_path: str | Path) -> str:
    return normalize_html(convert_docx_to_html(docx_path))


def convert_docx_to_normalized_html_with_assets(docx_path: str | Path, assets_dir: str | Path) -> str:
    # When assets are being extracted, preserve the original HTML structure
    # so image `img` elements are not removed by normalization heuristics.
    return convert_docx_to_html_with_assets(docx_path, assets_dir)


def inject_a4_css(html_document: str) -> str:
    style_block = f"\n{A4_STYLESHEET}\n"
    head_close = html_document.find("</head>")
    if head_close == -1:
        return html_document + style_block
    return html_document[:head_close] + style_block + html_document[head_close:]


def inject_page_css(html_document: str, page_setup: dict | None = None) -> str:
    """Like inject_a4_css but uses actual page dimensions from the DOCX AST.

    Falls back to A4/12mm defaults when page_setup is absent or incomplete.
    """
    if not page_setup:
        return inject_a4_css(html_document)

    w = page_setup.get("width_in")
    h = page_setup.get("height_in")
    size_css = f"{w}in {h}in" if (w and h) else "A4"

    margins = page_setup.get("margins_in", {})
    top    = margins.get("top",    0.5)
    right  = margins.get("right",  0.5)
    bottom = margins.get("bottom", 0.5)
    left   = margins.get("left",   0.5)

    page_rule = (
        f"@page {{\n"
        f"    size: {size_css};\n"
        f"    margin: {top}in {right}in {bottom}in {left}in;\n"
        f"}}"
    )

    # Replace only the @page block in A4_STYLESHEET, keep the rest
    import re as _re
    base = _re.sub(r"@page\s*\{[^}]*\}", page_rule, A4_STYLESHEET, count=1)
    style_block = f"\n<style>\n{base}\n</style>\n"

    head_close = html_document.find("</head>")
    if head_close == -1:
        return html_document + style_block
    return html_document[:head_close] + style_block + html_document[head_close:]


def convert_docx_to_a4_html(docx_path: str | Path) -> str:
    return inject_a4_css(convert_docx_to_normalized_html(docx_path))


def convert_docx_to_a4_html_with_assets(docx_path: str | Path, assets_dir: str | Path) -> str:
    return inject_a4_css(convert_docx_to_normalized_html_with_assets(docx_path, assets_dir))


def convert_html_to_pdf(
    html_document: str,
    pdf_path: str | Path | None = None,
) -> bytes:
    try:
        from weasyprint import HTML
    except ImportError as exc:
        raise PdfConversionError("weasyprint is required for PDF conversion") from exc

    pdf_bytes = HTML(string=html_document).write_pdf(presentational_hints=True)
    if pdf_path is not None:
        Path(pdf_path).write_bytes(pdf_bytes)
    return pdf_bytes


def convert_docx_to_pdf(
    docx_path: str | Path,
    pdf_path: str | Path | None = None,
) -> bytes:
    html_document = convert_docx_to_a4_html(docx_path)
    return convert_html_to_pdf(html_document, pdf_path=pdf_path)


def _read_xml(archive: zipfile.ZipFile, name: str) -> ET.Element:
    with archive.open(name) as handle:
        return ET.fromstring(handle.read())


def _analyze_docx_structure(docx_path: str | Path) -> dict:
    path = Path(docx_path)
    result = {"paragraphs": 0, "tables": 0, "images": 0}
    try:
        with zipfile.ZipFile(path) as archive:
            doc = _read_xml(archive, "word/document.xml")
            for el in doc.iter():
                ln = _local_name(el.tag)
                if ln == "p":
                    result["paragraphs"] += 1
                elif ln == "tbl":
                    result["tables"] += 1
                elif ln in {"blip", "blipfill"}:
                    result["images"] += 1
    except Exception:
        pass
    return result


def _analyze_html_structure(html_text: str) -> dict:
    parser = _HtmlNormalizer()
    parser.feed(html_text)
    parser.close()

    counts = {"paragraphs": 0, "tables": 0, "images": 0}

    def walk(node: _Node | str) -> None:
        if isinstance(node, str):
            return
        tag = node.tag
        if tag == "p":
            counts["paragraphs"] += 1
        elif tag == "table":
            counts["tables"] += 1
        elif tag == "img":
            counts["images"] += 1
        for child in node.children:
            walk(child)

    for child in parser.root.children:
        walk(child)

    return counts


def _validate_visual(docx_path: str, generated_html: str, reference_path: Optional[str] = None) -> dict:
    """Run a lightweight Phase 4 validation: structural comparisons.

    This produces counts of paragraphs/tables/images for the original DOCX
    and the generated HTML. If a reference HTML is provided, it will be
    analyzed as well.
    """
    report: dict = {"docx": {}, "generated": {}, "reference": None, "summary": []}

    docx_stats = _analyze_docx_structure(docx_path)
    html_stats = _analyze_html_structure(generated_html)

    report["docx"] = docx_stats
    report["generated"] = html_stats

    if reference_path:
        try:
            with open(reference_path, "r", encoding="utf-8") as fh:
                ref_html = fh.read()
            ref_stats = _analyze_html_structure(ref_html)
            report["reference"] = ref_stats
        except Exception as exc:
            report["reference"] = {"error": str(exc)}

    # simple comparisons
    for key in ("paragraphs", "tables", "images"):
        d = docx_stats.get(key, 0)
        g = html_stats.get(key, 0)
        if d != g:
            report["summary"].append(f"Mismatch {key}: docx={d} vs generated={g}")
        else:
            report["summary"].append(f"Match {key}: {d}")

    if report["reference"] and isinstance(report["reference"], dict) and "error" not in report["reference"]:
        for key in ("paragraphs", "tables", "images"):
            r = report["reference"].get(key, 0)
            g = html_stats.get(key, 0)
            if r != g:
                report["summary"].append(f"Reference mismatch {key}: reference={r} vs generated={g}")
            else:
                report["summary"].append(f"Reference match {key}: {g}")

    return report


def _extract_paragraph_style(paragraph: ET.Element) -> dict:
    ppr = paragraph.find(f"{W_NS}pPr")
    if ppr is None:
        return {}

    result: dict = {}
    jc = ppr.find(f"{W_NS}jc")
    if jc is not None:
        val = jc.get(f"{W_NS}val") or jc.get("val")
        if val:
            result["alignment"] = val

    spacing = ppr.find(f"{W_NS}spacing")
    if spacing is not None:
        spacing_obj: dict = {}
        for key in ("before", "after", "line"):
            v = spacing.get(f"{W_NS}{key}") or spacing.get(key)
            if v is not None:
                spacing_obj[key] = v
        if spacing_obj:
            result["spacing"] = spacing_obj

    ind = ppr.find(f"{W_NS}ind")
    if ind is not None:
        ind_obj: dict = {}
        for key in ("left", "right", "firstLine", "hanging"):
            v = ind.get(f"{W_NS}{key}") or ind.get(key)
            if v is not None:
                ind_obj[key] = v
        if ind_obj:
            result["indent"] = ind_obj

    return result


def _extract_run_style(run_pr: ET.Element) -> dict:
    result: dict = {}
    # font size (w:sz is in half-points)
    sz = run_pr.find(f"{W_NS}sz")
    if sz is not None:
        val = sz.get(f"{W_NS}val") or sz.get("val")
        if val:
            try:
                pts = int(val) / 2
                # prefer integer when possible
                result["font_size_pt"] = int(pts) if pts.is_integer() else round(pts, 1)
            except Exception:
                result["font_size"] = val

    rfonts = run_pr.find(f"{W_NS}rFonts")
    if rfonts is not None:
        # attempt common attributes
        for key in ("ascii", "hAnsi", "cs"):
            v = rfonts.get(f"{W_NS}{key}") or rfonts.get(key)
            if v:
                result["font_family"] = v
                break

    color = run_pr.find(f"{W_NS}color")
    if color is not None:
        v = color.get(f"{W_NS}val") or color.get("val")
        if v:
            result["color"] = v

    # include explicit flags for quick checks
    if run_pr.find(f"{W_NS}b") is not None:
        result["bold"] = True
    if run_pr.find(f"{W_NS}i") is not None:
        result["italic"] = True
    if run_pr.find(f"{W_NS}u") is not None:
        result["underline"] = True

    return result


def _extract_cell_style(cell: ET.Element) -> dict:
    tcpr = cell.find(f"{W_NS}tcPr")
    if tcpr is None:
        return {}
    result: dict = {}
    valign = tcpr.find(f"{W_NS}vAlign")
    if valign is not None:
        v = valign.get(f"{W_NS}val") or valign.get("val")
        if v:
            result["vertical_align"] = v
    # horizontal alignment may be declared at the cell or paragraph level
    jc = tcpr.find(f"{W_NS}jc")
    if jc is not None:
        v = jc.get(f"{W_NS}val") or jc.get("val")
        if v:
            result["alignment"] = v

    # fallback: look for first paragraph alignment inside the cell
    first_p = cell.find(f"{W_NS}p")
    if first_p is not None:
        pjc = first_p.find(f"{W_NS}pPr/{W_NS}jc")
        if pjc is not None:
            v = pjc.get(f"{W_NS}val") or pjc.get("val")
            if v and "alignment" not in result:
                result["alignment"] = v

    return result


def _paragraph_style_to_css(pstyle: dict) -> str:
    parts: list[str] = []
    if not pstyle:
        return ""

    align = pstyle.get("alignment")
    if align:
        parts.append(f"text-align: {align};")

    spacing = pstyle.get("spacing") or {}
    before = spacing.get("before")
    after = spacing.get("after")
    line = spacing.get("line")
    if before is not None:
        px = _twips_to_px(before)
        if px is not None:
            parts.append(f"margin-top: {px}px;")
    if after is not None:
        px = _twips_to_px(after)
        if px is not None:
            parts.append(f"margin-bottom: {px}px;")
    if line is not None:
        try:
            ln = int(line)
            # heuristic: convert twentieths of a point to a relative line-height
            lh = round(ln / 240, 2)
            parts.append(f"line-height: {lh};")
        except Exception:
            pass

    indent = pstyle.get("indent") or {}
    left = indent.get("left")
    right = indent.get("right")
    first = indent.get("firstLine") or indent.get("first")
    if left is not None:
        px = _twips_to_px(left)
        if px is not None:
            parts.append(f"padding-left: {px}px;")
    if right is not None:
        px = _twips_to_px(right)
        if px is not None:
            parts.append(f"padding-right: {px}px;")
    if first is not None:
        px = _twips_to_px(first)
        if px is not None:
            parts.append(f"text-indent: {px}px;")

    return " ".join(parts)


def _run_style_to_css(rstyle: dict) -> str:
    parts: list[str] = []
    if not rstyle:
        return ""
    fs = rstyle.get("font_size_pt")
    if fs is not None:
        parts.append(f"font-size: {fs}pt;")
    ff = rstyle.get("font_family")
    if ff:
        parts.append(f"font-family: {ff};")
    color = rstyle.get("color")
    if color:
        c = color if color.startswith("#") else f"#{color}"
        parts.append(f"color: {c};")
    if rstyle.get("bold"):
        parts.append("font-weight: 700;")
    if rstyle.get("italic"):
        parts.append("font-style: italic;")
    if rstyle.get("underline"):
        parts.append("text-decoration: underline;")
    return " ".join(parts)


def _cell_style_to_css(cstyle: dict) -> str:
    parts: list[str] = []
    if not cstyle:
        return ""
    valign = cstyle.get("vertical_align")
    if valign:
        v = valign
        if v == "center":
            v = "middle"
        parts.append(f"vertical-align: {v};")
    # horizontal alignment
    align = cstyle.get("alignment")
    if align:
        a = align
        if a == "both":
            a = "justify"
        parts.append(f"text-align: {a};")
    return " ".join(parts)


def _parse_styles(styles_root: ET.Element) -> Dict[str, str]:
    styles: Dict[str, str] = {}
    for style in styles_root.findall(f"{W_NS}style"):
        style_id = style.get(f"{W_NS}styleId")
        style_type = style.get(f"{W_NS}type")
        name = style.find(f"{W_NS}name")
        if style_id and style_type == "paragraph" and name is not None:
            styles[style_id] = name.get(f"{W_NS}val", "")
    return styles


def _parse_relationships(rels_root: ET.Element) -> Dict[str, Dict[str, str]]:
    relationships: Dict[str, Dict[str, str]] = {}
    for rel in rels_root.findall(f"{{{REL_NS}}}Relationship"):
        rel_id = rel.get("Id")
        if rel_id:
            relationships[rel_id] = {
                "type": rel.get("Type", ""),
                "target": rel.get("Target", ""),
            }
    return relationships


def _load_media(archive: zipfile.ZipFile) -> Dict[str, bytes]:
    media: Dict[str, bytes] = {}
    for name in archive.namelist():
        if name.startswith("word/media/"):
            with archive.open(name) as handle:
                media[name] = handle.read()
    return media


def _find_media_bytes(media: Dict[str, bytes], target: str) -> tuple[Optional[str], Optional[bytes]]:
    """Find the best-match media entry for a relationship target.

    Returns a tuple of (resolved_media_path, bytes) or (None, None) when not found.
    Matching strategy (deterministic):
    - Normalize target via _resolve_media_path and try exact match
    - Try the raw target as given
    - Try stripping any leading "word/" prefix
    - Fallback to matching by basename if there is a unique match
    """
    if not target:
        return None, None

    candidates = []

    # primary normalized form
    primary = _resolve_media_path(target)
    candidates.append(primary)

    # raw target as-is
    candidates.append(target)

    # strip leading word/ if present
    if target.startswith("word/"):
        candidates.append(target[len("word/"):])

    # strip leading slashes
    candidates.append(target.lstrip("/"))

    # try candidates in order
    for cand in candidates:
        if cand in media:
            return cand, media[cand]

    # fallback: match by basename if unique
    basename = os.path.basename(target)
    if basename:
        matches = [k for k in media.keys() if os.path.basename(k) == basename]
        if len(matches) == 1:
            key = matches[0]
            return key, media[key]

    return None, None


def _render_paragraph(
    paragraph: ET.Element,
    styles: Dict[str, str],
    relationships: Dict[str, Dict[str, str]],
    media: Dict[str, bytes],
    assets_dir: Optional[Path] = None,
) -> str:
    style_id = None
    style_element = paragraph.find(f"{W_NS}pPr/{W_NS}pStyle")
    if style_element is not None:
        style_id = style_element.get(f"{W_NS}val")

    tag = _paragraph_tag(style_id, styles)
    fragments = [
        _render_inline(child, relationships, media, assets_dir) for child in paragraph if _local_name(child.tag) in {"r", "hyperlink"}
    ]
    content = _join_html_fragments(fragments)

    # extract paragraph-level styling attributes (Phase 1)
    p_style = _extract_paragraph_style(paragraph)
    style_attr = html.escape(json.dumps(p_style, separators=(",", ":"), ensure_ascii=False), quote=True)
    css = _paragraph_style_to_css(p_style)
    css_attr = f' style="{html.escape(css, quote=True)}"' if css else ""

    if not content.strip():
        return f'<p class="docx-empty-paragraph" data-docx-style="{style_attr}"{css_attr}></p>'

    return f"<{tag} data-docx-style=\"{style_attr}\"{css_attr}>{content}</{tag}>"


def _paragraph_tag(style_id: Optional[str], styles: Dict[str, str]) -> str:
    if not style_id:
        return "p"

    style_name = styles.get(style_id, style_id).strip().lower()
    match = re.search(r"heading\s*([1-6])", style_name)
    if match:
        return f"h{match.group(1)}"

    return "p"


def _render_table(
    table: ET.Element,
    styles: Dict[str, str],
    relationships: Dict[str, Dict[str, str]],
    media: Dict[str, bytes],
    assets_dir: Optional[Path] = None,
) -> str:
    column_widths = _parse_table_column_widths(table)
    logical_rows = _build_table_rows(table, styles, relationships, media, assets_dir)
    rows = []
    for row_entries in logical_rows:
        cells = []
        for entry in row_entries:
            if entry["skip"]:
                continue

            cell_attrs = []
            if entry["colspan"] > 1:
                cell_attrs.append(f' colspan="{entry["colspan"]}"')
            if entry["rowspan"] > 1:
                cell_attrs.append(f' rowspan="{entry["rowspan"]}"')
            if entry["width"] is not None:
                cell_attrs.append(f' width="{entry["width"]}"')
            # attach extracted cell style metadata when available and apply inline CSS
            if entry.get("cell_style"):
                style_attr = html.escape(json.dumps(entry["cell_style"], separators=(",", ":"), ensure_ascii=False), quote=True)
                css = _cell_style_to_css(entry["cell_style"])
                css_attr = f' style="{html.escape(css, quote=True)}"' if css else ""
                cell_attrs.append(f' data-docx-style="{style_attr}"')
                if css_attr:
                    cell_attrs.append(css_attr)
            cells.append(f"<td{''.join(cell_attrs)}>{_join_html_fragments(entry['blocks'])}</td>")
        rows.append(f"<tr>{''.join(cells)}</tr>")

    # generate colgroup using percentages so table can reliably be 100% width
    colgroup = ""
    if column_widths:
        total = sum(w for w in column_widths if w is not None)
        columns = []
        if total and total > 0:
            for width in column_widths:
                if width is None:
                    columns.append("<col>")
                else:
                    pct = round(width / total * 100, 2)
                    columns.append(f'<col style="width: {pct}%;">')
        else:
            columns = [f'<col width="{width}">' if width is not None else "<col>" for width in column_widths]

        colgroup = f"<colgroup>{''.join(columns)}</colgroup>"

    # ensure table occupies full container width (Phase 3 requirement)
    table_style = ' style="width: 100%;"'
    return f"<table{table_style}>{colgroup}{''.join(rows)}</table>"


def _build_table_rows(
    table: ET.Element,
    styles: Dict[str, str],
    relationships: Dict[str, Dict[str, str]],
    media: Dict[str, bytes],
    assets_dir: Optional[Path] = None,
) -> list[list[dict[str, object]]]:
    row_cells: list[list[dict[str, object]]] = []
    for row_index, row in enumerate(table.findall(f"{W_NS}tr")):
        entries = []
        current_column = 0
        for cell in row.findall(f"{W_NS}tc"):
            colspan = _parse_cell_grid_span(cell)
            vmerge = _parse_vertical_merge(cell)
            cell_width = _parse_cell_width(cell)
            cell_blocks = []
            for child in cell:
                tag = _local_name(child.tag)
                if tag == "p":
                    rendered = _render_paragraph(child, styles, relationships, media, assets_dir)
                elif tag == "tbl":
                    rendered = _render_table(child, styles, relationships, media, assets_dir)
                else:
                    rendered = ""
                if rendered:
                    cell_blocks.append(rendered)

            # extract cell-level styles (vertical alignment etc.) for Phase 1
            cell_style = _extract_cell_style(cell)

            entries.append(
                {
                    "row_index": row_index,
                    "column_index": current_column,
                    "colspan": colspan,
                    "rowspan": 1,
                    "vmerge": vmerge,
                    "width": cell_width,
                    "blocks": cell_blocks,
                    "skip": False,
                    "cell_style": cell_style,
                }
            )
            current_column += colspan
        row_cells.append(entries)

    for row_index, entries in enumerate(row_cells):
        for entry in entries:
            if entry["vmerge"] != "restart":
                continue

            rowspan = 1
            for next_row in row_cells[row_index + 1 :]:
                match = _find_table_cell_by_column(next_row, entry["column_index"], entry["colspan"])
                if match is None or match["vmerge"] != "continue":
                    break
                match["skip"] = True
                rowspan += 1

            entry["rowspan"] = rowspan

        for entry in entries:
            if entry["vmerge"] == "continue":
                entry["skip"] = True

    return row_cells


def _find_table_cell_by_column(
    row_entries: list[dict[str, object]],
    column_index: int,
    colspan: int,
) -> Optional[dict[str, object]]:
    for entry in row_entries:
        if entry["column_index"] == column_index and entry["colspan"] == colspan:
            return entry
    return None


def _render_inline(
    node: ET.Element,
    relationships: Dict[str, Dict[str, str]],
    media: Dict[str, bytes],
    assets_dir: Optional[Path] = None,
) -> str:
    tag = _local_name(node.tag)

    if tag == "hyperlink":
        link_target = None
        rel_id = node.get(f"{R_NS}id")
        if rel_id:
            relationship = relationships.get(rel_id)
            if relationship:
                link_target = relationship.get("target")

        inner = _join_html_fragments(
            _render_inline(child, relationships, media, assets_dir) for child in node if _local_name(child.tag) == "r"
        )
        if not inner:
            return ""
        if link_target:
            return f'<a href="{html.escape(link_target, quote=True)}">{inner}</a>'
        return inner

    if tag != "r":
        return ""

    prefix = []
    suffix = []
    run_properties = node.find(f"{W_NS}rPr")
    run_style = {}
    if run_properties is not None:
        if run_properties.find(f"{W_NS}b") is not None:
            prefix.append("<strong>")
            suffix.insert(0, "</strong>")
        if run_properties.find(f"{W_NS}i") is not None:
            prefix.append("<em>")
            suffix.insert(0, "</em>")
        if run_properties.find(f"{W_NS}u") is not None:
            prefix.append("<u>")
            suffix.insert(0, "</u>")

        # extract run-level style metadata (Phase 1)
        run_style = _extract_run_style(run_properties)

    parts = []
    for child in node:
        child_tag = _local_name(child.tag)
        if child_tag == "t":
            preserve_spaces = child.get("{http://www.w3.org/XML/1998/namespace}space") == "preserve"
            parts.append(_escape_html_text(child.text or "", preserve_spaces=preserve_spaces))
        elif child_tag == "tab":
            parts.append("&emsp;")
        elif child_tag in {"br", "cr"}:
            parts.append("<br>")
        elif child_tag == "drawing":
            parts.append(_render_drawing(child, relationships, media, assets_dir))

    content = "".join(parts)
    if not content:
        return ""

    # If run-level metadata exists, wrap with a span and apply inline CSS (Phase 2)
    span_prefix = ""
    span_suffix = ""
    if run_style:
        style_attr = html.escape(json.dumps(run_style, separators=(",", ":"), ensure_ascii=False), quote=True)
        css = _run_style_to_css(run_style)
        css_attr = f' style="{html.escape(css, quote=True)}"' if css else ""
        span_prefix = f'<span data-docx-style="{style_attr}"{css_attr}>'
        span_suffix = "</span>"

    return "".join(prefix) + span_prefix + content + span_suffix + "".join(suffix)


def _render_drawing(
    drawing: ET.Element,
    relationships: Dict[str, Dict[str, str]],
    media: Dict[str, bytes],
    assets_dir: Optional[Path] = None,
) -> str:
    # Try DrawingML blip first
    embed = None
    blip = drawing.find(f".//{A_NS}blip")
    if blip is not None:
        embed = blip.get(f"{R_NS}embed")

    # Fallback: VML <imagedata r:id="rIdX"> (common in older DOCX)
    if not embed:
        vml = drawing.find('.//{urn:schemas-microsoft-com:vml}imagedata')
        if vml is not None:
            # r:id attribute is stored with the R namespace
            embed = vml.get(f"{R_NS}id") or vml.get(f"{R_NS}embed") or vml.get("id")

    if not embed:
        return ""

    relationship = relationships.get(embed)
    if not relationship:
        return ""

    target = relationship.get("target", "")

    media_path, image_bytes = _find_media_bytes(media, target)
    if image_bytes is None:
        return ""

    mime_type = mimetypes.guess_type(media_path)[0] or "application/octet-stream"
    width, height = _extract_image_dimensions(drawing)
    image_src = _export_image_asset(assets_dir, media_path, image_bytes) if assets_dir is not None else _encode_image_data_uri(mime_type, image_bytes)
    dimension_attrs = []
    if width is not None:
        dimension_attrs.append(f' width="{width}"')
    if height is not None:
        dimension_attrs.append(f' height="{height}"')

    # Extract anchor positioning to prevent paragraph alignment from affecting absolute-positioned images
    style_overrides = ""
    WP_NS = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
    anchor = drawing.find(f".//{WP_NS}anchor")
    if anchor is not None:
        positionH = anchor.find(f"{WP_NS}positionH")
        positionV = anchor.find(f"{WP_NS}positionV")
        
        if positionH is not None:
            relativeFrom_h = positionH.get("relativeFrom", "")
            posOffset_h_elem = positionH.find(f"{WP_NS}posOffset")
            
            css_parts = []
            
            # For absolutely-positioned images (margin/page), apply absolute CSS positioning
            if relativeFrom_h in ("margin", "page") and posOffset_h_elem is not None:
                try:
                    posOffset_h = int(posOffset_h_elem.text or 0)
                    # Convert EMUs to inches: 914400 EMUs = 1 inch
                    # For far-right positioned images (> 6 million EMUs = 6.5+ inches),
                    # reposition to left side to avoid right-alignment issues
                    if posOffset_h > 6000000:
                        # Reposition to left side of page
                        css_parts.append("position: absolute")
                        css_parts.append("left: 0.5in")
                        css_parts.append("top: auto")
                    else:
                        # Reposition based on anchor coordinates
                        inches_h = posOffset_h / 914400
                        css_parts.append("position: absolute")
                        css_parts.append(f"left: {inches_h:.2f}in")
                        css_parts.append("top: auto")
                    
                    # Ensure the image doesn't inherit paragraph text alignment
                    css_parts.append("float: none")
                    css_parts.append("display: block")
                    
                    style_overrides = f' style="{"; ".join(css_parts)};"'
                except (ValueError, TypeError):
                    # If parsing fails, fall back to simple override
                    style_overrides = ' style="float: none; display: block;"'

    dimension_suffix = "".join(dimension_attrs)
    return f'<img class="docx-image" decoding="async" loading="eager" src="{html.escape(image_src, quote=True)}" alt=""{dimension_suffix}{style_overrides} />'


def _extract_image_dimensions(drawing: ET.Element) -> tuple[Optional[int], Optional[int]]:
    for element in drawing.iter():
        if _local_name(element.tag) not in {"extent", "ext"}:
            continue

        cx = element.get("cx")
        cy = element.get("cy")
        if not cx or not cy:
            return None, None

        try:
            width = max(1, round(int(cx) / 914400 * 96))
            height = max(1, round(int(cy) / 914400 * 96))
        except ValueError:
            return None, None

        return width, height

    return None, None


def _parse_table_column_widths(table: ET.Element) -> list[Optional[int]]:
    tbl_grid = table.find(f"{W_NS}tblGrid")
    if tbl_grid is None:
        return []

    widths = []
    for grid_col in tbl_grid.findall(f"{W_NS}gridCol"):
        widths.append(_twips_to_px(grid_col.get(f"{W_NS}w")))
    return widths


def _parse_cell_grid_span(cell: ET.Element) -> int:
    grid_span = cell.find(f"{W_NS}tcPr/{W_NS}gridSpan")
    if grid_span is None:
        return 1

    value = grid_span.get(f"{W_NS}val")
    if not value:
        return 1

    try:
        return max(1, int(value))
    except ValueError:
        return 1


def _parse_vertical_merge(cell: ET.Element) -> str:
    vmerge = cell.find(f"{W_NS}tcPr/{W_NS}vMerge")
    if vmerge is None:
        return ""

    value = vmerge.get(f"{W_NS}val")
    if value == "restart":
        return "restart"
    if value == "continue" or value is None:
        return "continue"
    return "continue"


def _parse_cell_width(cell: ET.Element) -> Optional[int]:
    width = cell.find(f"{W_NS}tcPr/{W_NS}tcW")
    if width is None:
        return None

    value = width.get(f"{W_NS}w")
    return _twips_to_px(value)


def _twips_to_px(value: Optional[str]) -> Optional[int]:
    if not value:
        return None

    try:
        return max(1, round(int(value) / 15))
    except ValueError:
        return None


def _export_image_asset(assets_dir: Path, media_path: str, image_bytes: bytes) -> str:
    asset_path = assets_dir / media_path
    asset_path.parent.mkdir(parents=True, exist_ok=True)
    if not asset_path.exists() or asset_path.read_bytes() != image_bytes:
        asset_path.write_bytes(image_bytes)
    return os.path.relpath(asset_path, start=assets_dir.parent).replace(os.sep, "/")


def _encode_image_data_uri(mime_type: str, image_bytes: bytes) -> str:
    encoded = base64.b64encode(image_bytes).decode("ascii")
    return f"data:{mime_type};base64,{encoded}"


def _resolve_media_path(target: str) -> str:
    """Normalize relationship target to a predictable archive media path.

    Handles targets like "media/image1.png", "../media/image1.png",
    "/media/image1.png" and already-prefixed "word/media/...".
    """
    if not target:
        return target

    # strip any leading slashes and parent-directory segments
    t = target.lstrip("/")
    while t.startswith("../"):
        t = t[3:]

    if t.startswith("word/"):
        return t

    return f"word/{t}"


def _wrap_html(body: str) -> str:
    return (
        "<!DOCTYPE html>\n"
        "<html lang=\"en\">\n"
        "<head>\n"
        "<meta charset=\"utf-8\">\n"
        "<meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">\n"
        "<title>DOCX Preview</title>\n"
        "</head>\n"
        "<body>\n"
        f"{body}\n"
        "</body>\n"
        "</html>\n"
    )


def _write_file_bytes(path: Path, data: bytes) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(data)


def _write_file_text(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8")


def _cli_main(argv: Optional[list[str]] = None) -> int:
    import argparse

    parser = argparse.ArgumentParser(description="Convert DOCX to HTML and PDF (A4)")
    parser.add_argument("docx", help="Path to input DOCX file")
    parser.add_argument("--out-html", "-o", help="Output HTML file path. Defaults to <input>.html")
    parser.add_argument("--out-pdf", "-p", help="Output PDF file path. Defaults to <input>.pdf")
    parser.add_argument("--assets-dir", "-a", help="Optional assets directory to extract images into")
    parser.add_argument("--validate-visual", action="store_true", help="Run Phase 4 visual validation (structural checks)")
    parser.add_argument("--reference", help="Optional reference HTML file to compare against")

    args = parser.parse_args(argv)
    docx_path = Path(args.docx)
    if not docx_path.exists():
        print(f"DOCX not found: {docx_path}", file=sys.stderr)
        return 2

    out_html = Path(args.out_html) if args.out_html else docx_path.with_suffix(".html")
    out_pdf = Path(args.out_pdf) if args.out_pdf else docx_path.with_suffix(".pdf")
    assets_dir = Path(args.assets_dir) if args.assets_dir else None

    try:
        if assets_dir is not None:
            html_doc = convert_docx_to_a4_html_with_assets(str(docx_path), assets_dir)
        else:
            html_doc = convert_docx_to_a4_html(str(docx_path))

        _write_file_text(out_html, html_doc)
        print(f"Wrote HTML: {out_html}")

        if args.validate_visual:
            report = _validate_visual(str(docx_path), html_doc, args.reference)
            report_path = out_html.with_suffix(out_html.suffix + ".validation.json")
            _write_file_text(report_path, json.dumps(report, indent=2, ensure_ascii=False))
            print(f"Wrote validation report: {report_path}")

        # produce PDF from the HTML we just generated so assets are consistent
        pdf_bytes = convert_html_to_pdf(html_doc, pdf_path=None)
        _write_file_bytes(out_pdf, pdf_bytes)
        print(f"Wrote PDF: {out_pdf}")
    except Exception as exc:
        print(f"Conversion failed: {exc}", file=sys.stderr)
        return 3

    return 0





def _local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _extract_normalization_source(root: "_Node") -> list["_Node | str"]:
    for child in root.children:
        if isinstance(child, _Node) and child.tag == "html":
            for html_child in child.children:
                if isinstance(html_child, _Node) and html_child.tag == "body":
                    return html_child.children
            return child.children
    return root.children


def _join_html_fragments(fragments: Iterable[str]) -> str:
    rendered = []
    previous_fragment = ""
    for fragment in fragments:
        if not fragment:
            continue
        if previous_fragment and _needs_inter_fragment_space(previous_fragment, fragment):
            rendered.append(" ")
        rendered.append(fragment)
        previous_fragment = fragment
    return "".join(rendered)


def _needs_inter_fragment_space(previous_fragment: str, next_fragment: str) -> bool:
    previous_text = html.unescape(re.sub(r"<[^>]+>", "", previous_fragment)).strip()
    next_text = html.unescape(re.sub(r"<[^>]+>", "", next_fragment)).strip()
    if not previous_text or not next_text:
        return False
    if previous_fragment.endswith(" ") or next_fragment.startswith(" "):
        return False

    return previous_text[-1].isalnum() and next_text[0].isalnum()


def _escape_html_text(text: str, preserve_spaces: bool = False) -> str:
    escaped = html.escape(text, quote=False)
    if preserve_spaces:
        return escaped.replace(" ", "&nbsp;")
    return escaped


def _normalize_children(children: list["_Node | str"]) -> list["_Node | str"]:
    normalized = []
    for child in children:
        if isinstance(child, str):
            normalized.append(child)
            continue

        child.children = _normalize_children(child.children)
        child.attrs = _normalize_attrs(child)

        if child.tag == "table":
            child.attrs["class"] = "doc-table"

        if child.tag not in VOID_HTML_TAGS and child.tag not in PRESERVE_EMPTY_HTML_TAGS and not _node_has_renderable_content(child):
            continue

        normalized.append(child)

    return normalized


def _trim_root_edge_spacers(children: list["_Node | str"]) -> list["_Node | str"]:
    start = 0
    end = len(children)

    while start < end and _is_empty_spacer_paragraph(children[start]):
        start += 1

    while end > start and _is_empty_spacer_paragraph(children[end - 1]):
        end -= 1

    return children[start:end]


def _is_empty_spacer_paragraph(node: "_Node | str") -> bool:
    if isinstance(node, str):
        return not node.strip()

    if node.tag != "p":
        return False

    class_names = set(node.attrs.get("class", "").split())
    if "docx-empty-paragraph" in class_names:
        return True

    text_content = []
    for child in node.children:
        if isinstance(child, str):
            text_content.append(html.unescape(child))

    return not "".join(text_content).replace("\u00a0", "").strip()


def _normalize_attrs(node: "_Node") -> dict[str, str]:
    normalized = {}
    for key, value in node.attrs.items():
        if key == "style":
            continue
        if key == "class" and node.tag == "table":
            continue
        normalized[key] = value
    return normalized


def _node_has_renderable_content(node: "_Node") -> bool:
    if node.tag in VOID_HTML_TAGS:
        return True

    if node.tag in PRESERVE_EMPTY_HTML_TAGS:
        return True

    if node.text.strip():
        return True

    for child in node.children:
        if isinstance(child, str):
            if child.strip():
                return True
        elif _node_has_renderable_content(child):
            return True

    return False


def _serialize_children(children: list["_Node | str"]) -> str:
    return "".join(_serialize_node(child) for child in children)


def _serialize_node(node: "_Node | str") -> str:
    if isinstance(node, str):
        return html.escape(node)

    attrs = "".join(f' {key}="{html.escape(value, quote=True)}"' for key, value in sorted(node.attrs.items()))
    if node.tag in VOID_HTML_TAGS:
        return f"<{node.tag}{attrs}>"

    inner = html.escape(node.text) if node.text else ""
    inner += _serialize_children(node.children)
    return f"<{node.tag}{attrs}>{inner}</{node.tag}>"


class _Node:
    def __init__(self, tag: str, attrs: Optional[dict[str, str]] = None, children: Optional[list["_Node | str"]] = None) -> None:
        self.tag = tag
        self.attrs = attrs or {}
        self.children = children or []
        self.text = ""


class _HtmlNormalizer(HTMLParser):
    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.root = _Node("root")
        self._stack = [self.root]

    def handle_starttag(self, tag: str, attrs: list[tuple[str, Optional[str]]]) -> None:
        node = _Node(tag.lower(), attrs={key.lower(): value or "" for key, value in attrs})
        self._stack[-1].children.append(node)
        if tag.lower() not in VOID_HTML_TAGS:
            self._stack.append(node)

    def handle_endtag(self, tag: str) -> None:
        tag = tag.lower()
        for index in range(len(self._stack) - 1, 0, -1):
            if self._stack[index].tag == tag:
                self._stack = self._stack[:index]
                break

    def handle_data(self, data: str) -> None:
        if data:
            self._stack[-1].children.append(data)


def main(argv: Optional[Iterable[str]] = None) -> int:
    args = list(argv if argv is not None else sys.argv[1:])
    if len(args) != 1:
        print("Usage: python docx_to_html.py input.docx", file=sys.stderr)
        return 2

    try:
        html_output = convert_docx_to_html(args[0])
    except DocxConversionError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1
    
    try:
        pdf_output = convert_html_to_pdf(html_output)
        Path("output.pdf").write_bytes(pdf_output)
    except PdfConversionError as exc:
        print(f"PDF conversion warning: {exc}", file=sys.stderr)

    sys.stdout.write(html_output)
    return 0


if __name__ == "__main__":
    raise SystemExit(_cli_main())