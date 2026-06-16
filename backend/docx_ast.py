"""Canonical AST builder for DOCX documents.

Parses a .docx archive and returns a JSON-serialisable dict conforming to
docx_ast_schema.json (version 1.0).  No HTML is produced here.
"""

from __future__ import annotations

import base64
import mimetypes
import os
import zipfile
from pathlib import Path
from typing import Dict, List, Optional
from xml.etree import ElementTree as ET

try:
    from .docx_node_ids import NODE_ID_ATTR, inject_docx_node_ids
except ImportError:
    from docx_node_ids import NODE_ID_ATTR, inject_docx_node_ids

W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
R_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
A_NS = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
WP_NS = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
VML_NS = "{urn:schemas-microsoft-com:vml}"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
W14_NS = "{http://schemas.microsoft.com/office/word/2010/wordml}"


class DocxAstError(RuntimeError):
    pass


_ORDERED_NUM_FMTS = frozenset({
    "decimal", "upperRoman", "lowerRoman", "upperLetter", "lowerLetter",
    "decimalZero", "ordinal", "cardinalText", "ordinalText",
    "hex", "decimalEnclosedCircle", "decimalFullWidth", "decimalHalfWidth",
    "decimalFullWidth2",
})

_NUM_FMT_TO_CSS = {
    "decimal": "decimal",
    "upperRoman": "upper-roman",
    "lowerRoman": "lower-roman",
    "upperLetter": "upper-alpha",
    "lowerLetter": "lower-alpha",
    "decimalZero": "decimal-leading-zero",
}

_BULLET_CHAR_TO_MARKER = {
    "•": "disc",    # •
    "‣": "disc",    # ‣
    "◦": "circle",  # ◦
    "○": "circle",  # ○
    "▪": "square",  # ▪
    "■": "square",  # ■
    "●": "disc",    # ●
}


# ---------------------------------------------------------------------------
# Public entry point
# ---------------------------------------------------------------------------

def build_docx_ast(docx_path: str | Path) -> dict:
    """Parse *docx_path* and return its canonical AST dict."""
    path = Path(docx_path)
    if not path.is_file():
        raise DocxAstError(f"DOCX file not found: {path}")

    try:
        with zipfile.ZipFile(path) as archive:
            document_xml = _read_xml(archive, "word/document.xml")
            styles_root = _read_xml(archive, "word/styles.xml")
            styles = _parse_styles(styles_root)
            table_styles = _parse_table_styles(styles_root)
            relationships = _parse_relationships(
                _read_xml(archive, "word/_rels/document.xml.rels")
            )
            media = _load_media(archive)

            try:
                numbering = _parse_numbering(_read_xml(archive, "word/numbering.xml"))
            except KeyError:
                numbering = {}

            # Resolve theme body font (minorHAnsi) so the renderer can apply
            # accurate font-metrics correction to CSS line-height.
            default_body_font: str = ""
            try:
                theme_xml = _read_xml(archive, "word/theme/theme1.xml")
                _THEME_NS = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
                minor_latin = theme_xml.find(
                    f".//{_THEME_NS}minorFont/{_THEME_NS}latin"
                )
                if minor_latin is not None:
                    default_body_font = minor_latin.get("typeface", "")
            except (KeyError, Exception):
                pass

            id_counters = {"p": 0, "r": 0, "sdt": 0}
            inject_docx_node_ids(document_xml, id_counters)

            header_parts: list[tuple[ET.Element, dict]] = []
            footer_parts: list[tuple[ET.Element, dict]] = []
            for rel in relationships.values():
                rtype = rel.get("type", "")
                target = rel.get("target", "")
                if not target or not rtype:
                    continue
                is_header = rtype.endswith("/header") or rtype.endswith("/headers")
                is_footer = rtype.endswith("/footer") or rtype.endswith("/footers")
                if not is_header and not is_footer:
                    continue
                part_path = (
                    f"word/{target}" if not target.startswith("word/") else target
                )
                try:
                    part_xml = _read_xml(archive, part_path)
                    inject_docx_node_ids(part_xml, id_counters)
                    part_dir, part_filename = part_path.rsplit("/", 1)
                    rels_path = f"{part_dir}/_rels/{part_filename}.rels"
                    try:
                        part_rels = _parse_relationships(_read_xml(archive, rels_path))
                    except KeyError:
                        part_rels = {}
                    if is_header:
                        header_parts.append((part_xml, part_rels))
                    else:
                        footer_parts.append((part_xml, part_rels))
                except KeyError:
                    pass
    except zipfile.BadZipFile as exc:
        raise DocxAstError(f"Invalid DOCX archive: {path}") from exc
    except KeyError as exc:
        raise DocxAstError(f"Missing required DOCX part: {exc.args[0]}") from exc

    body_elem = document_xml.find(f"{W_NS}body")
    if body_elem is None:
        raise DocxAstError("DOCX document body is missing")

    sections: list[dict] = []
    for hdr, hdr_rels in header_parts:
        sections.append({"kind": "header", "blocks": _build_blocks(hdr, styles, hdr_rels, media, table_styles)})
    sections.append({"kind": "body", "blocks": _build_blocks(body_elem, styles, relationships, media, table_styles)})
    for ftr, ftr_rels in footer_parts:
        sections.append({"kind": "footer", "blocks": _build_blocks(ftr, styles, ftr_rels, media, table_styles)})

    images = {
        path_key: {
            "media_path": path_key,
            "mime_type": mimetypes.guess_type(path_key)[0] or "application/octet-stream",
            "data_uri": "data:{};base64,{}".format(
                mimetypes.guess_type(path_key)[0] or "application/octet-stream",
                base64.b64encode(img_bytes).decode("ascii"),
            ),
        }
        for path_key, img_bytes in media.items()
    }

    page_setup = _extract_page_setup(body_elem, styles_root, default_body_font)

    result = {
        "version": "1.0",
        "sections": sections,
        "relationships": relationships,
        "images": images,
        "page_setup": page_setup,
    }
    _annotate_list_paragraphs(result, numbering)
    return result


# ---------------------------------------------------------------------------
# Block builders
# ---------------------------------------------------------------------------

def _build_blocks(
    root: ET.Element,
    styles: dict,
    relationships: dict,
    media: dict,
    table_styles: dict | None = None,
) -> list[dict]:
    blocks: list[dict] = []
    for child in root:
        tag = _local_name(child.tag)
        if tag == "p":
            block = _build_paragraph(child, styles, relationships, media)
        elif tag == "tbl":
            block = _build_table(child, styles, relationships, media, table_styles)
        else:
            continue
        # Explicit w:br type="page" inside runs → split into separate pages
        if tag == "p" and not block.get("page_break_before"):
            block = _promote_explicit_page_breaks(block, blocks)
            if block is None:
                continue

        # w:pageBreakBefore paragraph property
        if block.get("style", {}).get("page_break_before"):
            block["page_break_before"] = True

        # lastRenderedPageBreak hints (fallback for reflowed content)
        if not block.get("page_break_before"):
            pb = _get_page_break_info(child)
            if pb:
                if pb.get("at_start"):
                    block["page_break_before"] = True
                else:
                    block["page_break_info"] = pb
        blocks.append(block)
    return blocks


def _promote_explicit_page_breaks(block: dict, blocks: list[dict]) -> dict | None:
    """Split a paragraph at explicit w:br type=page breaks.

    Runs before the first page break are appended to `blocks` as a completed
    paragraph. The first run after the break starts a new paragraph with
    page_break_before=True, which is returned. If the break is the very first
    thing in the paragraph the original block is returned with page_break_before
    set. Returns None only if the paragraph was entirely consumed into `blocks`.
    """
    children = block.get("children", [])
    # Flatten: collect (run_index, child_index) of every page break
    split_points: list[tuple[int, int]] = []
    for ri, run in enumerate(children):
        if run.get("type") != "run":
            continue
        for ci, ch in enumerate(run.get("children", [])):
            if ch.get("type") == "break" and ch.get("break_type") == "page":
                split_points.append((ri, ci))

    if not split_points:
        return block

    def _slice_run(run: dict, start: int, end: int | None) -> dict | None:
        inner = run.get("children", [])[start:end]
        if not inner:
            return None
        return {**run, "children": inner}

    def _build_para(runs: list[dict], base: dict) -> dict:
        return {k: v for k, v in base.items() if k != "children"} | {"children": runs}

    segments: list[list[dict]] = []
    current: list[dict] = []
    ri_prev, ci_prev = 0, 0

    for ri, ci in split_points:
        # Collect runs before this break
        for i in range(ri_prev, ri):
            run = children[i]
            if run.get("type") != "run":
                current.append(run)
                continue
            sliced = _slice_run(run, ci_prev if i == ri_prev else 0, None)
            if sliced:
                current.append(sliced)
            ci_prev = 0
        # Partial run up to (not including) the break char
        run = children[ri]
        if run.get("type") == "run":
            before = _slice_run(run, ci_prev if ri == ri_prev else 0, ci)
            if before:
                current.append(before)
        segments.append(current)
        current = []
        ri_prev, ci_prev = ri, ci + 1  # skip the break itself

    # Tail after last break
    for i in range(ri_prev, len(children)):
        run = children[i]
        if run.get("type") != "run":
            current.append(run)
            continue
        sliced = _slice_run(run, ci_prev if i == ri_prev else 0, None)
        if sliced:
            current.append(sliced)
        ci_prev = 0
    segments.append(current)

    # First segment → flush to blocks (inherits original block's page_break_before)
    first_para = _build_para(segments[0], block)
    if segments[0]:
        blocks.append(first_para)

    # Middle segments (if any multi-break)
    for seg in segments[1:-1]:
        p = _build_para(seg, block)
        p["page_break_before"] = True
        blocks.append(p)

    # Last segment → returned as the new "current" block
    last = _build_para(segments[-1], block)
    last["page_break_before"] = True
    if not segments[-1]:
        # Nothing after the last break — no block to return
        return None
    return last


def _get_page_break_info(elem: ET.Element) -> dict | None:
    """Return info about the first lastRenderedPageBreak within elem, or None."""
    tag = _local_name(elem.tag)
    if tag == "p":
        chars_before = 0
        for run in elem.findall(f"{W_NS}r"):
            children = list(run)
            for ci, child in enumerate(children):
                if _local_name(child.tag) == "lastRenderedPageBreak":
                    return {"at_start": chars_before == 0}
            for t in run.findall(f"{W_NS}t"):
                chars_before += len(t.text or "")
        return None

    if tag == "tbl":
        # Only look at direct tr children to avoid nested tables
        for ri, row in enumerate(c for c in elem if _local_name(c.tag) == "tr"):
            for ci, cell in enumerate(row.findall(f"{W_NS}tc")):
                paras = cell.findall(f"{W_NS}p")
                chars_before = 0
                for pi, para in enumerate(paras):
                    for run in para.findall(f"{W_NS}r"):
                        children = list(run)
                        for child in children:
                            if _local_name(child.tag) == "lastRenderedPageBreak":
                                return {
                                    "at_start": chars_before == 0 and pi == 0,
                                    "row": ri,
                                    "cell": ci,
                                    "para": pi,
                                }
                        for t in run.findall(f"{W_NS}t"):
                            chars_before += len(t.text or "")
        return None

    return None


def _build_paragraph(
    paragraph: ET.Element,
    styles: dict,
    relationships: dict,
    media: dict,
) -> dict:
    pstyle = _extract_paragraph_style(paragraph, styles)
    children: list[dict] = []
    for child in paragraph:
        tag = _local_name(child.tag)
        if tag == "r":
            children.append(_build_run(child, relationships, media))
        elif tag == "hyperlink":
            children.append(_build_hyperlink(child, relationships, media))
        elif tag == "sdt":
            cb = _build_checkbox(child)
            if cb is not None:
                children.append(cb)
    return {
        "type": "paragraph",
        "style": pstyle,
        "node_id": paragraph.get(NODE_ID_ATTR),
        "children": children,
    }


def _build_checkbox(sdt: ET.Element) -> dict | None:
    sdtpr = sdt.find(f"{W_NS}sdtPr")
    if sdtpr is None:
        return None
    if sdtpr.find(f"{W14_NS}checkbox") is None:
        return None
    checked_elem = sdtpr.find(f"{W14_NS}checkbox/{W14_NS}checked")
    checked = (checked_elem.get(f"{W14_NS}val") if checked_elem is not None else "0") == "1"
    return {
        "type": "run",
        "style": {},
        "node_id": sdt.get(NODE_ID_ATTR),
        "children": [{"type": "checkbox", "checked": checked}],
    }


def _extract_paragraph_style(paragraph: ET.Element, styles: dict) -> dict:
    style: dict = {}
    ppr = paragraph.find(f"{W_NS}pPr")
    if ppr is None:
        return style

    style_elem = ppr.find(f"{W_NS}pStyle")
    if style_elem is not None:
        sid = style_elem.get(f"{W_NS}val")
        if sid:
            style["style_id"] = sid
            style["style_name"] = styles.get(sid, sid)

    jc = ppr.find(f"{W_NS}jc")
    if jc is not None:
        val = jc.get(f"{W_NS}val") or jc.get("val")
        if val:
            style["alignment"] = val

    spacing_elem = ppr.find(f"{W_NS}spacing")
    if spacing_elem is not None:
        spacing: dict = {}
        for key in ("before", "after", "line"):
            v = spacing_elem.get(f"{W_NS}{key}") or spacing_elem.get(key)
            if v is not None:
                spacing[key] = v
        if spacing:
            style["spacing"] = spacing

    ind_elem = ppr.find(f"{W_NS}ind")
    if ind_elem is not None:
        indent: dict = {}
        for key in ("left", "right", "firstLine", "hanging"):
            v = ind_elem.get(f"{W_NS}{key}") or ind_elem.get(key)
            if v is not None:
                indent[key] = v
        if indent:
            style["indent"] = indent

    if ppr.find(f"{W_NS}pageBreakBefore") is not None:
        style["page_break_before"] = True

    num_pr = ppr.find(f"{W_NS}numPr")
    if num_pr is not None:
        num_id_elem = num_pr.find(f"{W_NS}numId")
        ilvl_elem = num_pr.find(f"{W_NS}ilvl")
        if num_id_elem is not None:
            v = num_id_elem.get(f"{W_NS}val") or num_id_elem.get("val")
            if v:
                style["num_id"] = v
        if ilvl_elem is not None:
            v = ilvl_elem.get(f"{W_NS}val") or ilvl_elem.get("val")
            if v is not None:
                try:
                    style["ilvl"] = int(v)
                except ValueError:
                    pass

    return style


# ---------------------------------------------------------------------------
# Inline builders
# ---------------------------------------------------------------------------

def _build_run(run: ET.Element, relationships: dict, media: dict) -> dict:
    rpr = run.find(f"{W_NS}rPr")
    run_style = _extract_run_style(rpr) if rpr is not None else {}

    children: list[dict] = []
    for child in run:
        tag = _local_name(child.tag)
        if tag == "t":
            children.append({
                "type": "text",
                "text": child.text or "",
                "preserve_spaces": child.get(XML_SPACE) == "preserve",
            })
        elif tag == "tab":
            children.append({"type": "break", "break_type": "tab"})
        elif tag in ("br", "cr"):
            br_type = child.get(f"{W_NS}type") or child.get("type")
            if br_type == "page":
                children.append({"type": "break", "break_type": "page"})
            else:
                children.append({"type": "break", "break_type": "line"})
        elif tag == "drawing":
            img = _build_image(child, relationships, media)
            if img is not None:
                children.append(img)

    return {
        "type": "run",
        "style": run_style,
        "node_id": run.get(NODE_ID_ATTR),
        "children": children,
    }


def _extract_run_style(rpr: ET.Element) -> dict:
    style: dict = {}

    sz = rpr.find(f"{W_NS}sz")
    if sz is not None:
        val = sz.get(f"{W_NS}val") or sz.get("val")
        if val:
            try:
                pts = int(val) / 2
                style["font_size_pt"] = int(pts) if float(pts).is_integer() else round(pts, 1)
            except (ValueError, TypeError):
                pass

    rfonts = rpr.find(f"{W_NS}rFonts")
    if rfonts is not None:
        for key in ("ascii", "hAnsi", "cs"):
            v = rfonts.get(f"{W_NS}{key}") or rfonts.get(key)
            if v:
                style["font_family"] = v
                break

    color = rpr.find(f"{W_NS}color")
    if color is not None:
        v = color.get(f"{W_NS}val") or color.get("val")
        if v:
            style["color"] = v

    if rpr.find(f"{W_NS}b") is not None:
        style["bold"] = True
    if rpr.find(f"{W_NS}i") is not None:
        style["italic"] = True
    if rpr.find(f"{W_NS}u") is not None:
        style["underline"] = True
    if rpr.find(f"{W_NS}strike") is not None:
        style["strike"] = True

    vert = rpr.find(f"{W_NS}vertAlign")
    if vert is not None:
        v = vert.get(f"{W_NS}val") or vert.get("val")
        if v in ("superscript", "subscript"):
            style["vertical_align"] = v

    return style


def _build_hyperlink(node: ET.Element, relationships: dict, media: dict) -> dict:
    href: Optional[str] = None
    rel_id = node.get(f"{R_NS}id")
    if rel_id:
        rel = relationships.get(rel_id)
        if rel:
            href = rel.get("target")

    children: list[dict] = []
    for child in node:
        if _local_name(child.tag) == "r":
            children.append(_build_run(child, relationships, media))

    return {"type": "hyperlink", "href": href, "children": children}


def _build_image(drawing: ET.Element, relationships: dict, media: dict) -> Optional[dict]:
    embed: Optional[str] = None

    blip = drawing.find(f".//{A_NS}blip")
    if blip is not None:
        embed = blip.get(f"{R_NS}embed")

    if not embed:
        vml = drawing.find(f".//{VML_NS}imagedata")
        if vml is not None:
            embed = vml.get(f"{R_NS}id") or vml.get(f"{R_NS}embed") or vml.get("id")

    if not embed:
        return None

    rel = relationships.get(embed)
    if not rel:
        return None

    target = rel.get("target", "")
    media_path, image_bytes = _find_media_bytes(media, target)

    mime_type: Optional[str] = None
    data_uri: Optional[str] = None
    if image_bytes is not None and media_path:
        mime_type = mimetypes.guess_type(media_path)[0] or "application/octet-stream"
        data_uri = "data:{};base64,{}".format(
            mime_type, base64.b64encode(image_bytes).decode("ascii")
        )

    width_px, height_px = _extract_image_dimensions(drawing)
    positioning = _extract_image_positioning(drawing)

    return {
        "type": "image",
        "rel_id": embed,
        "media_path": media_path,
        "data_uri": data_uri,
        "mime_type": mime_type,
        "width_px": width_px,
        "height_px": height_px,
        "positioning": positioning,
    }


def _extract_image_dimensions(drawing: ET.Element) -> tuple[Optional[int], Optional[int]]:
    for element in drawing.iter():
        if _local_name(element.tag) not in {"extent", "ext"}:
            continue
        cx = element.get("cx")
        cy = element.get("cy")
        if not cx or not cy:
            return None, None
        try:
            return (
                max(1, round(int(cx) / 914400 * 96)),
                max(1, round(int(cy) / 914400 * 96)),
            )
        except ValueError:
            return None, None
    return None, None


def _extract_image_positioning(drawing: ET.Element) -> Optional[dict]:
    anchor = drawing.find(f".//{WP_NS}anchor")
    if anchor is not None:
        pos: dict = {"kind": "anchor"}
        pos_h = anchor.find(f"{WP_NS}positionH")
        if pos_h is not None:
            pos["relative_from_h"] = pos_h.get("relativeFrom", "")
            off = pos_h.find(f"{WP_NS}posOffset")
            if off is not None and off.text:
                try:
                    pos["offset_h_emu"] = int(off.text)
                except ValueError:
                    pass
        pos_v = anchor.find(f"{WP_NS}positionV")
        if pos_v is not None:
            pos["relative_from_v"] = pos_v.get("relativeFrom", "")
            off = pos_v.find(f"{WP_NS}posOffset")
            if off is not None and off.text:
                try:
                    pos["offset_v_emu"] = int(off.text)
                except ValueError:
                    pass
        return pos

    if drawing.find(f".//{WP_NS}inline") is not None:
        return {"kind": "inline"}

    return None


# ---------------------------------------------------------------------------
# Table builders
# ---------------------------------------------------------------------------

def _build_table(
    table: ET.Element,
    styles: dict,
    relationships: dict,
    media: dict,
    table_styles: dict | None = None,
) -> dict:
    column_widths = _parse_column_widths(table)

    tbl_pr = table.find(f"{W_NS}tblPr")
    tbl_style_id: Optional[str] = None
    if tbl_pr is not None:
        ts_elem = tbl_pr.find(f"{W_NS}tblStyle")
        if ts_elem is not None:
            tbl_style_id = ts_elem.get(f"{W_NS}val")

    resolved = _resolve_tbl_style(tbl_style_id, table_styles or {})
    # Direct tblBorders on the element override style defaults
    tbl_borders = dict(resolved.get("borders", {}))
    tbl_cell_margins = dict(resolved.get("cell_margins", {}))
    if tbl_pr is not None:
        direct_b = _parse_tbl_borders_elem(tbl_pr)
        tbl_borders.update(direct_b)
        direct_m = _parse_tbl_cell_margins_elem(tbl_pr)
        tbl_cell_margins.update(direct_m)

    rows = _build_table_rows(table, styles, relationships, media, table_styles)
    return {
        "type": "table",
        "column_widths_twips": column_widths,
        "tbl_borders": tbl_borders,
        "tbl_cell_margins": tbl_cell_margins,
        "rows": rows,
    }


def _parse_tbl_borders_elem(tbl_pr: ET.Element) -> dict:
    """Extract tblBorders from a tblPr element."""
    result: dict = {}
    borders_elem = tbl_pr.find(f"{W_NS}tblBorders")
    if borders_elem is None:
        return result
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        b = borders_elem.find(f"{W_NS}{edge}")
        if b is None:
            continue
        val = b.get(f"{W_NS}val") or b.get("val", "")
        if val in ("none", "nil", ""):
            continue
        sz = b.get(f"{W_NS}sz") or b.get("sz")
        color = b.get(f"{W_NS}color") or b.get("color", "auto")
        result[edge] = {
            "style": "double" if val == "double" else "solid",
            "width_pt": round(int(sz) / 8, 2) if sz else 0.5,
            "color": color.upper() if color and color.upper() != "AUTO" else "000000",
        }
    return result


def _parse_tbl_cell_margins_elem(tbl_pr: ET.Element) -> dict:
    """Extract tblCellMar default cell margins (twips) from a tblPr element."""
    result: dict = {}
    tcmar = tbl_pr.find(f"{W_NS}tblCellMar")
    if tcmar is None:
        return result
    for edge in ("top", "left", "bottom", "right"):
        ee = tcmar.find(f"{W_NS}{edge}")
        if ee is not None:
            w_val = ee.get(f"{W_NS}w") or ee.get("w")
            if w_val:
                try:
                    result[edge] = int(w_val)
                except ValueError:
                    pass
    return result


def _parse_table_styles(styles_root: ET.Element) -> dict:
    """Parse table styles from styles.xml into {styleId: {based_on, borders, cell_margins}}."""
    raw: dict = {}
    for style in styles_root.findall(f"{W_NS}style"):
        if style.get(f"{W_NS}type") != "table":
            continue
        sid = style.get(f"{W_NS}styleId")
        if not sid:
            continue
        based = style.find(f"{W_NS}basedOn")
        based_on = based.get(f"{W_NS}val") if based is not None else None
        tbl_pr = style.find(f"{W_NS}tblPr")
        borders: dict = {}
        cell_margins: dict = {}
        if tbl_pr is not None:
            borders = _parse_tbl_borders_elem(tbl_pr)
            cell_margins = _parse_tbl_cell_margins_elem(tbl_pr)
        raw[sid] = {"based_on": based_on, "borders": borders, "cell_margins": cell_margins}
    return raw


def _resolve_tbl_style(style_id: Optional[str], table_styles: dict) -> dict:
    """Walk basedOn chain and return merged {borders, cell_margins}."""
    chain: list[dict] = []
    visited: set[str] = set()
    current = style_id
    while current and current not in visited:
        visited.add(current)
        entry = table_styles.get(current)
        if entry is None:
            break
        chain.append(entry)
        current = entry.get("based_on")
    borders: dict = {}
    cell_margins: dict = {}
    for entry in reversed(chain):
        borders = {**borders, **entry.get("borders", {})}
        cell_margins = {**cell_margins, **entry.get("cell_margins", {})}
    return {"borders": borders, "cell_margins": cell_margins}


def _parse_column_widths(table: ET.Element) -> list:
    tbl_grid = table.find(f"{W_NS}tblGrid")
    if tbl_grid is None:
        return []
    widths: list = []
    for gc in tbl_grid.findall(f"{W_NS}gridCol"):
        val = gc.get(f"{W_NS}w")
        if val:
            try:
                widths.append(int(val))
            except ValueError:
                widths.append(None)
        else:
            widths.append(None)
    return widths


def _build_table_rows(
    table: ET.Element,
    styles: dict,
    relationships: dict,
    media: dict,
    table_styles: dict | None = None,
) -> list[dict]:
    # First pass: collect raw cell data with vmerge tracking
    raw_rows: list[list[dict]] = []
    for row_index, row in enumerate(table.findall(f"{W_NS}tr")):
        height_twips: Optional[int] = None
        trpr = row.find(f"{W_NS}trPr")
        if trpr is not None:
            tr_h = trpr.find(f"{W_NS}trHeight")
            if tr_h is not None:
                val = tr_h.get(f"{W_NS}val") or tr_h.get("val")
                if val:
                    try:
                        height_twips = int(val)
                    except ValueError:
                        pass

        entries: list[dict] = []
        current_col = 0
        for cell in row.findall(f"{W_NS}tc"):
            colspan = _parse_cell_colspan(cell)
            vmerge = _parse_cell_vmerge(cell)
            width_twips = _parse_cell_width_twips(cell)
            cell_style = _extract_cell_style(cell)
            children = _build_blocks(cell, styles, relationships, media, table_styles)
            entries.append({
                "_col": current_col,
                "colspan": colspan,
                "rowspan": 1,
                "_vmerge": vmerge,
                "width_twips": width_twips,
                "_height_twips": height_twips,
                "style": cell_style,
                "children": children,
                "_skip": False,
            })
            current_col += colspan
        raw_rows.append(entries)

    # Second pass: resolve rowspan from vmerge restart/continue pairs
    for row_index, entries in enumerate(raw_rows):
        for entry in entries:
            if entry["_vmerge"] != "restart":
                continue
            rowspan = 1
            for next_row in raw_rows[row_index + 1:]:
                match = next(
                    (e for e in next_row
                     if e["_col"] == entry["_col"] and e["colspan"] == entry["colspan"]),
                    None,
                )
                if match is None or match["_vmerge"] != "continue":
                    break
                match["_skip"] = True
                rowspan += 1
            entry["rowspan"] = rowspan

        for entry in entries:
            if entry["_vmerge"] == "continue":
                entry["_skip"] = True

    # Third pass: emit clean rows, stripping internal tracking keys
    rows: list[dict] = []
    for entries in raw_rows:
        cells: list[dict] = []
        for entry in entries:
            if entry["_skip"]:
                continue
            cells.append({
                "colspan": entry["colspan"],
                "rowspan": entry["rowspan"],
                "vmerge": entry["_vmerge"],
                "width_twips": entry["width_twips"],
                "style": entry["style"],
                "children": entry["children"],
            })
        height = entries[0]["_height_twips"] if entries else None
        rows.append({"height_twips": height, "cells": cells})
    return rows


def _parse_cell_colspan(cell: ET.Element) -> int:
    gs = cell.find(f"{W_NS}tcPr/{W_NS}gridSpan")
    if gs is None:
        return 1
    val = gs.get(f"{W_NS}val")
    if not val:
        return 1
    try:
        return max(1, int(val))
    except ValueError:
        return 1


def _parse_cell_vmerge(cell: ET.Element) -> str:
    vm = cell.find(f"{W_NS}tcPr/{W_NS}vMerge")
    if vm is None:
        return ""
    val = vm.get(f"{W_NS}val")
    return "restart" if val == "restart" else "continue"


def _parse_cell_width_twips(cell: ET.Element) -> Optional[int]:
    w = cell.find(f"{W_NS}tcPr/{W_NS}tcW")
    if w is None:
        return None
    val = w.get(f"{W_NS}w")
    if not val:
        return None
    try:
        return int(val)
    except ValueError:
        return None


def _extract_cell_style(cell: ET.Element) -> dict:
    tcpr = cell.find(f"{W_NS}tcPr")
    style: dict = {}
    if tcpr is None:
        return style

    valign = tcpr.find(f"{W_NS}vAlign")
    if valign is not None:
        v = valign.get(f"{W_NS}val") or valign.get("val")
        if v:
            style["vertical_align"] = v

    jc = tcpr.find(f"{W_NS}jc")
    if jc is not None:
        v = jc.get(f"{W_NS}val") or jc.get("val")
        if v:
            style["alignment"] = v

    if "alignment" not in style:
        first_p = cell.find(f"{W_NS}p")
        if first_p is not None:
            pjc = first_p.find(f"{W_NS}pPr/{W_NS}jc")
            if pjc is not None:
                v = pjc.get(f"{W_NS}val") or pjc.get("val")
                if v:
                    style["alignment"] = v

    # Background shading (w:shd w:fill — "auto" means no fill)
    shd = tcpr.find(f"{W_NS}shd")
    if shd is not None:
        fill = shd.get(f"{W_NS}fill") or shd.get("fill")
        if fill and fill.upper() != "AUTO":
            style["background"] = fill.upper()

    # Per-edge cell borders (w:tcBorders; sz is in eighths of a point)
    tc_borders = tcpr.find(f"{W_NS}tcBorders")
    if tc_borders is not None:
        borders: dict = {}
        for edge in ("top", "left", "bottom", "right"):
            b = tc_borders.find(f"{W_NS}{edge}")
            if b is None:
                continue
            val = b.get(f"{W_NS}val") or b.get("val", "")
            if val in ("none", "nil", ""):
                continue
            sz = b.get(f"{W_NS}sz") or b.get("sz")
            color = b.get(f"{W_NS}color") or b.get("color", "auto")
            borders[edge] = {
                "style": "double" if val == "double" else "solid",
                "width_pt": round(int(sz) / 8, 2) if sz else 0.5,
                "color": color.upper() if color and color.upper() != "AUTO" else "000000",
            }
        if borders:
            style["borders"] = borders

    return style


# ---------------------------------------------------------------------------
# Page setup extraction
# ---------------------------------------------------------------------------

# Win-metrics ratio (winAscent+winDescent)/UPM for common fonts.
# Word's lineRule="auto" uses win metrics as the line pitch base; CSS uses em (font-size).
# The ratio corrects CSS line-height to match Word's rendered line height.
_FONT_WIN_RATIO: dict[str, float] = {
    "calibri":        1.328,  # (2206+514)/2048
    "calibri light":  1.328,
    "aptos":          1.328,  # successor to Calibri, same design metrics
    "aptos display":  1.328,
    "aptos narrow":   1.328,
    "arial":          1.275,  # (2112+497)/2048
    "arial narrow":   1.275,
    "times new roman":1.275,
    "segoe ui":       1.329,  # (2210+514)/2048
    "cambria":        1.328,
    "georgia":        1.300,
    "verdana":        1.261,
}
_FONT_WIN_RATIO_DEFAULT = 1.3  # safe fallback for unknown fonts


def _extract_page_setup(
    body: ET.Element,
    styles_root: ET.Element | None = None,
    default_body_font: str = "",
) -> dict:
    """Extract page dimensions, margins, and default typography from the document."""
    sect = body.find(f"{W_NS}sectPr")
    result: dict = {}

    if sect is not None:
        pg_sz = sect.find(f"{W_NS}pgSz")
        if pg_sz is not None:
            w = pg_sz.get(f"{W_NS}w")
            h = pg_sz.get(f"{W_NS}h")
            if w and h:
                try:
                    result["width_in"] = round(int(w) / 1440, 4)
                    result["height_in"] = round(int(h) / 1440, 4)
                except ValueError:
                    pass

        pg_mar = sect.find(f"{W_NS}pgMar")
        if pg_mar is not None:
            margins: dict = {}
            for edge in ("top", "right", "bottom", "left"):
                v = pg_mar.get(f"{W_NS}{edge}")
                if v:
                    try:
                        margins[edge] = round(int(v) / 1440, 4)
                    except ValueError:
                        pass
            if margins:
                result["margins_in"] = margins

    if styles_root is not None:
        dd = styles_root.find(f"{W_NS}docDefaults/{W_NS}rPrDefault/{W_NS}rPr")
        if dd is not None:
            sz = dd.find(f"{W_NS}sz")
            if sz is not None:
                v = sz.get(f"{W_NS}val") or sz.get("val")
                if v:
                    try:
                        result["default_font_size_pt"] = int(v) / 2
                    except ValueError:
                        pass
        # Extract default paragraph spacing and line height from docDefaults
        dd_ppr = styles_root.find(f"{W_NS}docDefaults/{W_NS}pPrDefault/{W_NS}pPr")
        if dd_ppr is not None:
            spacing = dd_ppr.find(f"{W_NS}spacing")
            if spacing is not None:
                line = spacing.get(f"{W_NS}line")
                line_rule = spacing.get(f"{W_NS}lineRule")
                after = spacing.get(f"{W_NS}after")
                before = spacing.get(f"{W_NS}before")
                if line and line_rule == "auto":
                    try:
                        lh_multiplier = int(line) / 240
                        win_ratio = _FONT_WIN_RATIO.get(
                            default_body_font.lower(), _FONT_WIN_RATIO_DEFAULT
                        )
                        result["default_line_height"] = round(lh_multiplier * win_ratio, 4)
                        result["default_win_ratio"] = round(win_ratio, 4)
                    except ValueError:
                        pass
                if after is not None:
                    try:
                        result["default_para_after_twips"] = int(after)
                    except ValueError:
                        pass
                if before is not None:
                    try:
                        result["default_para_before_twips"] = int(before)
                    except ValueError:
                        pass

        # Fall back to Normal style if docDefaults doesn't have spacing
        if "default_line_height" not in result:
            for style in styles_root.findall(f"{W_NS}style"):
                name_elem = style.find(f"{W_NS}name")
                if name_elem is not None and name_elem.get(f"{W_NS}val") == "Normal":
                    spacing = style.find(f"{W_NS}pPr/{W_NS}spacing")
                    if spacing is not None:
                        line = spacing.get(f"{W_NS}line")
                        line_rule = spacing.get(f"{W_NS}lineRule")
                        if line and line_rule in ("auto", None):
                            try:
                                lh_multiplier = int(line) / 240
                                win_ratio = _FONT_WIN_RATIO.get(
                                    default_body_font.lower(), _FONT_WIN_RATIO_DEFAULT
                                )
                                result["default_line_height"] = round(lh_multiplier * win_ratio, 4)
                                result["default_win_ratio"] = round(win_ratio, 4)
                            except ValueError:
                                pass
                    break

    return result


# ---------------------------------------------------------------------------
# Archive utilities
# ---------------------------------------------------------------------------

def _read_xml(archive: zipfile.ZipFile, name: str) -> ET.Element:
    with archive.open(name) as handle:
        return ET.fromstring(handle.read())


def _parse_styles(styles_root: ET.Element) -> dict[str, str]:
    styles: dict[str, str] = {}
    for style in styles_root.findall(f"{W_NS}style"):
        style_id = style.get(f"{W_NS}styleId")
        style_type = style.get(f"{W_NS}type")
        name_elem = style.find(f"{W_NS}name")
        if style_id and style_type == "paragraph" and name_elem is not None:
            styles[style_id] = name_elem.get(f"{W_NS}val", "")
    return styles


def _parse_numbering(numbering_root: ET.Element) -> dict:
    """Parse word/numbering.xml → {numId: {ilvl: {numFmt, lvlText}}}."""
    abstract: dict = {}
    for an in numbering_root.findall(f"{W_NS}abstractNum"):
        aid = an.get(f"{W_NS}abstractNumId")
        if aid is None:
            continue
        levels: dict = {}
        for lvl in an.findall(f"{W_NS}lvl"):
            ilvl_val = lvl.get(f"{W_NS}ilvl")
            if ilvl_val is None:
                continue
            nf = lvl.find(f"{W_NS}numFmt")
            fmt = (nf.get(f"{W_NS}val") or nf.get("val") or "bullet") if nf is not None else "bullet"
            lt = lvl.find(f"{W_NS}lvlText")
            lvl_text = ""
            if lt is not None:
                lvl_text = lt.get(f"{W_NS}val") or lt.get("val") or ""
            try:
                levels[int(ilvl_val)] = {"numFmt": fmt, "lvlText": lvl_text}
            except ValueError:
                pass
        abstract[aid] = levels

    result: dict = {}
    for num in numbering_root.findall(f"{W_NS}num"):
        nid = num.get(f"{W_NS}numId")
        aid_elem = num.find(f"{W_NS}abstractNumId")
        if nid and aid_elem is not None:
            aid = aid_elem.get(f"{W_NS}val")
            result[nid] = abstract.get(aid or "", {})
    return result


def _annotate_list_paragraphs(ast: dict, numbering: dict) -> None:
    """Walk AST and add list_marker / list_ordered to paragraphs that have num_id."""
    def annotate_blocks(blocks: list) -> None:
        for block in blocks:
            if block.get("type") == "paragraph":
                style = block.get("style", {})
                num_id = style.get("num_id")
                if num_id:
                    ilvl = style.get("ilvl", 0)
                    level = numbering.get(str(num_id), {}).get(ilvl, {})
                    fmt = level.get("numFmt", "bullet")
                    lvl_text = level.get("lvlText", "")
                    if fmt in _ORDERED_NUM_FMTS:
                        style["list_ordered"] = True
                        style["list_marker"] = _NUM_FMT_TO_CSS.get(fmt, "decimal")
                    else:
                        style["list_ordered"] = False
                        # PUA or unrecognised bullet → square fallback
                        style["list_marker"] = _BULLET_CHAR_TO_MARKER.get(lvl_text, "square")
            elif block.get("type") == "table":
                for row in block.get("rows", []):
                    for cell in row.get("cells", []):
                        annotate_blocks(cell.get("children", []))

    for section in ast.get("sections", []):
        annotate_blocks(section.get("blocks", []))


def _parse_relationships(rels_root: ET.Element) -> dict[str, dict[str, str]]:
    relationships: dict[str, dict[str, str]] = {}
    for rel in rels_root.findall(f"{{{REL_NS}}}Relationship"):
        rel_id = rel.get("Id")
        if rel_id:
            relationships[rel_id] = {
                "type": rel.get("Type", ""),
                "target": rel.get("Target", ""),
            }
    return relationships


def _load_media(archive: zipfile.ZipFile) -> dict[str, bytes]:
    media: dict[str, bytes] = {}
    for name in archive.namelist():
        if name.startswith("word/media/"):
            with archive.open(name) as handle:
                media[name] = handle.read()
    return media


def _find_media_bytes(media: dict[str, bytes], target: str) -> tuple[Optional[str], Optional[bytes]]:
    if not target:
        return None, None

    candidates = [_resolve_media_path(target), target]
    if target.startswith("word/"):
        candidates.append(target[len("word/"):])
    candidates.append(target.lstrip("/"))

    for cand in candidates:
        if cand in media:
            return cand, media[cand]

    basename = os.path.basename(target)
    if basename:
        matches = [k for k in media if os.path.basename(k) == basename]
        if len(matches) == 1:
            return matches[0], media[matches[0]]

    return None, None


def _resolve_media_path(target: str) -> str:
    if not target:
        return target
    t = target.lstrip("/")
    while t.startswith("../"):
        t = t[3:]
    return t if t.startswith("word/") else f"word/{t}"


def _local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    import json
    import sys

    if len(sys.argv) != 2:
        print("Usage: python docx_ast.py input.docx", file=sys.stderr)
        sys.exit(2)

    try:
        ast = build_docx_ast(sys.argv[1])
    except DocxAstError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        sys.exit(1)

    # Print a compact summary (omit data URIs to keep output readable)
    def _summarise(obj: object, depth: int = 0) -> object:
        if isinstance(obj, dict):
            return {
                k: ("[base64 data omitted]" if k == "data_uri" else _summarise(v, depth + 1))
                for k, v in obj.items()
            }
        if isinstance(obj, list):
            return [_summarise(v, depth + 1) for v in obj]
        return obj

    print(json.dumps(_summarise(ast), indent=2, ensure_ascii=False))
