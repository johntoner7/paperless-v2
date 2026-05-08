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

W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
R_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
A_NS = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
WP_NS = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
VML_NS = "{urn:schemas-microsoft-com:vml}"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"


class DocxAstError(RuntimeError):
    pass


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
            styles = _parse_styles(_read_xml(archive, "word/styles.xml"))
            relationships = _parse_relationships(
                _read_xml(archive, "word/_rels/document.xml.rels")
            )
            media = _load_media(archive)

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
        sections.append({"kind": "header", "blocks": _build_blocks(hdr, styles, hdr_rels, media)})
    sections.append({"kind": "body", "blocks": _build_blocks(body_elem, styles, relationships, media)})
    for ftr, ftr_rels in footer_parts:
        sections.append({"kind": "footer", "blocks": _build_blocks(ftr, styles, ftr_rels, media)})

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

    page_setup = _extract_page_setup(body_elem)

    return {
        "version": "1.0",
        "sections": sections,
        "relationships": relationships,
        "images": images,
        "page_setup": page_setup,
    }


# ---------------------------------------------------------------------------
# Block builders
# ---------------------------------------------------------------------------

def _build_blocks(root: ET.Element, styles: dict, relationships: dict, media: dict) -> list[dict]:
    blocks: list[dict] = []
    for child in root:
        tag = _local_name(child.tag)
        if tag == "p":
            blocks.append(_build_paragraph(child, styles, relationships, media))
        elif tag == "tbl":
            blocks.append(_build_table(child, styles, relationships, media))
    return blocks


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
    return {"type": "paragraph", "style": pstyle, "children": children}


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
            children.append({"type": "break", "break_type": "line"})
        elif tag == "drawing":
            img = _build_image(child, relationships, media)
            if img is not None:
                children.append(img)

    return {"type": "run", "style": run_style, "children": children}


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
) -> dict:
    column_widths = _parse_column_widths(table)
    tbl_borders = _extract_tbl_borders(table)
    rows = _build_table_rows(table, styles, relationships, media)
    return {"type": "table", "column_widths_twips": column_widths, "tbl_borders": tbl_borders, "rows": rows}


def _extract_tbl_borders(table: ET.Element) -> dict:
    tbl_pr = table.find(f"{W_NS}tblPr")
    if tbl_pr is None:
        return {}
    borders_elem = tbl_pr.find(f"{W_NS}tblBorders")
    if borders_elem is None:
        return {}
    result: dict = {}
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
) -> list[dict]:
    # First pass: collect raw cell data with vmerge tracking
    raw_rows: list[list[dict]] = []
    for row_index, row in enumerate(table.findall(f"{W_NS}tr")):
        entries: list[dict] = []
        current_col = 0
        for cell in row.findall(f"{W_NS}tc"):
            colspan = _parse_cell_colspan(cell)
            vmerge = _parse_cell_vmerge(cell)
            width_twips = _parse_cell_width_twips(cell)
            cell_style = _extract_cell_style(cell)
            children = _build_blocks(cell, styles, relationships, media)
            entries.append({
                "_col": current_col,
                "colspan": colspan,
                "rowspan": 1,
                "_vmerge": vmerge,
                "width_twips": width_twips,
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
        rows.append({"cells": cells})
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

def _extract_page_setup(body: ET.Element) -> dict:
    """Extract page dimensions and margins from w:sectPr (twips → inches)."""
    sect = body.find(f"{W_NS}sectPr")
    if sect is None:
        return {}

    result: dict = {}

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
