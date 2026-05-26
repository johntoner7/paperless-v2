"""Apply identity-based patches to a DOCX file, return modified DOCX bytes."""
import io
import re
import zipfile
import xml.etree.ElementTree as ET

try:
    from .docx_node_ids import NODE_ID_ATTR, inject_docx_node_ids
except ImportError:
    from docx_node_ids import NODE_ID_ATTR, inject_docx_node_ids

W_NS   = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
W14_NS = '{http://schemas.microsoft.com/office/word/2010/wordml}'
XML_SPACE = '{http://www.w3.org/XML/1998/namespace}space'
REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'

# Seed the ET namespace registry with the most common DOCX prefixes.
# _collect_namespaces() will register the rest from the actual source document.
DOCX_NS = {
    'wpc': 'http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas',
    'mc':  'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'r':   'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'w':   'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
    'w15': 'http://schemas.microsoft.com/office/word/2012/wordml',
    'a':   'http://schemas.openxmlformats.org/drawingml/2006/main',
    'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
    'wp':  'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'o':   'urn:schemas-microsoft-com:office:office',
    'v':   'urn:schemas-microsoft-com:vml',
    'wne': 'http://schemas.microsoft.com/office/word/2006/wordml',
}


def _local(tag: str) -> str:
    return tag.split('}', 1)[-1] if '}' in tag else tag


def apply_patches(docx_bytes: bytes, patches: list[dict]) -> bytes:
    """Apply identity-based patches to DOCX bytes and return the modified DOCX bytes.

    Preferred patch shape:
      {"node_id": "r_12", "operation": "replace_text", "text": "edited text"}
      {"node_id": "sdt_7", "operation": "toggle_checkbox", "checked": true}

    Legacy positional patches are still accepted as a fallback.
    """
    for prefix, uri in DOCX_NS.items():
        ET.register_namespace(prefix, uri)

    with zipfile.ZipFile(io.BytesIO(docx_bytes)) as zin:
        files = {name: zin.read(name) for name in zin.namelist()}
        infos = {info.filename: info for info in zin.infolist()}

    # Collect the full namespace map from each XML part and register all prefixes
    # with ET so it uses the original prefix aliases when serialising.
    # We also keep the per-part map so we can re-inject xmlns: declarations that
    # ET drops for namespaces only referenced as string tokens (e.g. in
    # mc:Ignorable="w15 w16se …") but not used in any element or attribute name.
    source_ns_maps: dict[str, dict[str, str]] = {}
    for name, data in files.items():
        if name.endswith('.xml') or name.endswith('.rels'):
            ns_map = _collect_namespaces(data)
            source_ns_maps[name] = ns_map

    roots = _load_docx_xml_roots(files)
    id_counters = {"p": 0, "r": 0, "sdt": 0}
    for root in roots.values():
        inject_docx_node_ids(root, id_counters)

    node_index = _index_docx_nodes(roots)
    document_root = roots.get('word/document.xml')
    if document_root is None:
        raise ValueError('word/document.xml is missing')

    for patch in patches:
        node_id = patch.get('node_id')
        if node_id:
            target = node_index.get(str(node_id))
            if target is None:
                continue
            _part_name, node = target
            _apply_node_patch(node, patch)
            continue

        _apply_legacy_patch(document_root, patch)

    # Strip injected node IDs — they are invalid OOXML and corrupt the file in Word.
    for root in roots.values():
        _strip_node_ids(root)

    for part_name, root in roots.items():
        files[part_name] = _serialize_xml(root, source_ns_maps.get(part_name))

    # Re-zip
    out = io.BytesIO()
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as zout:
        for name, data in files.items():
            info = infos.get(name)
            zout.writestr(info or name, data)
    return out.getvalue()


def _collect_namespaces(xml_bytes: bytes) -> dict[str, str]:
    """Collect all namespace prefix→URI mappings from XML bytes and register them with ET."""
    ns_map: dict[str, str] = {}
    try:
        for _, (prefix, uri) in ET.iterparse(io.BytesIO(xml_bytes), events=('start-ns',)):
            if prefix not in ns_map:
                ns_map[prefix] = uri
                if prefix:
                    ET.register_namespace(prefix, uri)
    except ET.ParseError:
        pass
    return ns_map


def _strip_node_ids(root: ET.Element) -> None:
    """Remove injected docx-node-id attributes before writing output."""
    for elem in root.iter():
        elem.attrib.pop(NODE_ID_ATTR, None)


def _serialize_xml(root: ET.Element, source_ns_map: dict[str, str] | None = None) -> bytes:
    xml_str = ET.tostring(root, encoding='unicode')
    if source_ns_map:
        xml_str = _inject_missing_xmlns(xml_str, source_ns_map)
    return b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + xml_str.encode('utf-8')


def _inject_missing_xmlns(xml_str: str, source_ns_map: dict[str, str]) -> str:
    """Re-add xmlns: declarations that ET drops for namespaces only referenced
    as string tokens (e.g. inside mc:Ignorable) but not used in element/attribute names."""
    # Scan the first 8 KB — enough to cover any root element's attribute list
    existing = set(re.findall(r'xmlns:(\w+)=', xml_str[:8000]))
    missing = [
        f'xmlns:{prefix}="{uri}"'
        for prefix, uri in source_ns_map.items()
        if prefix and prefix not in existing
    ]
    if not missing:
        return xml_str
    # Insert right after the root element name: <w:document<INJECT HERE> ...
    match = re.match(r'^<[^\s>]+', xml_str)
    if not match:
        return xml_str
    insert_at = match.end()
    return xml_str[:insert_at] + ' ' + ' '.join(missing) + xml_str[insert_at:]


def _load_docx_xml_roots(files: dict[str, bytes]) -> dict[str, ET.Element]:
    roots: dict[str, ET.Element] = {
        'word/document.xml': ET.fromstring(files['word/document.xml']),
    }
    rels = _parse_relationships(files.get('word/_rels/document.xml.rels'))
    for rel in rels.values():
        rel_type = rel.get('type', '')
        target = rel.get('target', '')
        if not rel_type or not target:
            continue
        is_header = rel_type.endswith('/header') or rel_type.endswith('/headers')
        is_footer = rel_type.endswith('/footer') or rel_type.endswith('/footers')
        if not is_header and not is_footer:
            continue
        part_path = f'word/{target}' if not target.startswith('word/') else target
        part_bytes = files.get(part_path)
        if part_bytes is None:
            continue
        roots[part_path] = ET.fromstring(part_bytes)
    return roots


def _parse_relationships(rels_bytes: bytes | None) -> dict[str, dict[str, str]]:
    if not rels_bytes:
        return {}
    root = ET.fromstring(rels_bytes)
    relationships: dict[str, dict[str, str]] = {}
    for rel in root.findall(f'{{{REL_NS}}}Relationship'):
        rel_id = rel.get('Id')
        if not rel_id:
            continue
        relationships[rel_id] = {
            'type': rel.get('Type', ''),
            'target': rel.get('Target', ''),
        }
    return relationships


def _index_docx_nodes(roots: dict[str, ET.Element]) -> dict[str, tuple[str, ET.Element]]:
    indexed: dict[str, tuple[str, ET.Element]] = {}
    for part_name, root in roots.items():
        for elem in root.iter():
            node_id = elem.get(NODE_ID_ATTR)
            if node_id:
                indexed[node_id] = (part_name, elem)
    return indexed


def _apply_node_patch(node: ET.Element, patch: dict) -> None:
    operation = patch.get('operation') or _infer_operation(patch)
    if operation == 'replace_text':
        _replace_node_text(node, patch.get('text', ''))
    elif operation == 'toggle_checkbox':
        _toggle_checkbox(node, bool(patch.get('checked')))


def _apply_legacy_patch(document_root: ET.Element, patch: dict) -> None:
    body = document_root.find(f'.//{W_NS}body')
    if body is None:
        return

    block_elems = [c for c in body if _local(c.tag) in ('p', 'tbl')]
    bi = patch.get('block')
    if bi is None or bi >= len(block_elems):
        return

    block_el = block_elems[bi]
    btag = _local(block_el.tag)
    ri = patch.get('row')
    ci = patch.get('cell')
    pi = patch.get('para', 0)
    runi = patch.get('run', 0)

    if btag == 'tbl' and ri is not None and ci is not None:
        rows = [c for c in block_el if _local(c.tag) == 'tr']
        if ri >= len(rows):
            return
        cells = [c for c in rows[ri] if _local(c.tag) == 'tc']
        if ci >= len(cells):
            return
        paras = [c for c in cells[ci] if _local(c.tag) == 'p']
        if pi >= len(paras):
            return
        para_el = paras[pi]
    elif btag == 'p':
        para_el = block_el
    else:
        return

    run_els = [c for c in para_el if _local(c.tag) in ('r', 'hyperlink', 'ins', 'del', 'sdt')]
    if runi >= len(run_els):
        return
    run_el = run_els[runi]
    if 'text' in patch:
        _replace_node_text(run_el, patch['text'])
    if 'checked' in patch:
        _toggle_checkbox(run_el, bool(patch['checked']))


def _replace_node_text(node: ET.Element, text: str) -> None:
    text_nodes = list(_iter_text_nodes(node))
    if text_nodes:
        # Preserve the surrounding run structure, but rewrite the visible text
        # across all text leaves rather than assuming there is only one.
        text_nodes[0].text = text
        text_nodes[0].set(XML_SPACE, 'preserve')
        for text_node in text_nodes[1:]:
            text_node.text = ''
        return

    if _local(node.tag) == 'r':
        t_new = ET.SubElement(node, f'{W_NS}t')
        t_new.text = text
        t_new.set(XML_SPACE, 'preserve')


def _iter_text_nodes(node: ET.Element):
    """Yield all visible text nodes under a run-like container in document order."""
    for child in node.iter():
        if _local(child.tag) == 't':
            yield child


def _toggle_checkbox(node: ET.Element, checked: bool) -> None:
    target = node
    if _local(node.tag) == 'r':
        sdt = node.find(f'{W_NS}sdt')
        if sdt is not None:
            target = sdt

    if _local(target.tag) != 'sdt':
        return

    sdt_pr = target.find(f'{W_NS}sdtPr')
    if sdt_pr is not None:
        cb = sdt_pr.find(f'{W14_NS}checkbox')
        if cb is not None:
            ch_el = cb.find(f'{W14_NS}checked')
            if ch_el is not None:
                ch_el.set(f'{W14_NS}val', '1' if checked else '0')
    sdt_content = target.find(f'{W_NS}sdtContent')
    if sdt_content is not None:
        for t in sdt_content.findall(f'.//{W_NS}t'):
            t.text = '☒' if checked else '☐'


def _infer_operation(patch: dict) -> str | None:
    if 'checked' in patch:
        return 'toggle_checkbox'
    if 'text' in patch:
        return 'replace_text'
    return None
