"""Apply positional patches to a DOCX file, return modified DOCX bytes."""
import io
import zipfile
import xml.etree.ElementTree as ET

W_NS   = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
W14_NS = '{http://schemas.microsoft.com/office/word/2010/wordml}'
XML_SPACE = '{http://www.w3.org/XML/1998/namespace}space'

# Register all common DOCX namespace prefixes so ET serialises them correctly
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
    """Apply positional patches to DOCX bytes and return the modified DOCX bytes.

    Each patch dict must contain:
      block (int)         — top-level block index (paragraph or table)
      para  (int)         — paragraph index within cell (0 for top-level paragraphs)
      run   (int)         — run index within paragraph
      row   (int|None)    — row index within table (None for top-level paragraphs)
      cell  (int|None)    — cell index within row  (None for top-level paragraphs)

    And exactly one of:
      text    (str)  — new text content for the run
      checked (bool) — new checked state for a checkbox sdt run

    Example patch for a table cell edit:
      {"block": 2, "row": 0, "cell": 1, "para": 0, "run": 0, "text": "edited text"}
    """
    for prefix, uri in DOCX_NS.items():
        ET.register_namespace(prefix, uri)

    with zipfile.ZipFile(io.BytesIO(docx_bytes)) as zin:
        files = {name: zin.read(name) for name in zin.namelist()}
        infos = {info.filename: info for info in zin.infolist()}

    root = ET.fromstring(files['word/document.xml'])
    body = root.find(f'.//{W_NS}body')
    if body is None:
        raise ValueError('word/document.xml has no <w:body>')

    # blocks = body children that are p or tbl (same order as docx_ast._build_blocks)
    block_elems = [c for c in body if _local(c.tag) in ('p', 'tbl')]

    for patch in patches:
        bi = patch.get('block')
        if bi is None or bi >= len(block_elems):
            continue
        block_el = block_elems[bi]
        btag = _local(block_el.tag)

        ri   = patch.get('row')
        ci   = patch.get('cell')
        pi   = patch.get('para', 0)
        runi = patch.get('run', 0)

        # Locate the paragraph element
        if btag == 'tbl' and ri is not None and ci is not None:
            rows  = [c for c in block_el if _local(c.tag) == 'tr']
            if ri >= len(rows):
                continue
            cells = [c for c in rows[ri] if _local(c.tag) == 'tc']
            if ci >= len(cells):
                continue
            paras = [c for c in cells[ci] if _local(c.tag) == 'p']
            if pi >= len(paras):
                continue
            para_el = paras[pi]
        elif btag == 'p':
            para_el = block_el
        else:
            continue

        # Locate the run element (same tags as docx_ast._build_paragraph)
        run_els = [c for c in para_el if _local(c.tag) in ('r', 'hyperlink', 'ins', 'del', 'sdt')]
        if runi >= len(run_els):
            continue
        run_el = run_els[runi]
        rtag = _local(run_el.tag)

        if 'text' in patch:
            t_els = run_el.findall(f'.//{W_NS}t')
            if t_els:
                t_els[0].text = patch['text']
                t_els[0].set(XML_SPACE, 'preserve')
                for t in t_els[1:]:
                    t.text = ''
            elif rtag == 'r':
                # No <w:t> exists yet (empty run) — create one
                t_new = ET.SubElement(run_el, f'{W_NS}t')
                t_new.text = patch['text']
                t_new.set(XML_SPACE, 'preserve')

        if 'checked' in patch and rtag == 'sdt':
            sdt_pr = run_el.find(f'{W_NS}sdtPr')
            if sdt_pr is not None:
                cb = sdt_pr.find(f'{W14_NS}checkbox')
                if cb is not None:
                    ch_el = cb.find(f'{W14_NS}checked')
                    if ch_el is not None:
                        ch_el.set(f'{W14_NS}val', '1' if patch['checked'] else '0')
            sdt_content = run_el.find(f'{W_NS}sdtContent')
            if sdt_content is not None:
                for t in sdt_content.findall(f'.//{W_NS}t'):
                    t.text = '☒' if patch['checked'] else '☐'

    # Re-serialise document.xml preserving the XML declaration
    xml_bytes = ET.tostring(root, encoding='unicode').encode('utf-8')
    xml_bytes = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + xml_bytes
    files['word/document.xml'] = xml_bytes

    # Re-zip
    out = io.BytesIO()
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as zout:
        for name, data in files.items():
            info = infos.get(name)
            zout.writestr(info or name, data)
    return out.getvalue()
