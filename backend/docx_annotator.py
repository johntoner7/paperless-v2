"""Phase 4.2 — AI annotation generator.

Calls the Anthropic API with a compact AST summary and returns a validated
annotations dict conforming to docx_annotation_schema.json.

Usage:
    from docx_annotator import generate_annotations
    annotations = generate_annotations(ast)        # requires ANTHROPIC_API_KEY env var
    html = render_ast_to_html(ast, annotations)
"""

from __future__ import annotations

import json
import re
from typing import Optional

import anthropic

# ---------------------------------------------------------------------------
# System prompt (static — benefits from prompt caching)
# ---------------------------------------------------------------------------

_SYSTEM_PROMPT = """\
You are a document layout analyst. Given a DOCX document's structural summary \
(an AST excerpt), you produce layout annotations — presentation-only hints that \
improve visual rendering without changing any content.

RULES:
1. Never modify text content.
2. Never generate HTML.
3. Only output valid JSON matching the OUTPUT FORMAT below.
4. Only annotate blocks where you have a clear, specific reason.
5. Be conservative — fewer, high-confidence annotations beat noisy guesses.

BLOCK KEY FORMAT:
Every block has a key like "s{section}:b{block}" where section and block are \
zero-based indices. Example: "s0:b3" = section 0 (body), block index 3.

AVAILABLE ANNOTATION FIELDS (all optional):
  heading_weight  int 1–6       Override rendered heading level
  spacing_before  string        Extra space above: "none"|"small"|"medium"|"large"
  spacing_after   string        Extra space below: same values
  emphasis        string        Visual prominence: "low"|"medium"|"high"
  image_display   string        "inline"|"block"|"float-right"|"full-width"
  suppress        boolean       De-emphasise decorative or empty spacer blocks
  group           string        Logical group name (e.g. "address-block")
  group_role      string        "header"|"label"|"value"|"separator"|"footer"

HEURISTICS TO APPLY:
- Empty paragraphs that are purely structural spacers → suppress: true
- Document title or company name (large text, top of doc) → heading_weight: 1 or 2
- Section labels before data tables → spacing_before: "medium", group_role: "header"
- Totals rows or summary sections → emphasis: "medium"
- Logo images (small, top of doc, in header paragraph) → image_display: "block"
- Wide content images → image_display: "full-width"
- Address blocks, billing info → group: "address-block"
- Invoice totals section → group: "totals-section"

OUTPUT FORMAT — return exactly this JSON shape, nothing else:
{
  "document_type": "invoice",
  "annotations": {
    "s0:b2": { "heading_weight": 1, "spacing_after": "small" },
    "s0:b5": { "suppress": true }
  }
}"""

# ---------------------------------------------------------------------------
# Validated annotation field sets
# ---------------------------------------------------------------------------

_VALID_SPACINGS = {"none", "small", "medium", "large"}
_VALID_EMPHASIS = {"low", "medium", "high"}
_VALID_IMG_DISPLAY = {"inline", "block", "float-right", "full-width"}
_VALID_GROUP_ROLES = {"header", "label", "value", "separator", "footer"}
_VALID_DOC_TYPES = {"invoice", "letter", "report", "contract", "form", "other"}


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def generate_annotations(
    ast: dict,
    model: str = "claude-sonnet-4-6",
) -> dict:
    """Call the Anthropic API and return a validated annotations dict.

    Returns {"version": "1.0", "annotations": {}} on any failure so callers
    can always proceed with rendering.
    """
    client = anthropic.Anthropic()
    prompt_ast = _build_prompt_ast(ast)
    user_content = (
        "Analyse this document AST and return layout annotations:\n\n"
        + json.dumps(prompt_ast, indent=2, ensure_ascii=False)
    )

    try:
        response = client.messages.create(
            model=model,
            max_tokens=2048,
            system=[
                {
                    "type": "text",
                    "text": _SYSTEM_PROMPT,
                    "cache_control": {"type": "ephemeral"},
                }
            ],
            messages=[{"role": "user", "content": user_content}],
        )
        raw = response.content[0].text
    except Exception as exc:
        print(f"[annotator] API call failed: {exc}")
        return _empty()

    return _parse_response(raw)


# ---------------------------------------------------------------------------
# Compact AST builder — strips data URIs and deep nesting for the prompt
# ---------------------------------------------------------------------------

def _build_prompt_ast(ast: dict) -> dict:
    sections = []
    for si, section in enumerate(ast.get("sections", [])):
        blocks = []
        for bi, block in enumerate(section.get("blocks", [])):
            blocks.append(_summarise_block(block, f"s{si}:b{bi}"))
        sections.append({"kind": section.get("kind"), "blocks": blocks})
    return {"sections": sections}


def _summarise_block(block: dict, key: str) -> dict:
    t = block.get("type")
    if t == "paragraph":
        return _summarise_paragraph(block, key)
    if t == "table":
        return _summarise_table(block, key)
    return {"type": t, "key": key}


def _summarise_paragraph(block: dict, key: str) -> dict:
    style = block.get("style", {})
    text = _extract_text(block)
    has_image = _has_image(block)

    result: dict = {"type": "paragraph", "key": key, "text": text[:300]}
    if style.get("style_id"):
        result["style_id"] = style["style_id"]
    if style.get("style_name"):
        result["style_name"] = style["style_name"]
    if style.get("alignment"):
        result["alignment"] = style["alignment"]
    if style.get("num_id"):
        result["is_list"] = True
        result["list_level"] = style.get("ilvl", 0)
    if has_image:
        result["has_image"] = True

    # Include dominant run style info (font size of first non-empty run)
    first_run_style = _first_run_style(block)
    if first_run_style:
        result["run_style"] = first_run_style

    return result


def _summarise_table(block: dict, key: str) -> dict:
    rows = block.get("rows", [])
    cols = max((len(r.get("cells", [])) for r in rows), default=0)
    # Sample first-row text for context
    header_text: list[str] = []
    if rows:
        for cell in rows[0].get("cells", []):
            header_text.append(_extract_text(cell)[:40])
    return {
        "type": "table",
        "key": key,
        "rows": len(rows),
        "cols": cols,
        "header_sample": header_text,
    }


def _extract_text(obj: object) -> str:
    parts: list[str] = []

    def walk(o: object) -> None:
        if isinstance(o, dict):
            if o.get("type") == "text":
                parts.append(o.get("text", ""))
            else:
                for v in o.values():
                    walk(v)
        elif isinstance(o, list):
            for v in o:
                walk(v)

    walk(obj)
    return "".join(parts)


def _has_image(obj: object) -> bool:
    if isinstance(obj, dict):
        if obj.get("type") == "image":
            return True
        return any(_has_image(v) for v in obj.values())
    if isinstance(obj, list):
        return any(_has_image(v) for v in obj)
    return False


def _first_run_style(block: dict) -> Optional[dict]:
    for child in block.get("children", []):
        if child.get("type") == "run":
            style = child.get("style", {})
            result = {}
            if style.get("font_size_pt"):
                result["font_size_pt"] = style["font_size_pt"]
            if style.get("bold"):
                result["bold"] = True
            if style.get("color"):
                result["color"] = style["color"]
            return result or None
    return None


# ---------------------------------------------------------------------------
# Response parsing and validation
# ---------------------------------------------------------------------------

def _parse_response(raw: str) -> dict:
    raw = raw.strip()

    # Strip markdown code fences if present
    if raw.startswith("```"):
        lines = raw.splitlines()
        # Drop first line (```json or ```) and last line (```)
        inner_lines = lines[1:]
        if inner_lines and inner_lines[-1].strip() == "```":
            inner_lines = inner_lines[:-1]
        raw = "\n".join(inner_lines)

    try:
        data = json.loads(raw)
    except json.JSONDecodeError:
        # Try to extract a JSON object from freeform text
        m = re.search(r"\{.*\}", raw, re.DOTALL)
        if not m:
            return _empty()
        try:
            data = json.loads(m.group())
        except json.JSONDecodeError:
            return _empty()

    if not isinstance(data, dict):
        return _empty()

    annotations: dict = {}
    raw_anns = data.get("annotations", {})
    if isinstance(raw_anns, dict):
        for key, ann in raw_anns.items():
            if not isinstance(key, str) or not isinstance(ann, dict):
                continue
            clean = _validate_annotation(ann)
            if clean:
                annotations[key] = clean

    result: dict = {"version": "1.0", "annotations": annotations}
    doc_type = data.get("document_type")
    if isinstance(doc_type, str) and doc_type in _VALID_DOC_TYPES:
        result["document_type"] = doc_type

    return result


def _validate_annotation(ann: dict) -> dict:
    clean: dict = {}

    hw = ann.get("heading_weight")
    if isinstance(hw, int) and 1 <= hw <= 6:
        clean["heading_weight"] = hw

    for field in ("spacing_before", "spacing_after"):
        if ann.get(field) in _VALID_SPACINGS:
            clean[field] = ann[field]

    if ann.get("emphasis") in _VALID_EMPHASIS:
        clean["emphasis"] = ann["emphasis"]

    if ann.get("image_display") in _VALID_IMG_DISPLAY:
        clean["image_display"] = ann["image_display"]

    if ann.get("suppress") is True:
        clean["suppress"] = True

    if isinstance(ann.get("group"), str):
        clean["group"] = ann["group"]

    if ann.get("group_role") in _VALID_GROUP_ROLES:
        clean["group_role"] = ann["group_role"]

    return clean


def _empty() -> dict:
    return {"version": "1.0", "annotations": {}}


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    import sys
    try:
        from .docx_ast import build_docx_ast
    except ImportError:
        from docx_ast import build_docx_ast

    if len(sys.argv) < 2:
        print("Usage: python docx_annotator.py input.docx [out.json]", file=sys.stderr)
        sys.exit(2)

    try:
        ast = build_docx_ast(sys.argv[1])
    except Exception as exc:
        print(f"AST build failed: {exc}", file=sys.stderr)
        sys.exit(1)

    annotations = generate_annotations(ast)

    out_path = sys.argv[2] if len(sys.argv) > 2 else None
    if out_path:
        with open(out_path, "w", encoding="utf-8") as fh:
            json.dump(annotations, fh, indent=2, ensure_ascii=False)
        print(f"Wrote: {out_path}")

    print(json.dumps(annotations, indent=2, ensure_ascii=False))
