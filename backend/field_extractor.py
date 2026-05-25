"""AST-based form field extraction (phase 1 rules + phase 2 AI suggestions)."""

from __future__ import annotations

import json
import re
from typing import Iterable, Optional

import anthropic


_UNDERSCORE_RE = re.compile(r"_{3,}")
_DATE_HINT_RE = re.compile(r"\bdate\b", re.IGNORECASE)
_DATE_VALUE_RE = re.compile(r"\b\d{1,2}[/-]\d{1,2}[/-]\d{2,4}\b")
_NUMBER_HINT_RE = re.compile(r"\b(amount|qty|quantity|total|number|no\.)\b", re.IGNORECASE)

_CHECKBOX_CHARS = {"\u2610", "\u2611", "\u2612"}  # ☐ ☑ ☒


def extract_fields_from_ast(ast: dict, document_id: Optional[str] = None) -> dict:
    """Extract form fields from a Docx AST using deterministic rules."""
    fields: list[dict] = []
    counter = 1

    body_blocks = _get_body_blocks(ast)
    for block_index, block in enumerate(body_blocks):
        if block.get("type") == "paragraph":
            for field in _extract_from_paragraph(block, block_index, 0):
                field["id"] = _field_id(counter)
                counter += 1
                fields.append(field)
        elif block.get("type") == "table":
            for field in _extract_from_table(block, block_index):
                field["id"] = _field_id(counter)
                counter += 1
                fields.append(field)

    return {
        "documentId": document_id,
        "fields": fields,
    }


def extract_fields_from_ast_with_ai(
    ast: dict,
    document_id: Optional[str] = None,
    model: str = "claude-sonnet-4-6",
) -> dict:
    """Extract fields using rules, then add AI-suggested fields."""
    base = extract_fields_from_ast(ast, document_id=document_id)
    ai_fields = suggest_fields_with_ai(ast, model=model)
    merged = _merge_fields(base["fields"], ai_fields)
    return {"documentId": document_id, "fields": merged}


def suggest_fields_with_ai(ast: dict, model: str = "claude-sonnet-4-6") -> list[dict]:
    """Call LLM to suggest additional fields not found by rules."""
    prompt_ast = _build_ai_prompt_ast(ast)
    if not prompt_ast.get("blocks"):
        return []

    system_prompt = (
        "You extract form fields from a DOCX AST summary. "
        "Return JSON only, matching the OUTPUT FORMAT. "
        "Do not invent fields that are not supported by the text."
    )
    user_prompt = (
        "Identify form fields users should fill in. Use block/para references.\n\n"
        + json.dumps(prompt_ast, indent=2, ensure_ascii=False)
        + "\n\nOUTPUT FORMAT:\n"
        + "{\"fields\":[{\"label\":\"...\",\"type\":\"text|date|checkbox|choice|number|multiline\","
        + "\"block\":0,\"para\":0,\"confidence\":0.0,\"labelText\":\"...\",\"valueText\":\"...\"}]}"
    )

    try:
        client = anthropic.Anthropic()
        response = client.messages.create(
            model=model,
            max_tokens=1024,
            system=[{"type": "text", "text": system_prompt, "cache_control": {"type": "ephemeral"}}],
            messages=[{"role": "user", "content": user_prompt}],
        )
        raw = response.content[0].text
    except Exception:
        return []

    parsed = _parse_ai_response(raw)
    return _normalize_ai_fields(parsed)


def _get_body_blocks(ast: dict) -> list:
    for section in ast.get("sections", []):
        if section.get("kind") == "body":
            return section.get("blocks", [])
    return []


def _extract_from_paragraph(paragraph: dict, block_index: int, para_index: int) -> Iterable[dict]:
    tokens = _flatten_paragraph(paragraph)
    text = _tokens_to_text(tokens)
    text_stripped = text.strip()
    if not text_stripped:
        return []

    node_ids = _collect_node_ids(paragraph)
    fields: list[dict] = []

    # Rule: checkbox glyphs or checkbox runs in a paragraph
    checkbox_count = _count_checkboxes(tokens)
    if checkbox_count > 0:
        label = _strip_checkbox_text(text).strip()
        if label:
            field_type = "choice" if checkbox_count > 1 else "checkbox"
            fields.append(
                _build_field(
                    label=label,
                    field_type=field_type,
                    value=None,
                    confidence=0.9,
                    block_index=block_index,
                    para_index=para_index,
                    node_ids=node_ids,
                    raw_label=text_stripped,
                    raw_value=None,
                )
            )

    # Rule: label/value in same paragraph using colon
    if ":" in text_stripped:
        label_part, value_part = text_stripped.split(":", 1)
        label = label_part.strip()
        value_text = value_part.strip()
        if label:
            has_placeholder = _UNDERSCORE_RE.search(value_part) is not None or _has_tab(tokens)
            field_type = _guess_type(label, value_text)
            confidence = 0.85 if has_placeholder else 0.6
            value = _normalize_value(value_text)
            fields.append(
                _build_field(
                    label=label,
                    field_type=field_type,
                    value=value,
                    confidence=confidence,
                    block_index=block_index,
                    para_index=para_index,
                    node_ids=node_ids,
                    raw_label=label_part + ":",
                    raw_value=value_text,
                )
            )
            return fields

    # Rule: underscores after a label without colon
    m = _UNDERSCORE_RE.search(text_stripped)
    if m:
        label = text_stripped[: m.start()].strip()
        if label:
            field_type = _guess_type(label, "")
            fields.append(
                _build_field(
                    label=label,
                    field_type=field_type,
                    value=None,
                    confidence=0.8,
                    block_index=block_index,
                    para_index=para_index,
                    node_ids=node_ids,
                    raw_label=label,
                    raw_value=text_stripped[m.start() :].strip(),
                )
            )

    return fields


def _extract_from_table(table: dict, block_index: int) -> Iterable[dict]:
    fields: list[dict] = []
    rows = table.get("rows", [])
    for row in rows:
        cells = row.get("cells", [])
        if len(cells) < 2:
            continue
        label_text, label_nodes = _cell_text_and_nodes(cells[0])
        value_text, value_nodes = _cell_text_and_nodes(cells[1])
        if not label_text:
            continue

        has_placeholder = not value_text or _UNDERSCORE_RE.search(value_text)
        checkbox_count = _count_checkboxes_in_cell(cells[1])
        if checkbox_count > 0:
            field_type = "choice" if checkbox_count > 1 else "checkbox"
            confidence = 0.85
            value = None
        else:
            field_type = _guess_type(label_text, value_text)
            confidence = 0.8 if has_placeholder else 0.6
            value = _normalize_value(value_text)

        node_ids = list(dict.fromkeys(label_nodes + value_nodes))
        fields.append(
            _build_field(
                label=label_text,
                field_type=field_type,
                value=value,
                confidence=confidence,
                block_index=block_index,
                para_index=0,
                node_ids=node_ids,
                raw_label=label_text,
                raw_value=value_text,
            )
        )

    return fields


def _build_ai_prompt_ast(ast: dict) -> dict:
    blocks: list[dict] = []
    body_blocks = _get_body_blocks(ast)
    for bi, block in enumerate(body_blocks):
        btype = block.get("type")
        if btype == "paragraph":
            tokens = _flatten_paragraph(block)
            text = _tokens_to_text(tokens).strip()
            if not text:
                continue
            blocks.append(
                {
                    "block": bi,
                    "type": "paragraph",
                    "text": text[:240],
                    "has_checkbox": _count_checkboxes(tokens) > 0,
                }
            )
        elif btype == "table":
            blocks.append({"block": bi, "type": "table", "rows": _summarize_table(block)})
    return {"blocks": blocks}


def _summarize_table(table: dict) -> list[list[str]]:
    rows_out: list[list[str]] = []
    for row in table.get("rows", [])[:6]:
        row_cells: list[str] = []
        for cell in row.get("cells", [])[:4]:
            row_cells.append(_cell_text(cell)[:80])
        rows_out.append(row_cells)
    return rows_out


def _cell_text(cell: dict) -> str:
    texts: list[str] = []
    for block in cell.get("children", []):
        if block.get("type") != "paragraph":
            continue
        tokens = _flatten_paragraph(block)
        text = _tokens_to_text(tokens).strip()
        if text:
            texts.append(text)
    return " ".join(texts).strip()


def _cell_text_and_nodes(cell: dict) -> tuple[str, list[str]]:
    texts: list[str] = []
    node_ids: list[str] = []
    for block in cell.get("children", []):
        if block.get("type") != "paragraph":
            continue
        tokens = _flatten_paragraph(block)
        text = _tokens_to_text(tokens).strip()
        if text:
            texts.append(text)
        node_ids.extend(_collect_node_ids(block))
    return " ".join(texts).strip(), node_ids


def _count_checkboxes_in_cell(cell: dict) -> int:
    count = 0
    for block in cell.get("children", []):
        if block.get("type") != "paragraph":
            continue
        tokens = _flatten_paragraph(block)
        count += _count_checkboxes(tokens)
    return count


def _flatten_paragraph(paragraph: dict) -> list[dict]:
    tokens: list[dict] = []
    for inline in paragraph.get("children", []):
        kind = inline.get("type")
        if kind == "run":
            tokens.extend(_flatten_run(inline))
        elif kind == "hyperlink":
            for run in inline.get("children", []):
                if run.get("type") == "run":
                    tokens.extend(_flatten_run(run))
    return tokens


def _flatten_run(run: dict) -> list[dict]:
    tokens: list[dict] = []
    for item in run.get("children", []):
        itype = item.get("type")
        if itype == "text":
            tokens.append({"kind": "text", "text": item.get("text", "")})
        elif itype == "break":
            if item.get("break_type") == "tab":
                tokens.append({"kind": "tab"})
            else:
                tokens.append({"kind": "text", "text": "\n"})
        elif itype == "checkbox":
            tokens.append({"kind": "checkbox", "checked": bool(item.get("checked"))})
    return tokens


def _tokens_to_text(tokens: list[dict]) -> str:
    parts: list[str] = []
    for token in tokens:
        kind = token.get("kind")
        if kind == "text":
            parts.append(token.get("text", ""))
        elif kind == "tab":
            parts.append("\t")
        elif kind == "checkbox":
            parts.append(" ")
    return "".join(parts)


def _strip_checkbox_text(text: str) -> str:
    out = text
    for ch in _CHECKBOX_CHARS:
        out = out.replace(ch, " ")
    return out


def _count_checkboxes(tokens: list[dict]) -> int:
    return sum(1 for token in tokens if token.get("kind") == "checkbox")


def _has_tab(tokens: list[dict]) -> bool:
    return any(token.get("kind") == "tab" for token in tokens)


def _collect_node_ids(paragraph: dict) -> list[str]:
    node_ids: list[str] = []
    if paragraph.get("node_id"):
        node_ids.append(str(paragraph.get("node_id")))
    for inline in paragraph.get("children", []):
        if inline.get("type") == "run" and inline.get("node_id"):
            node_ids.append(str(inline.get("node_id")))
        elif inline.get("type") == "hyperlink":
            for run in inline.get("children", []):
                if run.get("type") == "run" and run.get("node_id"):
                    node_ids.append(str(run.get("node_id")))
    return list(dict.fromkeys(node_ids))


def _guess_type(label: str, value_text: str) -> str:
    combined = f"{label} {value_text}".strip()
    if _DATE_HINT_RE.search(combined) or _DATE_VALUE_RE.search(value_text):
        return "date"
    if _NUMBER_HINT_RE.search(combined):
        return "number"
    return "text"


def _normalize_value(value_text: str) -> Optional[str]:
    if not value_text:
        return None
    if _UNDERSCORE_RE.search(value_text):
        return None
    return value_text


def _build_field(
    label: str,
    field_type: str,
    value: Optional[str],
    confidence: float,
    block_index: int,
    para_index: int,
    node_ids: list[str],
    raw_label: Optional[str],
    raw_value: Optional[str],
) -> dict:
    return {
        "label": label,
        "type": field_type,
        "value": value,
        "required": False,
        "confidence": round(confidence, 3),
        "source": {
            "block": block_index,
            "para": para_index,
            "nodeIds": node_ids,
        },
        "raw": {
            "labelText": raw_label,
            "valueText": raw_value,
        },
    }


def _field_id(index: int) -> str:
    return f"field_{index:03d}"


def _parse_ai_response(raw: str) -> dict:
    raw = raw.strip()
    if raw.startswith("```"):
        lines = raw.splitlines()
        inner = lines[1:]
        if inner and inner[-1].strip() == "```":
            inner = inner[:-1]
        raw = "\n".join(inner)

    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        m = re.search(r"\{.*\}", raw, re.DOTALL)
        if not m:
            return {}
        try:
            return json.loads(m.group())
        except json.JSONDecodeError:
            return {}


def _normalize_ai_fields(payload: dict) -> list[dict]:
    fields = payload.get("fields")
    if not isinstance(fields, list):
        return []
    output: list[dict] = []
    for item in fields:
        if not isinstance(item, dict):
            continue
        label = str(item.get("label") or "").strip()
        if not label:
            continue
        field_type = str(item.get("type") or "text").strip()
        if field_type not in {"text", "date", "checkbox", "choice", "number", "multiline"}:
            field_type = "text"
        block = item.get("block")
        para = item.get("para")
        if not isinstance(block, int):
            block = -1
        if not isinstance(para, int):
            para = 0
        confidence = item.get("confidence")
        try:
            confidence = float(confidence)
        except (TypeError, ValueError):
            confidence = 0.45

        output.append(
            {
                "label": label,
                "type": field_type,
                "value": None,
                "required": False,
                "confidence": round(min(max(confidence, 0.0), 1.0), 3),
                "source": {
                    "block": block,
                    "para": para,
                    "nodeIds": [],
                },
                "raw": {
                    "labelText": item.get("labelText"),
                    "valueText": item.get("valueText"),
                },
            }
        )
    return output


def _merge_fields(rule_fields: list[dict], ai_fields: list[dict]) -> list[dict]:
    seen = set()
    merged: list[dict] = []

    for field in rule_fields:
        key = _field_merge_key(field)
        seen.add(key)
        merged.append(field)

    for field in ai_fields:
        key = _field_merge_key(field)
        if key in seen:
            continue
        seen.add(key)
        merged.append(field)

    # Re-assign sequential IDs
    for idx, field in enumerate(merged, start=1):
        field["id"] = _field_id(idx)

    return merged


def _field_merge_key(field: dict) -> str:
    label = _normalize_label(field.get("label") or "")
    src = field.get("source") or {}
    block = src.get("block")
    para = src.get("para")
    return f"{label}|{block}|{para}"


def _normalize_label(label: str) -> str:
    return re.sub(r"[^a-z0-9]+", " ", label.lower()).strip()
