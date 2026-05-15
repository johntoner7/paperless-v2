"""Stable DOCX node ID injection helpers.

The importer and writer both use these helpers so the same DOCX traversal
order produces the same identity map. Only the node types that the pipeline
patches need IDs: paragraphs, runs, and checkbox containers.
"""

from __future__ import annotations

import re
from collections.abc import MutableMapping
from typing import Optional
from xml.etree import ElementTree as ET


NODE_ID_ATTR = "docx-node-id"
_NODE_PREFIX_BY_TAG = {
    "p": "p",
    "r": "r",
    "sdt": "sdt",
}
_NODE_ID_RE = re.compile(r"^(?P<prefix>[a-z]+)_(?P<index>\d+)$")


def inject_docx_node_ids(root: ET.Element, counters: Optional[MutableMapping[str, int]] = None) -> None:
    """Assign stable IDs to patchable DOCX nodes in document order.

    Existing IDs are preserved. Missing IDs are filled in with monotonically
    increasing identifiers like ``p_12`` or ``r_87``.
    """
    if counters is None:
        counters = {prefix: 0 for prefix in _NODE_PREFIX_BY_TAG.values()}

    for elem in root.iter():
        prefix = _NODE_PREFIX_BY_TAG.get(_local_name(elem.tag))
        if prefix is None:
            continue

        node_id = elem.get(NODE_ID_ATTR)
        if node_id:
            _bump_counter_from_existing_id(counters, node_id)
            continue

        next_index = counters.get(prefix, 0)
        node_id = f"{prefix}_{next_index}"
        elem.set(NODE_ID_ATTR, node_id)
        counters[prefix] = next_index + 1


def _bump_counter_from_existing_id(counters: MutableMapping[str, int], node_id: str) -> None:
    match = _NODE_ID_RE.match(node_id)
    if not match:
        return
    prefix = match.group("prefix")
    index = int(match.group("index"))
    current = counters.get(prefix, 0)
    if index + 1 > current:
        counters[prefix] = index + 1


def _local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]