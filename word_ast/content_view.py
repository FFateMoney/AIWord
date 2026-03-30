"""Content-only document view for AI reading tasks."""

from __future__ import annotations

import re


_HEADING_RE = re.compile(r"heading\s*([1-9]\d*)", re.IGNORECASE)
_HEADING_ZH_RE = re.compile(r"标题\s*([1-9]\d*)")


def to_content_view(ast: dict) -> dict:
    """Return a content-focused JSON view with formatting stripped."""
    document = ast.get("document", {})
    styles = document.get("styles", {})
    body = document.get("body", [])
    return {
        "schema_version": ast.get("schema_version", "1.0"),
        "document": {
            "body": [_transform_block(block, styles) for block in body],
        },
    }


def _transform_block(block: dict, styles: dict) -> dict:
    block_type = block.get("type")
    if block_type == "Paragraph":
        return _transform_paragraph(block, styles)
    if block_type == "Table":
        return _transform_table(block, styles)
    if block_type == "TOC":
        return _transform_toc(block, styles)
    return {
        "id": block.get("id"),
        "type": block_type,
    }


def _transform_paragraph(block: dict, styles: dict) -> dict:
    content = [_transform_inline(item) for item in block.get("content", [])]
    text = "".join(item.get("text", "") for item in content if item.get("type") == "Text")
    heading_level = _heading_level(block.get("style"), styles)

    result = {
        "id": block.get("id"),
        "type": "Heading" if heading_level is not None else "Paragraph",
        "text": text,
        "content": content,
    }
    if heading_level is not None:
        result["level"] = heading_level
    return result


def _transform_inline(item: dict) -> dict:
    item_type = item.get("type")
    if item_type == "Text":
        return {
            "type": "Text",
            "text": item.get("text", ""),
        }
    if item_type == "InlineImage":
        result = {
            "type": "InlineImage",
            "data": item.get("data", ""),
        }
        for key in ("content_type", "width", "height", "alt"):
            if key in item:
                result[key] = item[key]
        return result
    if item_type == "Hyperlink":
        content = [_transform_inline(child) for child in item.get("content", [])]
        result = {
            "type": "Hyperlink",
            "content": content,
            "text": "".join(
                child.get("text", "") for child in content if child.get("type") == "Text"
            ),
        }
        if "url" in item:
            result["url"] = item["url"]
        return result
    result = {"type": item_type}
    if "text" in item:
        result["text"] = item["text"]
    if "content" in item:
        result["content"] = [_transform_inline(child) for child in item.get("content", [])]
    return result


def _transform_table(block: dict, styles: dict) -> dict:
    rows = []
    for row in block.get("rows", []):
        cells = []
        for cell in row.get("cells", []):
            content = [_transform_block(child, styles) for child in cell.get("content", [])]
            cells.append(
                {
                    "id": cell.get("id"),
                    "text": "\n".join(
                        child.get("text", "") for child in content if child.get("text")
                    ),
                    "col_span": cell.get("col_span", 1),
                    "row_span": cell.get("row_span", 1),
                    "content": content,
                }
            )
        rows.append({"cells": cells})
    return {
        "id": block.get("id"),
        "type": "Table",
        "rows": rows,
    }


def _transform_toc(block: dict, styles: dict) -> dict:
    result = {
        "id": block.get("id"),
        "type": "TOC",
    }
    title = block.get("title")
    if isinstance(title, dict):
        title_block = _transform_paragraph(title, styles)
        result["title"] = title_block
        result["text"] = title_block.get("text", "")
    return result


def _heading_level(style_id: str | None, styles: dict) -> int | None:
    if not style_id:
        return None

    style_def = styles.get(style_id, {}) if isinstance(styles, dict) else {}
    candidates = [style_id]
    if isinstance(style_def, dict):
        candidates.extend(
            value
            for value in (style_def.get("style_id"), style_def.get("name"))
            if isinstance(value, str)
        )

    for candidate in candidates:
        match = _HEADING_RE.search(candidate)
        if match:
            return int(match.group(1))
        match = _HEADING_ZH_RE.search(candidate)
        if match:
            return int(match.group(1))
    return None
