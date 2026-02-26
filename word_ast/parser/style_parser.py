_STYLE_TYPE_MAP = {
    1: "paragraph",
    2: "character",
    3: "table",
    4: "numbering",
}


def _normalize_style_type(style_type) -> str:
    """Return a JSON-serializable lowercase style type."""
    # python-docx exposes WD_STYLE_TYPE enums; keep compatibility with older/newer reprs.
    value = getattr(style_type, "value", style_type)
    if isinstance(value, int):
        return _STYLE_TYPE_MAP.get(value, str(value).lower())

    text = str(style_type)
    if "." in text:
        text = text.split(".")[-1]
    return text.lower()


def parse_styles(doc) -> dict:
    styles = {}
    for style in doc.styles:
        if style.type is None:
            continue
        styles[style.style_id] = {
            "style_id": style.style_id,
            "name": style.name,
            "type": _normalize_style_type(style.type),
            "based_on": style.base_style.style_id if style.base_style else None,
        }
    return styles
