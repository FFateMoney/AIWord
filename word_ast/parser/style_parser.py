def parse_styles(doc) -> dict:
    def _normalize_style_type(style_type) -> str:
        if hasattr(style_type, "name") and style_type.name:
            return style_type.name.lower()
        return str(style_type).split()[0].lower()

    styles = {}
    for style in doc.styles:
        if style.type is None:
            continue
        styles[style.style_id] = {
            "style_id": style.style_id,
            "name": style.name,
            "type": _normalize_style_type(style.type),
            "based_on": getattr(style.base_style, "style_id", None) if hasattr(style, "base_style") else None,
        }
    return styles
