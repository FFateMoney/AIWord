def parse_styles(doc) -> dict:
    styles = {}
    for style in doc.styles:
        if style.type is None:
            continue
        styles[style.style_id] = {
            "style_id": style.style_id,
            "name": style.name,
            "type": str(style.type).split(".")[-1].lower(),
            "based_on": style.base_style.style_id if style.base_style else None,
        }
    return styles
