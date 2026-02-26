from docx.text.paragraph import Paragraph

from word_ast.utils.units import pt_to_half_points


def _color_to_hex(color) -> str | None:
    if color is None:
        return None
    rgb = color.rgb
    if rgb is None:
        return None
    return f"#{rgb}"


def _font_to_overrides(font) -> dict:
    overrides = {}
    if font is None:
        return overrides

    if font.bold is not None:
        overrides["bold"] = font.bold
    if font.italic is not None:
        overrides["italic"] = font.italic
    if font.underline is not None:
        overrides["underline"] = bool(font.underline)

    color = _color_to_hex(font.color)
    if color:
        overrides["color"] = color

    size = pt_to_half_points(font.size.pt if font.size else None)
    if size is not None:
        overrides["size"] = size

    if font.name:
        overrides["font"] = {"ascii": font.name, "eastAsia": font.name}

    return overrides


def parse_paragraph_block(paragraph: Paragraph, block_id: str) -> dict:
    content = []
    for run in paragraph.runs:
        item: dict = {"type": "Text", "text": run.text}
        overrides = _font_to_overrides(run.font)
        if overrides:
            item["overrides"] = overrides
        content.append(item)

    default_run = _font_to_overrides(getattr(paragraph.style, "font", None))

    block = {
        "id": block_id,
        "type": "Paragraph",
        "style": paragraph.style.style_id if paragraph.style else None,
        "content": content,
    }
    if default_run:
        block["default_run"] = default_run

    return block
