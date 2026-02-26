from docx.text.paragraph import Paragraph

from word_ast.utils.units import pt_to_half_points


def _color_to_hex(color) -> str | None:
    if color is None:
        return None
    rgb = color.rgb
    if rgb is None:
        return None
    return f"#{rgb}"


def parse_paragraph_block(paragraph: Paragraph, block_id: str) -> dict:
    content = []
    for run in paragraph.runs:
        item: dict = {"type": "Text", "text": run.text}
        overrides = {}
        if run.bold is not None:
            overrides["bold"] = run.bold
        if run.italic is not None:
            overrides["italic"] = run.italic
        if run.underline is not None:
            overrides["underline"] = bool(run.underline)
        color = _color_to_hex(run.font.color)
        if color:
            overrides["color"] = color
        size = pt_to_half_points(run.font.size.pt if run.font.size else None)
        if size is not None:
            overrides["size"] = size
        if run.font.name:
            overrides["font"] = {"ascii": run.font.name, "eastAsia": run.font.name}
        if overrides:
            item["overrides"] = overrides
        content.append(item)

    return {
        "id": block_id,
        "type": "Paragraph",
        "style": paragraph.style.style_id if paragraph.style else None,
        "content": content,
    }
