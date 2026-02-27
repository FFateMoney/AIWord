from docx.enum.dml import MSO_COLOR_TYPE
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

from word_ast.utils.units import pt_to_half_points


def _color_to_hex(color, *, skip_theme: bool = False) -> str | None:
    if color is None:
        return None
    if skip_theme and getattr(color, "type", None) == MSO_COLOR_TYPE.THEME:
        return None
    rgb = color.rgb
    if rgb is None:
        return None
    return f"#{rgb}"


def _read_east_asia_font(font) -> str | None:
    """Read the East Asian font name from the underlying XML element."""
    try:
        rPr = font._element.rPr
        if rPr is None:
            return None
        rFonts = rPr.find(qn('w:rFonts'))
        if rFonts is None:
            return None
        return rFonts.get(qn('w:eastAsia'))
    except (AttributeError, TypeError):
        return None


def _font_to_overrides(font, *, skip_theme_color: bool = False) -> dict:
    overrides = {}
    if font is None:
        return overrides

    if font.bold is not None:
        overrides["bold"] = font.bold
    if font.italic is not None:
        overrides["italic"] = font.italic
    if font.underline is not None:
        overrides["underline"] = bool(font.underline)

    color = _color_to_hex(font.color, skip_theme=skip_theme_color)
    if color:
        overrides["color"] = color

    size = pt_to_half_points(font.size.pt if font.size else None)
    if size is not None:
        overrides["size"] = size

    ascii_font = font.name
    ea_font = _read_east_asia_font(font)
    if ascii_font:
        overrides["font_ascii"] = ascii_font
    if ea_font:
        overrides["font_east_asia"] = ea_font

    return overrides


def _merge_runs(content: list[dict]) -> list[dict]:
    """Merge consecutive Text nodes that share identical overrides."""
    if not content:
        return content
    merged: list[dict] = [content[0]]
    for item in content[1:]:
        prev = merged[-1]
        if (
            prev["type"] == "Text"
            and item["type"] == "Text"
            and prev.get("overrides") == item.get("overrides")
        ):
            prev["text"] += item["text"]
        else:
            merged.append(item)
    return merged


def parse_paragraph_block(paragraph: Paragraph, block_id: str) -> dict:
    content = []
    for run in paragraph.runs:
        item: dict = {"type": "Text", "text": run.text}
        overrides = _font_to_overrides(run.font)
        if overrides:
            item["overrides"] = overrides
        content.append(item)
    content = _merge_runs(content)

    default_run = _font_to_overrides(
        getattr(paragraph.style, "font", None), skip_theme_color=True
    )

    block = {
        "id": block_id,
        "type": "Paragraph",
        "style": paragraph.style.style_id if paragraph.style else None,
        "content": content,
    }
    if default_run:
        block["default_run"] = default_run

    return block
