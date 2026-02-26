from docx.shared import RGBColor, Pt

from word_ast.utils.units import half_points_to_pt


def _apply_paragraph_style(paragraph, style_id: str | None, styles: dict | None):
    if not style_id:
        return

    candidates = []
    if isinstance(styles, dict):
        style_def = styles.get(style_id)
        style_name = style_def.get("name") if isinstance(style_def, dict) else None
        if style_name:
            candidates.append(style_name)
    candidates.append(style_id)

    for candidate in candidates:
        try:
            paragraph.style = candidate
            return
        except Exception:
            continue


def render_paragraph(doc, block: dict, styles: dict | None = None):
    paragraph = doc.add_paragraph()
    _apply_paragraph_style(paragraph, block.get("style"), styles)

    paragraph_defaults = block.get("default_run", {})
    for piece in block.get("content", []):
        if piece.get("type") != "Text":
            continue
        run = paragraph.add_run(piece.get("text", ""))
        overrides = {**paragraph_defaults, **piece.get("overrides", {})}
        if "bold" in overrides:
            run.bold = overrides["bold"]
        if "italic" in overrides:
            run.italic = overrides["italic"]
        if "underline" in overrides:
            run.underline = overrides["underline"]
        if "size" in overrides:
            size_pt = half_points_to_pt(overrides["size"])
            if size_pt is not None:
                run.font.size = Pt(size_pt)
        if "color" in overrides and overrides["color"].startswith("#"):
            hex_color = overrides["color"][1:]
            if len(hex_color) == 6:
                run.font.color.rgb = RGBColor.from_string(hex_color)
        font = overrides.get("font")
        if isinstance(font, dict) and font.get("ascii"):
            run.font.name = font["ascii"]
