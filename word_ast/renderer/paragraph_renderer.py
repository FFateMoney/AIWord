from docx.shared import RGBColor, Pt

from word_ast.utils.units import half_points_to_pt


def render_paragraph(doc, block: dict):
    paragraph = doc.add_paragraph()
    if block.get("style"):
        try:
            paragraph.style = block["style"]
        except Exception:
            pass

    for piece in block.get("content", []):
        if piece.get("type") != "Text":
            continue
        run = paragraph.add_run(piece.get("text", ""))
        overrides = piece.get("overrides", {})
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
