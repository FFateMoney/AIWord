import json
from pathlib import Path

from docx import Document
from docx.shared import Twips

from .paragraph_renderer import render_paragraph
from .style_renderer import render_styles
from .table_renderer import render_table


def _render_meta(doc, meta: dict):
    page = meta.get("page", {})
    margin = page.get("margin", {})
    section = doc.sections[0]
    for key, field in (("top_margin", "top"), ("bottom_margin", "bottom"), ("left_margin", "left"), ("right_margin", "right")):
        if field in margin:
            setattr(section, key, Twips(margin[field]))


def render_ast(ast_or_path: dict | str | Path, output_path: str | Path):
    if isinstance(ast_or_path, (str, Path)):
        ast = json.loads(Path(ast_or_path).read_text(encoding="utf-8"))
    else:
        ast = ast_or_path

    doc = Document()
    render_styles(doc, ast["document"].get("styles", {}))
    _render_meta(doc, ast["document"].get("meta", {}))

    body = ast["document"].get("body", [])
    for block in body:
        t = block.get("type")
        if t == "Paragraph":
            render_paragraph(doc, block)
        elif t == "Table":
            render_table(doc, block)

    doc.save(str(output_path))
