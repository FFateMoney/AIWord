import json
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn
from docx.shared import Twips


from .paragraph_renderer import render_paragraph
from .style_renderer import render_styles
from .table_renderer import render_table

_HEADING_STYLE_NAMES = frozenset(
    f"heading {i}" for i in range(1, 10)
)


def _remove_heading_colors(doc):
    """Remove the blue theme color from built-in heading styles.

    The default python-docx template defines heading styles with blue accent
    colors.  Chinese Word documents normally use black headings, so we strip
    the ``<w:color>`` element from every heading style (both paragraph and
    linked character styles) to let them inherit the default text color.
    """
    styles_element = doc.styles.element
    for style_el in styles_element.iterchildren(qn("w:style")):
        name_el = style_el.find(qn("w:name"))
        if name_el is None:
            continue
        name_val = name_el.get(qn("w:val"), "")
        # Match "heading 1" … "heading 9" and their linked Char styles
        if name_val.lower() not in _HEADING_STYLE_NAMES and not any(
            name_val.lower() == f"heading {i} char" for i in range(1, 10)
        ):
            continue
        rPr = style_el.find(qn("w:rPr"))
        if rPr is None:
            continue
        color = rPr.find(qn("w:color"))
        if color is not None:
            rPr.remove(color)


def _set_compat_mode_15(doc):
    """Set ``compatibilityMode`` to 15 (Word 2013+).

    The default python-docx template ships with ``compatibilityMode`` 14
    (Word 2010), which causes modern Word to open the file in compatibility
    mode.
    """
    settings = doc.settings.element
    compat = settings.find(qn("w:compat"))
    if compat is None:
        return
    uri = "http://schemas.microsoft.com/office/word"
    for cs in compat.iterchildren(qn("w:compatSetting")):
        if (
            cs.get(qn("w:name")) == "compatibilityMode"
            and cs.get(qn("w:uri")) == uri
        ):
            cs.set(qn("w:val"), "15")
            return


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
    _remove_heading_colors(doc)
    _set_compat_mode_15(doc)
    styles = ast["document"].get("styles", {})
    render_styles(doc, styles)
    _render_meta(doc, ast["document"].get("meta", {}))

    body = ast["document"].get("body", [])
    for block in body:
        t = block.get("type")
        if t == "Paragraph":
            render_paragraph(doc, block, styles)
        elif t == "Table":
            render_table(doc, block)

    doc.save(str(output_path))
