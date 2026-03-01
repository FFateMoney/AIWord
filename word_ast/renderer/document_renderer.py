import json
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn
from docx.shared import Twips
from lxml import etree


from .paragraph_renderer import render_paragraph
from .style_renderer import render_styles
from .table_renderer import render_table
from .toc_renderer import render_toc

_HEADING_STYLE_NAMES = frozenset(
    f"heading {i}" for i in range(1, 10)
)

_HEADING_CHAR_NAMES = frozenset(
    f"heading {i} char" for i in range(1, 10)
)


def _is_heading_style_name(name: str) -> bool:
    low = name.lower()
    return low in _HEADING_STYLE_NAMES or low in _HEADING_CHAR_NAMES


def _strip_heading_colors_from_element(styles_element):
    """Remove ``<w:color>`` from heading styles in a styles XML element."""
    for style_el in styles_element.iterchildren(qn("w:style")):
        name_el = style_el.find(qn("w:name"))
        if name_el is None:
            continue
        if not _is_heading_style_name(name_el.get(qn("w:val"), "")):
            continue
        rPr = style_el.find(qn("w:rPr"))
        if rPr is None:
            continue
        color = rPr.find(qn("w:color"))
        if color is not None:
            rPr.remove(color)


def _remove_heading_colors(doc):
    """Remove the blue theme color from built-in heading styles.

    The default python-docx template defines heading styles with blue accent
    colors.  Chinese Word documents normally use black headings, so we strip
    the ``<w:color>`` element from every heading style (both paragraph and
    linked character styles) to let them inherit the default text color.

    Both ``styles.xml`` and ``stylesWithEffects.xml`` are cleaned because
    some Word versions read heading colours from the latter.
    """
    # 1. Clean styles.xml (exposed via python-docx API)
    _strip_heading_colors_from_element(doc.styles.element)

    # 2. Clean stylesWithEffects.xml (only accessible via the raw OPC part)
    _STYLES_WITH_EFFECTS_REL = (
        "http://schemas.microsoft.com/office/2007/relationships/stylesWithEffects"
    )
    for rel in doc.part.rels.values():
        if rel.reltype == _STYLES_WITH_EFFECTS_REL:
            swe_part = rel.target_part
            swe_tree = etree.fromstring(swe_part.blob)
            _strip_heading_colors_from_element(swe_tree)
            swe_part._blob = etree.tostring(swe_tree, xml_declaration=True,
                                            encoding="UTF-8", standalone=True)
            break


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
    # Restore page dimensions when available
    if "width" in page:
        section.page_width = Twips(page["width"])
    if "height" in page:
        section.page_height = Twips(page["height"])
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
            render_table(doc, block, styles)
        elif t == "TOC":
            render_toc(doc, block, styles)

    doc.save(str(output_path))
