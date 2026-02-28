from docx.oxml.ns import qn
from lxml import etree

from .paragraph_renderer import render_paragraph

_SAFE_XML_PARSER = etree.XMLParser(resolve_entities=False)


def _apply_raw_tcPr(tc_element, raw_tcPr: str) -> None:
    try:
        new_tcPr = etree.fromstring(raw_tcPr, _SAFE_XML_PARSER)
    except etree.XMLSyntaxError:
        return
    old_tcPr = tc_element.find(qn("w:tcPr"))
    if old_tcPr is not None:
        tc_element.remove(old_tcPr)
    tc_element.insert(0, new_tcPr)



def _apply_table_style(table, style_id: str | None, styles: dict | None):
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
            table.style = candidate
            return
        except (KeyError, ValueError):
            continue


def render_table(doc, block: dict, styles: dict | None = None):
    rows = block.get("rows", [])
    if not rows:
        return
    col_count = max((len(r.get("cells", [])) for r in rows), default=1)
    table = doc.add_table(rows=len(rows), cols=col_count)
    _apply_table_style(table, block.get("style"), styles)
    for r_idx, row in enumerate(rows):
        cells = row.get("cells", [])
        for c_idx, cell in enumerate(cells):
            if c_idx >= col_count:
                continue
            tc = table.cell(r_idx, c_idx)
            if "_raw_tcPr" in cell:
                _apply_raw_tcPr(tc._element, cell["_raw_tcPr"])
            # Remove default empty paragraph(s)
            for p_el in tc._element.findall(qn('w:p')):
                tc._element.remove(p_el)
            # Render each paragraph with full formatting
            for p_block in cell.get("content", []):
                render_paragraph(tc, p_block, styles)
