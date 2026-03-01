import copy

from docx.oxml.ns import qn
from docx.oxml.parser import parse_xml

from .paragraph_renderer import render_paragraph


def _apply_raw_tcPr(tc_element, raw_tcPr: str) -> None:
    try:
        new_tcPr = parse_xml(raw_tcPr)
    except Exception:
        return
    old_tcPr = tc_element.find(qn("w:tcPr"))
    if old_tcPr is not None:
        tc_element.remove(old_tcPr)
    tc_element.insert(0, new_tcPr)


def _apply_raw_tblPr(tbl_element, raw_tblPr: str) -> None:
    """Replace the table's <w:tblPr> with the round-tripped raw XML.

    tblStyle is removed from the incoming raw_tblPr to avoid overriding
    the style already set by _apply_table_style; all other table properties
    (width, alignment, borders, etc.) are preserved.
    """
    try:
        new_tblPr = parse_xml(raw_tblPr)
    except Exception:
        return
    # Remove tblStyle from raw so it doesn't conflict with the style already
    # applied via the python-docx API.
    tbl_style_el = new_tblPr.find(qn("w:tblStyle"))
    if tbl_style_el is not None:
        new_tblPr.remove(tbl_style_el)
    # Preserve the tblStyle that _apply_table_style wrote.
    old_tblPr = tbl_element.find(qn("w:tblPr"))
    if old_tblPr is not None:
        existing_style = old_tblPr.find(qn("w:tblStyle"))
        if existing_style is not None:
            new_tblPr.insert(0, copy.deepcopy(existing_style))
        tbl_element.remove(old_tblPr)
    tbl_element.insert(0, new_tblPr)


def _apply_raw_trPr(tr_element, raw_trPr: str) -> None:
    """Replace a table row's <w:trPr> with the round-tripped raw XML."""
    try:
        new_trPr = parse_xml(raw_trPr)
    except Exception:
        return
    old_trPr = tr_element.find(qn("w:trPr"))
    if old_trPr is not None:
        tr_element.remove(old_trPr)
    tr_element.insert(0, new_trPr)



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
    # Apply raw table properties after setting the style so that structural
    # attributes (width, alignment, borders) are restored while tblStyle is
    # kept consistent with the style already applied above.
    if "_raw_tblPr" in block:
        _apply_raw_tblPr(table._tbl, block["_raw_tblPr"])
    for r_idx, row in enumerate(rows):
        tr_element = table.rows[r_idx]._tr
        if "_raw_trPr" in row:
            _apply_raw_trPr(tr_element, row["_raw_trPr"])
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
