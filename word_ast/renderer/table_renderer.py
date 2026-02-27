from docx.oxml.ns import qn

from .paragraph_renderer import render_paragraph


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
        except Exception:
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
            # Remove default empty paragraph(s)
            for p_el in tc._element.findall(qn('w:p')):
                tc._element.remove(p_el)
            # Render each paragraph with full formatting
            for p_block in cell.get("content", []):
                render_paragraph(tc, p_block, styles)
