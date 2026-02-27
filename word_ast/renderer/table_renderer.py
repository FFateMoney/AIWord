def render_table(doc, block: dict):
    rows = block.get("rows", [])
    if not rows:
        return
    col_count = max((len(r.get("cells", [])) for r in rows), default=1)
    table = doc.add_table(rows=len(rows), cols=col_count)
    style_id = block.get("style")
    if style_id:
        try:
            table.style = style_id
        except (KeyError, ValueError):
            pass
    for r_idx, row in enumerate(rows):
        cells = row.get("cells", [])
        for c_idx, cell in enumerate(cells):
            if c_idx >= col_count:
                continue
            text = ""
            for p in cell.get("content", []):
                for piece in p.get("content", []):
                    if piece.get("type") == "Text":
                        text += piece.get("text", "")
            table.cell(r_idx, c_idx).text = text
