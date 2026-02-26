def _cell_text(cell: dict) -> str:
    text = ""
    for paragraph in cell.get("content", []):
        for piece in paragraph.get("content", []):
            if piece.get("type") == "Text":
                text += piece.get("text", "")
    return text


def _table_col_count(rows: list[dict]) -> int:
    max_col = 0
    for row in rows:
        for cell in row.get("cells", []):
            start = int(cell.get("grid_col", 0))
            span = int(cell.get("col_span", 1))
            max_col = max(max_col, start + span)
    return max_col if max_col > 0 else 1


def render_table(doc, block: dict):
    rows = block.get("rows", [])
    if not rows:
        return

    col_count = _table_col_count(rows)
    table = doc.add_table(rows=len(rows), cols=col_count)

    for r_idx, row in enumerate(rows):
        for cell in row.get("cells", []):
            start_col = int(cell.get("grid_col", 0))
            col_span = int(cell.get("col_span", 1))
            row_span = int(cell.get("row_span", 1))

            if start_col >= col_count:
                continue

            end_col = min(col_count - 1, start_col + col_span - 1)
            end_row = min(len(rows) - 1, r_idx + row_span - 1)

            top_left = table.cell(r_idx, start_col)
            if end_col > start_col:
                top_left = top_left.merge(table.cell(r_idx, end_col))
            if end_row > r_idx:
                top_left = top_left.merge(table.cell(end_row, start_col))
                if end_col > start_col:
                    top_left = top_left.merge(table.cell(end_row, end_col))

            top_left.text = _cell_text(cell)
