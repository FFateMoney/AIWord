from docx.oxml.ns import qn
from docx.table import Table


def _find_child(element, tag: str):
    return element.find(qn(tag))


def _findall(element, tag: str):
    return element.findall(qn(tag))


def _grid_span(cell) -> int:
    tc_pr = _find_child(cell._tc, "w:tcPr")
    if tc_pr is None:
        return 1
    grid_span = _find_child(tc_pr, "w:gridSpan")
    if grid_span is None:
        return 1
    val = grid_span.get(qn("w:val"))
    return int(val) if val and val.isdigit() else 1


def _is_vmerge_continuation(cell) -> bool:
    tc_pr = _find_child(cell._tc, "w:tcPr")
    if tc_pr is None:
        return False
    vmerge = _find_child(tc_pr, "w:vMerge")
    if vmerge is None:
        return False
    val = vmerge.get(qn("w:val"))
    return val == "continue" or val is None


def _vertical_span(table: Table, row_idx: int, col_idx: int) -> int:
    span = 1
    for next_row_idx in range(row_idx + 1, len(table.rows)):
        row = table.rows[next_row_idx]
        if col_idx >= len(row.cells):
            break
        next_cell = row.cells[col_idx]
        if id(next_cell._tc) == id(table.rows[row_idx].cells[col_idx]._tc):
            span += 1
            continue

        tc_pr = _find_child(next_cell._tc, "w:tcPr")
        if tc_pr is None:
            break
        vmerge = _find_child(tc_pr, "w:vMerge")
        if vmerge is None:
            break
        val = vmerge.get(qn("w:val"))
        if val == "continue" or val is None:
            span += 1
        else:
            break
    return span


def parse_table_block(table: Table, block_id: str) -> dict:
    rows = []
    for row_idx, row in enumerate(table.rows):
        cells = []
        seen = set()
        logical_col = 0
        for col_idx, cell in enumerate(row.cells):
            tc_id = id(cell._tc)
            if tc_id in seen:
                logical_col += 1
                continue
            seen.add(tc_id)

            if _is_vmerge_continuation(cell):
                logical_col += _grid_span(cell)
                continue

            col_span = _grid_span(cell)
            row_span = _vertical_span(table, row_idx, col_idx)
            cell_text = "\n".join(p.text for p in cell.paragraphs)
            cell_id = f"{block_id}.r{row_idx}c{logical_col}"
            cells.append(
                {
                    "id": cell_id,
                    "grid_col": logical_col,
                    "content": [
                        {
                            "id": f"{cell_id}.p0",
                            "type": "Paragraph",
                            "style": "Normal",
                            "content": [{"type": "Text", "text": cell_text}],
                        }
                    ],
                    "col_span": col_span,
                    "row_span": row_span,
                }
            )
            logical_col += col_span
        rows.append({"cells": cells})
    return {"id": block_id, "type": "Table", "rows": rows}
