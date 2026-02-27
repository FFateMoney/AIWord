from docx.table import Table
from docx.table import _Cell


def _grid_span(tc) -> int:
    tc_pr = tc.tcPr
    if tc_pr is not None and tc_pr.gridSpan is not None and tc_pr.gridSpan.val is not None:
        return int(tc_pr.gridSpan.val)
    return 1


def _v_merge(tc) -> str | None:
    tc_pr = tc.tcPr
    if tc_pr is None or tc_pr.vMerge is None:
        return None
    return tc_pr.vMerge.val or "continue"


def _tc_at_column(tr, col_idx: int):
    cursor = 0
    for tc in tr.tc_lst:
        if cursor == col_idx:
            return tc
        cursor += _grid_span(tc)
    return None


def parse_table_block(table: Table, block_id: str) -> dict:
    style_id = table.style.style_id if table.style else None
    rows = []
    xml_rows = table._tbl.tr_lst
    for row_idx, tr in enumerate(xml_rows):
        cells = []
        col_cursor = 0
        for tc in tr.tc_lst:
            col_span = _grid_span(tc)
            v_merge = _v_merge(tc)
            if v_merge == "continue":
                col_cursor += col_span
                continue

            row_span = 1
            if v_merge == "restart":
                for next_row_idx in range(row_idx + 1, len(xml_rows)):
                    next_tc = _tc_at_column(xml_rows[next_row_idx], col_cursor)
                    if next_tc is None:
                        break
                    if _v_merge(next_tc) != "continue" or _grid_span(next_tc) != col_span:
                        break
                    row_span += 1

            cell = _Cell(tc, table)
            cell_text = "\n".join(p.text for p in cell.paragraphs)
            cells.append(
                {
                    "id": f"{block_id}.r{row_idx}c{col_cursor}",
                    "content": [{"id": f"{block_id}.r{row_idx}c{col_cursor}.p0", "type": "Paragraph", "style": "Normal", "content": [{"type": "Text", "text": cell_text}]}],
                    "col_span": col_span,
                    "row_span": row_span,
                }
            )
            col_cursor += col_span
        rows.append({"cells": cells})
    return {"id": block_id, "type": "Table", "style": style_id, "rows": rows}
