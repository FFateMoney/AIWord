from docx.table import Table


def parse_table_block(table: Table, block_id: str) -> dict:
    rows = []
    for row_idx, row in enumerate(table.rows):
        cells = []
        seen = set()
        for col_idx, cell in enumerate(row.cells):
            tc_id = id(cell._tc)
            if tc_id in seen:
                continue
            seen.add(tc_id)
            cell_text = "\n".join(p.text for p in cell.paragraphs)
            cells.append(
                {
                    "id": f"{block_id}.r{row_idx}c{col_idx}",
                    "content": [{"id": f"{block_id}.r{row_idx}c{col_idx}.p0", "type": "Paragraph", "style": "Normal", "content": [{"type": "Text", "text": cell_text}]}],
                    "col_span": 1,
                    "row_span": 1,
                }
            )
        rows.append({"cells": cells})
    return {"id": block_id, "type": "Table", "rows": rows}
