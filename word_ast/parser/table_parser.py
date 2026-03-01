from docx.oxml.ns import qn
from docx.table import Table
from docx.table import _Cell
from lxml import etree

from word_ast.parser.paragraph_parser import parse_paragraph_block


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

    # Capture the raw table-level properties so the renderer can restore
    # alignment, width, borders, and other tblPr attributes.
    raw_tblPr = None
    try:
        tblPr_el = table._tbl.tblPr
        if tblPr_el is not None:
            raw_tblPr = etree.tostring(tblPr_el, encoding="unicode")
    except (AttributeError, TypeError):
        pass

    rows = []
    xml_rows = table._tbl.tr_lst
    for row_idx, tr in enumerate(xml_rows):
        # Capture raw row properties (e.g. row height, tblHeader flag).
        raw_trPr = None
        try:
            trPr_el = tr.find(qn("w:trPr"))
            if trPr_el is not None:
                raw_trPr = etree.tostring(trPr_el, encoding="unicode")
        except (AttributeError, TypeError):
            pass

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
            cell_paragraphs = []
            for p_idx, p in enumerate(cell.paragraphs):
                p_block = parse_paragraph_block(p, f"{block_id}.r{row_idx}c{col_cursor}.p{p_idx}")
                cell_paragraphs.append(p_block)

            raw_tcPr = None
            try:
                tcPr_el = tc.tcPr
                if tcPr_el is not None:
                    raw_tcPr = etree.tostring(tcPr_el, encoding="unicode")
            except (AttributeError, TypeError):
                pass

            cell_data = {
                "id": f"{block_id}.r{row_idx}c{col_cursor}",
                "content": cell_paragraphs,
                "col_span": col_span,
                "row_span": row_span,
            }
            if raw_tcPr:
                cell_data["_raw_tcPr"] = raw_tcPr
            cells.append(cell_data)
            col_cursor += col_span
        row_data: dict = {"cells": cells}
        if raw_trPr:
            row_data["_raw_trPr"] = raw_trPr
        rows.append(row_data)
    block: dict = {"id": block_id, "type": "Table", "style": style_id, "rows": rows}
    if raw_tblPr:
        block["_raw_tblPr"] = raw_tblPr
    return block
