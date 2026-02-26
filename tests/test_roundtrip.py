from pathlib import Path

import pytest

from word_ast.renderer.table_renderer import _table_col_count
from word_ast import parse_docx, render_ast

try:
    from docx import Document

    HAS_DOCX = True
except ModuleNotFoundError:
    HAS_DOCX = False


def test_table_col_count_with_spans():
    rows = [
        {"cells": [{"grid_col": 0, "col_span": 2}, {"grid_col": 2, "col_span": 1}]},
        {"cells": [{"grid_col": 0, "col_span": 1}, {"grid_col": 1, "col_span": 1}, {"grid_col": 2, "col_span": 1}]},
    ]
    assert _table_col_count(rows) == 3


@pytest.mark.skipif(not HAS_DOCX, reason="python-docx is not installed")
def test_roundtrip_text_and_table(tmp_path: Path):
    src = tmp_path / "src.docx"
    out = tmp_path / "out.docx"

    doc = Document()
    p = doc.add_paragraph()
    run = p.add_run("Hello")
    run.bold = True
    table = doc.add_table(rows=1, cols=2)
    table.cell(0, 0).text = "A"
    table.cell(0, 1).text = "B"
    doc.save(src)

    ast = parse_docx(src)
    assert ast["schema_version"] == "1.0"
    assert ast["document"]["body"][0]["type"] == "Paragraph"
    assert ast["document"]["body"][1]["type"] == "Table"

    render_ast(ast, out)

    rebuilt = Document(out)
    assert rebuilt.paragraphs[0].text == "Hello"
    assert rebuilt.tables[0].cell(0, 0).text == "A"
    assert rebuilt.tables[0].cell(0, 1).text == "B"
