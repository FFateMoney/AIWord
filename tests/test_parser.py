from pathlib import Path

from docx import Document

from word_ast.parser.document_parser import parse_docx


def test_parse_character_format_bold(tmp_path: Path):
    path = tmp_path / "bold.docx"
    doc = Document()
    p = doc.add_paragraph()
    p.add_run("normal ")
    run = p.add_run("bold")
    run.bold = True
    doc.save(path)

    ast = parse_docx(path)
    runs = ast["document"]["body"][0]["content"]
    bold_run = next(r for r in runs if r.get("overrides", {}).get("bold"))
    assert bold_run["overrides"]["bold"] is True


def test_parse_styles_type_is_normalized(tmp_path: Path):
    path = tmp_path / "styles.docx"
    doc = Document()
    doc.add_paragraph("hello")
    doc.save(path)

    ast = parse_docx(path)
    styles = ast["document"]["styles"]
    assert styles["Normal"]["type"] == "paragraph"


def test_parse_merged_table_spans(tmp_path: Path):
    path = tmp_path / "merged.docx"
    doc = Document()
    table = doc.add_table(rows=3, cols=3)
    for row_idx in range(3):
        for col_idx in range(3):
            table.cell(row_idx, col_idx).text = f"{row_idx},{col_idx}"

    table.cell(0, 0).merge(table.cell(0, 1))
    table.cell(1, 2).merge(table.cell(2, 2))
    doc.save(path)

    ast = parse_docx(path)
    table_ast = next(block for block in ast["document"]["body"] if block["type"] == "Table")

    merged_h = table_ast["rows"][0]["cells"][0]
    assert merged_h["col_span"] == 2
    assert merged_h["row_span"] == 1

    merged_v = table_ast["rows"][1]["cells"][2]
    assert merged_v["col_span"] == 1
    assert merged_v["row_span"] == 2

    assert len(table_ast["rows"][2]["cells"]) == 2
