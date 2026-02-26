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
