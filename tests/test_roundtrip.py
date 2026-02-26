from pathlib import Path

from docx import Document
from docx.shared import Pt, RGBColor

from word_ast import parse_docx, render_ast


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


def test_render_uses_style_name_fallback_for_style_id(tmp_path: Path):
    out = tmp_path / "out-heading.docx"

    ast = {
        "schema_version": "1.0",
        "document": {
            "meta": {},
            "styles": {
                "1": {"style_id": "1", "name": "Heading 1", "type": "paragraph", "based_on": "Normal"},
                "a": {"style_id": "a", "name": "Normal", "type": "paragraph", "based_on": None},
            },
            "body": [
                {"id": "p0", "type": "Paragraph", "style": "1", "content": [{"type": "Text", "text": "Title"}]},
                {"id": "p1", "type": "Paragraph", "style": "a", "content": [{"type": "Text", "text": "Body"}]},
            ],
            "passthrough": {},
        },
    }

    render_ast(ast, out)

    rebuilt = Document(out)
    assert rebuilt.paragraphs[0].style.name == "Heading 1"
    assert rebuilt.paragraphs[1].style.name == "Normal"


def test_roundtrip_preserves_paragraph_style_font_defaults(tmp_path: Path):
    src = tmp_path / "styled.docx"
    out = tmp_path / "styled-out.docx"

    doc = Document()
    heading_style = doc.styles["Heading 1"]
    heading_style.font.name = "Arial"
    heading_style.font.size = Pt(24)
    heading_style.font.color.rgb = RGBColor(0, 0, 0)
    heading_style.font.bold = True

    heading = doc.add_paragraph("Styled title")
    heading.style = heading_style
    doc.save(src)

    ast = parse_docx(src)
    render_ast(ast, out)

    rebuilt = Document(out)
    run = rebuilt.paragraphs[0].runs[0]
    assert run.font.name == "Arial"
    assert run.font.size.pt == 24
    assert str(run.font.color.rgb) == "000000"
    assert run.bold is True
