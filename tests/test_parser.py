from pathlib import Path

import pytest

from word_ast.parser.style_parser import _normalize_style_type

try:
    from docx import Document

    HAS_DOCX = True
except ModuleNotFoundError:
    HAS_DOCX = False

from word_ast import parse_docx


def test_normalize_style_type_is_json_safe():
    assert _normalize_style_type(1) == "paragraph"
    assert _normalize_style_type(2) == "character"
    assert _normalize_style_type("WD_STYLE_TYPE.TABLE") == "table"


@pytest.mark.skipif(not HAS_DOCX, reason="python-docx is not installed")
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
