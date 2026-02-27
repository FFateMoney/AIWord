import base64
import io
import zipfile
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor, Inches
from lxml import etree

from word_ast import parse_docx, render_ast

# Minimal 1×1 transparent PNG used in image round-trip tests
_PNG_1X1 = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8"
    "z8BQDwADhQGAWjR9awAAAABJRU5ErkJggg=="
)


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


def _get_east_asia_font(run) -> str | None:
    """Helper to read the East Asian font name from a run's XML."""
    rPr = run._element.find(qn('w:rPr'))
    if rPr is None:
        return None
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        return None
    return rFonts.get(qn('w:eastAsia'))


def test_roundtrip_preserves_east_asian_font(tmp_path: Path):
    """East Asian font (w:eastAsia) must survive a parse → render round-trip."""
    src = tmp_path / "ea.docx"
    out = tmp_path / "ea-out.docx"

    doc = Document()
    p = doc.add_paragraph()
    run = p.add_run("你好世界")
    run.font.name = "Calibri"
    run.font.size = Pt(14)
    # Set East Asian font via XML
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    rFonts.set(qn('w:eastAsia'), '宋体')
    doc.save(src)

    ast = parse_docx(src)
    render_ast(ast, out)

    rebuilt = Document(out)
    r = rebuilt.paragraphs[0].runs[0]
    assert r.font.name == "Calibri"
    assert _get_east_asia_font(r) == "宋体"
    assert r.font.size.pt == 14


def test_roundtrip_multi_run_different_fonts(tmp_path: Path):
    """Multiple runs with different fonts/colors/sizes must not bleed into each other."""
    src = tmp_path / "multi.docx"
    out = tmp_path / "multi-out.docx"

    doc = Document()
    p = doc.add_paragraph()

    r1 = p.add_run("Red ")
    r1.font.color.rgb = RGBColor(255, 0, 0)
    r1.font.size = Pt(14)
    r1.font.name = "Arial"

    r2 = p.add_run("Blue")
    r2.font.color.rgb = RGBColor(0, 0, 255)
    r2.font.size = Pt(10)
    r2.font.name = "Times New Roman"

    doc.save(src)

    ast = parse_docx(src)
    render_ast(ast, out)

    rebuilt = Document(out)
    runs = rebuilt.paragraphs[0].runs
    assert len(runs) == 2

    assert runs[0].font.name == "Arial"
    assert runs[0].font.size.pt == 14
    assert str(runs[0].font.color.rgb) == "FF0000"

    assert runs[1].font.name == "Times New Roman"
    assert runs[1].font.size.pt == 10
    assert str(runs[1].font.color.rgb) == "0000FF"


def test_roundtrip_style_defaults_with_run_overrides(tmp_path: Path):
    """When a style sets defaults and a run overrides only some properties,
    the non-overridden properties must come from the style defaults."""
    src = tmp_path / "override.docx"
    out = tmp_path / "override-out.docx"

    doc = Document()
    heading = doc.styles["Heading 1"]
    heading.font.name = "Arial"
    heading.font.size = Pt(24)
    heading.font.color.rgb = RGBColor(0, 0, 128)
    heading.font.bold = True
    # Set East Asian font on style
    rPr = heading.font._element.rPr
    rFonts = rPr.find(qn('w:rFonts'))
    rFonts.set(qn('w:eastAsia'), '黑体')

    p = doc.add_paragraph()
    p.style = heading

    # Run 1: no overrides (inherit all from style)
    p.add_run("Title ")

    # Run 2: override only color
    r2 = p.add_run("RED")
    r2.font.color.rgb = RGBColor(255, 0, 0)

    # Run 3: override only ASCII font (east-asian font must come from style default)
    r3 = p.add_run(" Serif")
    r3.font.name = "Times New Roman"

    doc.save(src)

    ast = parse_docx(src)
    render_ast(ast, out)

    rebuilt = Document(out)
    runs = rebuilt.paragraphs[0].runs

    # Run 1: all from style defaults
    assert runs[0].font.name == "Arial"
    assert _get_east_asia_font(runs[0]) == "黑体"
    assert runs[0].bold is True
    assert runs[0].font.size.pt == 24
    assert str(runs[0].font.color.rgb) == "000080"

    # Run 2: color overridden, rest from defaults
    assert runs[1].font.name == "Arial"
    assert _get_east_asia_font(runs[1]) == "黑体"
    assert str(runs[1].font.color.rgb) == "FF0000"
    assert runs[1].font.size.pt == 24

    # Run 3: ASCII font overridden, east-asian still from defaults
    assert runs[2].font.name == "Times New Roman"
    assert _get_east_asia_font(runs[2]) == "黑体"
    assert runs[2].font.size.pt == 24


def _assert_no_heading_colors_in_styles(styles_element):
    """Assert no heading style in *styles_element* has a ``<w:color>``."""
    for style_el in styles_element.iterchildren(qn("w:style")):
        name_el = style_el.find(qn("w:name"))
        if name_el is None:
            continue
        name_val = name_el.get(qn("w:val"), "")
        if "heading" not in name_val.lower():
            continue
        rPr = style_el.find(qn("w:rPr"))
        if rPr is None:
            continue
        color = rPr.find(qn("w:color"))
        assert color is None, f"Style '{name_val}' should not have a <w:color> element"


def test_rendered_headings_have_no_blue_color(tmp_path: Path):
    """Heading styles in rendered documents must not carry the blue theme
    color from the default python-docx template.  Both ``styles.xml`` and
    ``stylesWithEffects.xml`` are checked."""
    out = tmp_path / "heading-color.docx"
    ast = {
        "schema_version": "1.0",
        "document": {
            "meta": {},
            "styles": {
                "Heading1": {"style_id": "Heading1", "name": "Heading 1", "type": "paragraph", "based_on": "Normal"},
            },
            "body": [
                {"id": "p0", "type": "Paragraph", "style": "Heading1",
                 "content": [{"type": "Text", "text": "Title"}]},
            ],
            "passthrough": {},
        },
    }
    render_ast(ast, out)

    # Check styles.xml via python-docx API
    rebuilt = Document(out)
    _assert_no_heading_colors_in_styles(rebuilt.styles.element)

    # Check stylesWithEffects.xml via raw ZIP
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    with zipfile.ZipFile(out) as zf:
        if "word/stylesWithEffects.xml" in zf.namelist():
            swe_root = etree.fromstring(zf.read("word/stylesWithEffects.xml"))
            _assert_no_heading_colors_in_styles(swe_root)


def test_rendered_compat_mode_is_15(tmp_path: Path):
    """Rendered documents should use compatibilityMode 15 (Word 2013+)
    so that modern Word does not enter compatibility mode."""
    out = tmp_path / "compat.docx"
    ast = {
        "schema_version": "1.0",
        "document": {
            "meta": {},
            "styles": {},
            "body": [
                {"id": "p0", "type": "Paragraph", "style": None,
                 "content": [{"type": "Text", "text": "hello"}]},
            ],
            "passthrough": {},
        },
    }
    render_ast(ast, out)

    rebuilt = Document(out)
    settings_el = rebuilt.settings.element
    compat = settings_el.find(qn("w:compat"))
    assert compat is not None
    for cs in compat.iterchildren(qn("w:compatSetting")):
        if (
            cs.get(qn("w:name")) == "compatibilityMode"
            and cs.get(qn("w:uri")) == "http://schemas.microsoft.com/office/word"
        ):
            assert cs.get(qn("w:val")) == "15"
            return
    raise AssertionError("compatibilityMode setting not found")


def test_roundtrip_heading_theme_color_not_reapplied(tmp_path: Path):
    """Round-tripping a document whose heading styles use a theme color must
    NOT re-apply that color as an explicit run-level override.  The default
    python-docx template defines headings with blue theme colours; after a
    round-trip the heading runs should carry no explicit ``<w:color>`` so
    that the cleaned-up style (which has no colour) determines the text
    colour (black / auto).
    """
    src = tmp_path / "theme-heading.docx"
    out = tmp_path / "theme-heading-out.docx"

    doc = Document()
    h1 = doc.add_paragraph("Title Level 1")
    h1.style = "Heading 1"
    h2 = doc.add_paragraph("Title Level 2")
    h2.style = "Heading 2"
    doc.save(src)

    ast = parse_docx(src)

    # The parser must NOT capture the template theme colour in default_run
    for block in ast["document"]["body"]:
        if block.get("style") in ("Heading1", "Heading2"):
            assert "color" not in block.get("default_run", {}), (
                f"default_run for {block['style']} should not contain a "
                f"theme-derived color, got {block.get('default_run')}"
            )

    render_ast(ast, out)

    rebuilt = Document(out)
    for para in rebuilt.paragraphs:
        if para.style.name in ("Heading 1", "Heading 2"):
            for run in para.runs:
                rPr = run._element.find(qn("w:rPr"))
                if rPr is not None:
                    color_el = rPr.find(qn("w:color"))
                    assert color_el is None, (
                        f"Run in '{para.style.name}' should not have an "
                        f"explicit <w:color>, but found val="
                        f"{color_el.get(qn('w:val')) if color_el is not None else None}"
                    )


def test_roundtrip_table_style_preserved(tmp_path: Path):
    """Table style (e.g. 'Table Grid') must survive a parse → render round-trip
    so that cell borders remain visible in the rendered document."""
    src = tmp_path / "table_style.docx"
    out = tmp_path / "table_style_out.docx"

    doc = Document()
    table = doc.add_table(rows=2, cols=2)
    table.style = "Table Grid"
    table.cell(0, 0).text = "A"
    table.cell(0, 1).text = "B"
    table.cell(1, 0).text = "C"
    table.cell(1, 1).text = "D"
    doc.save(src)

    ast = parse_docx(src)
    table_ast = next(b for b in ast["document"]["body"] if b["type"] == "Table")
    assert table_ast.get("style") == "TableGrid", (
        f"Expected style 'TableGrid', got {table_ast.get('style')!r}"
    )

    render_ast(ast, out)

    rebuilt = Document(out)
    assert rebuilt.tables[0].style.name == "Table Grid"


def test_roundtrip_inline_image_preserved(tmp_path: Path):
    """Inline images embedded in paragraphs must survive a parse → render
    round-trip; the rendered document must contain exactly one inline shape."""
    src = tmp_path / "image.docx"
    out = tmp_path / "image_out.docx"

    doc = Document()
    para = doc.add_paragraph()
    run = para.add_run()
    run.add_picture(io.BytesIO(_PNG_1X1), width=Inches(1), height=Inches(1))
    doc.save(src)

    ast = parse_docx(src)
    para_ast = ast["document"]["body"][0]
    image_nodes = [c for c in para_ast["content"] if c.get("type") == "InlineImage"]
    assert len(image_nodes) == 1, "Parser must produce exactly one InlineImage node"
    node = image_nodes[0]
    assert "data" in node and node["data"]
    assert node.get("width", 0) > 0
    assert node.get("height", 0) > 0

    render_ast(ast, out)

    rebuilt = Document(out)
    assert len(rebuilt.inline_shapes) == 1, (
        "Rendered document must contain exactly one inline shape"
    )


def test_table_style_resolved_by_name_when_id_differs(tmp_path: Path):
    """When the AST table style ID does not exist in the python-docx template
    (e.g. 'a8' in Chinese Word documents mapping to 'Table Grid'), the renderer
    must fall back to looking up the style by name via the styles dict."""
    out = tmp_path / "table_name.docx"

    ast = {
        "schema_version": "1.0",
        "document": {
            "meta": {},
            "styles": {
                "a8": {"style_id": "a8", "name": "Table Grid", "type": "table", "based_on": None},
            },
            "body": [
                {
                    "id": "t0", "type": "Table", "style": "a8",
                    "rows": [
                        {"cells": [
                            {"id": "t0.r0c0", "content": [
                                {"id": "t0.r0c0.p0", "type": "Paragraph", "style": None,
                                 "content": [{"type": "Text", "text": "cell"}]}
                            ], "col_span": 1, "row_span": 1},
                        ]},
                    ],
                },
            ],
            "passthrough": {},
        },
    }

    render_ast(ast, out)

    rebuilt = Document(out)
    assert rebuilt.tables[0].style.name == "Table Grid"


def test_roundtrip_table_cell_font_preserved(tmp_path: Path):
    """Table cell text with explicit East Asian font must preserve the font
    through a parse → render round-trip so that CJK characters render correctly."""
    src = tmp_path / "cell_font.docx"
    out = tmp_path / "cell_font_out.docx"

    doc = Document()
    table = doc.add_table(rows=1, cols=1)
    table.style = "Table Grid"
    cell = table.cell(0, 0)
    # Clear default paragraph and add one with styled run
    p = cell.paragraphs[0]
    run = p.add_run("发展")
    run.font.name = "宋体"
    run.font.size = Pt(12)
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        from docx.oxml import OxmlElement
        rFonts = OxmlElement('w:rFonts')
        rPr.append(rFonts)
    rFonts.set(qn('w:eastAsia'), '宋体')
    doc.save(src)

    ast = parse_docx(src)
    render_ast(ast, out)

    rebuilt = Document(out)
    runs = rebuilt.tables[0].cell(0, 0).paragraphs[0].runs
    text_runs = [r for r in runs if r.text]
    assert len(text_runs) >= 1
    r = text_runs[0]
    assert r.text == "发展"
    assert r.font.name == "宋体"
    assert _get_east_asia_font(r) == "宋体"
