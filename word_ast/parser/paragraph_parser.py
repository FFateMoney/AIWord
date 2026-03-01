import base64
import copy

from docx.enum.dml import MSO_COLOR_TYPE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
from docx.text.run import Run
from lxml import etree

from word_ast.utils.units import pt_to_half_points

_WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
_A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
_R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
# 1 EMU = 1/914400 inch; 1 twip = 1/1440 inch → 1 twip = 914400/1440 = 635 EMU
_EMU_PER_TWIP = 635


def _color_to_hex(color, *, skip_theme: bool = False) -> str | None:
    if color is None:
        return None
    if skip_theme and getattr(color, "type", None) == MSO_COLOR_TYPE.THEME:
        return None
    rgb = color.rgb
    if rgb is None:
        return None
    return f"#{rgb}"


def _read_east_asia_font(font) -> str | None:
    """Read the East Asian font name from the underlying XML element."""
    try:
        rPr = font._element.rPr
        if rPr is None:
            return None
        rFonts = rPr.find(qn('w:rFonts'))
        if rFonts is None:
            return None
        return rFonts.get(qn('w:eastAsia'))
    except (AttributeError, TypeError):
        return None


# Run properties that can be inherited from a style and should be captured in
# _raw_rPr even when not set directly on the run element.
_INHERITABLE_RPR_TAGS = {
    qn("w:rFonts"),    # 字体（ascii/eastAsia/hAnsi/cs）font family
    qn("w:sz"),        # 字号（半磅）font size in half-points
    qn("w:szCs"),      # 复杂文字字号 complex script font size
    qn("w:color"),     # 颜色（非主题色）font color (non-theme)
    qn("w:lang"),      # 语言 language
    qn("w:kern"),      # 字距 kerning
    qn("w:spacing"),   # 字符间距 character spacing
}


def _inherit_style_rPr(rPr_el, paragraph) -> None:  # rPr_el: lxml _Element
    """遍历 run 所在段落的样式继承链，将缺失的 <w:rPr> 子元素补入 rPr_el。

    Walk the paragraph's style inheritance chain and inject any <w:rPr> child
    elements that are absent from the run's own rPr.

    Mutates *rPr_el* in place by appending deep-copies of inheritable tags
    found on ancestor styles.  Tags already present on the run are never
    overwritten (run-level values take precedence over style values).

    This ensures that formatting defined on the style (e.g. a font family from
    "Heading 1") is captured in _raw_rPr even when the run itself carries no
    explicit rPr, preventing font substitution on round-trip.
    """
    present_tags = {child.tag for child in rPr_el}

    style = paragraph.style
    while style is not None:
        try:
            style_el = style.element
            style_rPr = style_el.rPr if style_el is not None else None
        except AttributeError:
            style_rPr = None
        if style_rPr is not None:
            for child in style_rPr:
                if child.tag in _INHERITABLE_RPR_TAGS and child.tag not in present_tags:
                    # Skip <w:color> with a w:themeColor attribute — theme colors
                    # must come from the style definition, not be materialised onto
                    # individual runs, to avoid corrupting the theme appearance.
                    if child.tag == qn("w:color") and child.get(qn("w:themeColor")):
                        continue
                    # deepcopy to avoid mutating the shared style XML
                    rPr_el.append(copy.deepcopy(child))
                    present_tags.add(child.tag)
        try:
            style = style.base_style
        except AttributeError:
            break


def _font_to_overrides(font, *, skip_theme_color: bool = False, paragraph=None) -> dict:
    """提取 run 的字体格式覆盖项，并序列化 _raw_rPr（含样式继承的完整信息）。

    Extract run-level font formatting overrides.  When *paragraph* is supplied
    the resulting ``_raw_rPr`` is enriched with properties inherited from the
    paragraph's style chain so that fonts and sizes are preserved on round-trip
    even when the run itself carries no explicit <w:rPr>.
    """
    overrides = {}
    if font is None:
        return overrides

    if font.bold is not None:
        overrides["bold"] = font.bold
    if font.italic is not None:
        overrides["italic"] = font.italic
    if font.underline is not None:
        overrides["underline"] = bool(font.underline)

    color = _color_to_hex(font.color, skip_theme=skip_theme_color)
    if color:
        overrides["color"] = color

    size = pt_to_half_points(font.size.pt if font.size else None)
    if size is not None:
        overrides["size"] = size

    # Read directly-set font names before any inheritance augmentation
    ascii_font = font.name
    ea_font = _read_east_asia_font(font)
    if ascii_font:
        overrides["font_ascii"] = ascii_font
    if ea_font:
        overrides["font_east_asia"] = ea_font

    try:
        rPr_el = font._element.rPr
        if rPr_el is None:
            if paragraph is not None:
                # 1d: run has no <w:rPr> — create a detached element, populate
                # via style inheritance, and store only if non-empty.
                rPr_el = OxmlElement("w:rPr")
                _inherit_style_rPr(rPr_el, paragraph)
                if len(rPr_el):
                    overrides["_raw_rPr"] = etree.tostring(rPr_el, encoding="unicode")
        else:
            if paragraph is not None:
                # Append inherited style properties absent from the run's own rPr
                _inherit_style_rPr(rPr_el, paragraph)
            overrides["_raw_rPr"] = etree.tostring(rPr_el, encoding="unicode")
    except (AttributeError, TypeError):
        pass

    return overrides


def _merge_runs(content: list[dict]) -> list[dict]:
    """Merge consecutive Text nodes that share identical overrides."""
    if not content:
        return content
    merged: list[dict] = [content[0]]
    for item in content[1:]:
        prev = merged[-1]
        if (
            prev["type"] == "Text"
            and item["type"] == "Text"
            and prev.get("overrides") == item.get("overrides")
        ):
            prev["text"] += item["text"]
        else:
            merged.append(item)
    return merged


def _parse_inline_image(run) -> dict | None:
    """Return an InlineImage node if *run* contains a ``<w:drawing>`` with an
    inline image, otherwise return ``None``."""
    r_el = run._element
    drawing = r_el.find(qn("w:drawing"))
    if drawing is None:
        return None
    inline = drawing.find(f"{{{_WP_NS}}}inline")
    if inline is None:
        return None
    ext = inline.find(f"{{{_WP_NS}}}extent")
    if ext is None:
        return None
    try:
        cx = int(ext.get("cx", 0))
        cy = int(ext.get("cy", 0))
    except (TypeError, ValueError):
        return None
    blip = inline.find(f".//{{{_A_NS}}}blip")
    if blip is None:
        return None
    r_id = blip.get(f"{{{_R_NS}}}embed")
    if not r_id:
        return None
    try:
        part = run.part
        image_part = part.related_parts[r_id]
        image_data = base64.b64encode(image_part.blob).decode("ascii")
        content_type = image_part.content_type
    except (KeyError, AttributeError):
        return None
    return {
        "type": "InlineImage",
        "data": image_data,
        "content_type": content_type,
        "width": cx // _EMU_PER_TWIP,
        "height": cy // _EMU_PER_TWIP,
    }


_ALIGNMENT_MAP = {0: "left", 1: "center", 2: "right", 3: "justify"}

# Paragraph properties that can be inherited from a style and should be
# captured in _raw_pPr even when not set directly on the paragraph element.
_INHERITABLE_PPR_TAGS = {
    qn("w:jc"),              # alignment (e.g. center)
    qn("w:ind"),             # indentation
    qn("w:spacing"),         # line/paragraph spacing
    qn("w:keepNext"),        # keep with next paragraph
    qn("w:keepLines"),       # keep lines together
    qn("w:pageBreakBefore"), # page break before paragraph
    qn("w:outlineLvl"),      # outline level
    qn("w:shd"),             # shading/background
    qn("w:pBdr"),            # paragraph border
}


def _inherit_style_pPr(pPr_el, paragraph) -> None:
    """Walk the paragraph's style inheritance chain and inject any <w:pPr>
    child elements that are absent from the paragraph's own pPr.

    Mutates *pPr_el* in place by appending deep-copies of inheritable tags
    found on ancestor styles.  Tags already present on the paragraph are never
    overwritten (paragraph-level values take precedence over style values).

    This ensures that formatting defined on the style (e.g. <w:jc> for center
    alignment) is captured in _raw_pPr even when python-docx does not surface
    it via ParagraphFormat properties (which only reflect directly-set values).
    """
    # Collect tags already present so we never overwrite paragraph-level values
    present_tags = {child.tag for child in pPr_el}

    style = paragraph.style
    while style is not None:
        try:
            style_pPr = style.element.pPr if style.element is not None else None
        except AttributeError:
            style_pPr = None
        if style_pPr is not None:
            for child in style_pPr:
                if child.tag in _INHERITABLE_PPR_TAGS and child.tag not in present_tags:
                    # deepcopy to avoid mutating the shared style XML
                    pPr_el.append(copy.deepcopy(child))
                    present_tags.add(child.tag)
        try:
            style = style.base_style
        except AttributeError:
            break


def _parse_paragraph_format(paragraph: Paragraph) -> dict:
    """Extract paragraph-level formatting (alignment, indentation, spacing)."""
    fmt: dict = {}
    pf = paragraph.paragraph_format

    if pf.alignment is not None:
        fmt["alignment"] = _ALIGNMENT_MAP.get(int(pf.alignment), "left")

    if pf.left_indent is not None:
        fmt["indent_left"] = pf.left_indent.twips
    if pf.right_indent is not None:
        fmt["indent_right"] = pf.right_indent.twips
    if pf.first_line_indent is not None:
        fmt["indent_first_line"] = pf.first_line_indent.twips

    if pf.space_before is not None:
        fmt["space_before"] = pf.space_before.twips
    if pf.space_after is not None:
        fmt["space_after"] = pf.space_after.twips

    try:
        pPr_el = paragraph._element.pPr
        if pPr_el is not None:
            # Inject inherited style properties that are absent from the
            # paragraph's own pPr (e.g. jc=center defined on a style).
            _inherit_style_pPr(pPr_el, paragraph)
            fmt["_raw_pPr"] = etree.tostring(pPr_el, encoding="unicode")
        else:
            # 1c: paragraph has no explicit <w:pPr> — create a detached element,
            # populate via style inheritance, and store only if non-empty.
            pPr_el = OxmlElement("w:pPr")
            _inherit_style_pPr(pPr_el, paragraph)
            if len(pPr_el):
                fmt["_raw_pPr"] = etree.tostring(pPr_el, encoding="unicode")
    except (AttributeError, TypeError):
        pass

    return fmt


def _iter_runs(paragraph: Paragraph):
    """Yield Run objects for all ``<w:r>`` elements in *paragraph*,
    including those nested inside wrapper elements such as ``<w:hyperlink>``,
    ``<w:ins>``, ``<w:del>``, ``<w:smartTag>``, ``<w:fldSimple>``,
    ``<w:sdt>``, and ``<w:customXml>``.

    python-docx ``paragraph.runs`` only returns direct ``<w:r>`` children,
    so we access the underlying lxml element to also reach runs wrapped in
    these container elements (used by TOC entries, cross-references, track
    changes, content controls, etc.).
    """
    _tag_r = qn("w:r")
    # Elements that may contain <w:r> children (directly or via sdtContent)
    _wrapper_tags = frozenset({
        qn("w:hyperlink"),
        qn("w:ins"),
        qn("w:del"),
        qn("w:smartTag"),
        qn("w:fldSimple"),
        qn("w:customXml"),
    })
    _tag_sdt = qn("w:sdt")
    _tag_sdt_content = qn("w:sdtContent")
    for child in paragraph._element:
        if child.tag == _tag_r:
            yield Run(child, paragraph)
        elif child.tag in _wrapper_tags:
            for r_el in child.findall(_tag_r):
                yield Run(r_el, paragraph)
        elif child.tag == _tag_sdt:
            sdt_content = child.find(_tag_sdt_content)
            if sdt_content is not None:
                for r_el in sdt_content.findall(_tag_r):
                    yield Run(r_el, paragraph)


def parse_paragraph_block(paragraph: Paragraph, block_id: str) -> dict:
    content = []
    for run in _iter_runs(paragraph):
        image_node = _parse_inline_image(run)
        if image_node is not None:
            content.append(image_node)
            continue
        item: dict = {"type": "Text", "text": run.text}
        overrides = _font_to_overrides(run.font, paragraph=paragraph)
        if overrides:
            item["overrides"] = overrides
        content.append(item)
    content = _merge_runs(content)

    default_run = _font_to_overrides(
        getattr(paragraph.style, "font", None), skip_theme_color=True
    )
    default_run.pop("_raw_rPr", None)

    para_fmt = _parse_paragraph_format(paragraph)

    block = {
        "id": block_id,
        "type": "Paragraph",
        "style": paragraph.style.style_id if paragraph.style else None,
        "content": content,
    }
    if para_fmt:
        block["paragraph_format"] = para_fmt
    if default_run:
        block["default_run"] = default_run

    return block
