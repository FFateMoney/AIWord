"""Merge 层：将 AI 修改的 AST 合并回含 _raw_* 的完整 AST。

AI merge layer: merges the AI-modified view back into the full AST that
contains ``_raw_*`` XML fields, applying AI changes to the XML so that the
Renderer can preserve both round-trip fidelity and AI-requested edits.

合并规则 / Merge rules:
- AI **未修改**的字段：保留 original_ast 中的值（含 _raw_*）
  AI unchanged fields: keep original values including _raw_* XML.
- AI **修改了**的字段：更新结构化字段并同步写入 _raw_* 对应的 XML 元素
  AI changed fields: update structural fields AND sync _raw_* XML.
- XML 解析失败时：删除对应 _raw_* 让 Renderer 走结构化路径（降级）
  XML parse failure: drop _raw_* so Renderer falls back to structural fields.

Block matching is by ``id``; run matching is positional.
"""
import copy

from lxml import etree
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.oxml.parser import parse_xml

# Alignment: AST semantic value → OOXML <w:jc w:val="..."/>
_ALIGN_TO_JC: dict[str, str] = {
    "left": "left",
    "center": "center",
    "right": "right",
    "justify": "both",
}

# Paragraph-format semantic fields managed by this layer
_PPR_FIELDS = (
    "alignment",
    "indent_left",
    "indent_right",
    "indent_first_line",
    "space_before",
    "space_after",
)

# Run-format semantic fields managed by this layer
_RPR_FIELDS = ("font_ascii", "font_east_asia", "size", "bold", "italic", "color")


def merge_ai_edits(original_ast: dict, ai_ast: dict) -> dict:
    """将 AI 修改的 ai_ast 合并回含 _raw_* 的 original_ast。

    Merges the AI-modified *ai_ast* into the full *original_ast* that
    contains ``_raw_*`` fields, applying AI changes to the XML so that
    the Renderer produces a document that reflects both the original
    formatting fidelity and any AI-requested edits.

    Block matching is by the ``id`` key; run (content item) matching is
    positional within each paragraph's ``content`` list.
    """
    result = copy.deepcopy(original_ast)

    ai_body = ai_ast.get("document", {}).get("body", [])
    ai_by_id = {
        b["id"]: b
        for b in ai_body
        if isinstance(b, dict) and "id" in b
    }

    for orig_block in result.get("document", {}).get("body", []):
        if not isinstance(orig_block, dict):
            continue
        block_id = orig_block.get("id")
        if block_id not in ai_by_id:
            continue
        ai_block = ai_by_id[block_id]
        if orig_block.get("type") == "Paragraph":
            _merge_paragraph_block(orig_block, ai_block)

    return result


def _merge_paragraph_block(orig_block: dict, ai_block: dict) -> None:
    """合并 AI 对段落块的修改，原地更新 orig_block。

    Merges AI edits into a paragraph block in place, updating both
    structural fields and the corresponding ``_raw_*`` XML.
    """
    # --- Merge paragraph_format ---
    orig_fmt = orig_block.get("paragraph_format", {})
    ai_fmt = ai_block.get("paragraph_format", {})
    merged_fmt = _merge_paragraph_format(orig_fmt, ai_fmt)
    if merged_fmt:
        orig_block["paragraph_format"] = merged_fmt
    elif "paragraph_format" in orig_block:
        del orig_block["paragraph_format"]

    # --- Merge content items (runs matched by position) ---
    orig_content = orig_block.get("content", [])
    ai_content = ai_block.get("content", [])
    for i, orig_piece in enumerate(orig_content):
        if i >= len(ai_content):
            break
        ai_piece = ai_content[i]
        if orig_piece.get("type") != "Text" or ai_piece.get("type") != "Text":
            continue

        # Update text if changed
        ai_text = ai_piece.get("text")
        if ai_text is not None and orig_piece.get("text") != ai_text:
            orig_piece["text"] = ai_text

        # Merge run-level overrides
        orig_ov = orig_piece.get("overrides", {})
        ai_ov = ai_piece.get("overrides", {})
        merged_ov = _merge_run_overrides(orig_ov, ai_ov)
        if merged_ov:
            orig_piece["overrides"] = merged_ov
        elif "overrides" in orig_piece:
            del orig_piece["overrides"]


def _merge_paragraph_format(orig_fmt: dict, ai_fmt: dict) -> dict:
    """合并段落格式，将 AI 的修改同步到 _raw_pPr XML。

    Merges paragraph format dicts, syncing any AI changes to the
    ``_raw_pPr`` XML string so the Renderer produces correct output.
    Returns the merged dict (a modified deep-copy of *orig_fmt*).
    """
    result = copy.deepcopy(orig_fmt)

    # Detect which semantic fields were changed by the AI
    changed: dict = {}
    for key in _PPR_FIELDS:
        orig_val = orig_fmt.get(key)
        ai_val = ai_fmt.get(key)
        if orig_val != ai_val:
            changed[key] = ai_val
            if ai_val is None:
                result.pop(key, None)
            else:
                result[key] = ai_val

    if not changed:
        return result

    # Sync changes into _raw_pPr XML
    if "_raw_pPr" in result:
        updated_xml = _apply_pPr_changes(result["_raw_pPr"], changed)
        if updated_xml is None:
            # XML parse failure — drop _raw_pPr; Renderer will use structural fields
            del result["_raw_pPr"]
        else:
            result["_raw_pPr"] = updated_xml

    return result


def _merge_run_overrides(orig_ov: dict, ai_ov: dict) -> dict:
    """合并 run 级格式覆盖，将 AI 的修改同步到 _raw_rPr XML。

    Merges run-level override dicts, syncing any AI changes to the
    ``_raw_rPr`` XML string.  Returns the merged dict.
    """
    result = copy.deepcopy(orig_ov)

    changed: dict = {}
    for key in _RPR_FIELDS:
        orig_val = orig_ov.get(key)
        ai_val = ai_ov.get(key)
        if orig_val != ai_val:
            changed[key] = ai_val
            if ai_val is None:
                result.pop(key, None)
            else:
                result[key] = ai_val

    if not changed:
        return result

    # Sync changes into _raw_rPr XML
    if "_raw_rPr" in result:
        updated_xml = _apply_rPr_changes(result["_raw_rPr"], changed)
        if updated_xml is None:
            del result["_raw_rPr"]
        else:
            result["_raw_rPr"] = updated_xml

    return result


# ---------------------------------------------------------------------------
# XML manipulation helpers
# ---------------------------------------------------------------------------

def _apply_pPr_changes(raw_pPr: str, changes: dict) -> str | None:
    """将语义字段变更写入 _raw_pPr XML 字符串，返回更新后的字符串。

    Applies *changes* (a dict of semantic field → new value) to the
    ``_raw_pPr`` XML string.  Returns the updated XML string, or ``None``
    if the input XML cannot be parsed (caller should then drop ``_raw_pPr``).
    """
    try:
        pPr_el = parse_xml(raw_pPr)
    except Exception:
        return None

    # alignment → <w:jc w:val="..."/>
    if "alignment" in changes:
        val = changes["alignment"]
        jc = pPr_el.find(qn("w:jc"))
        if val is None:
            if jc is not None:
                pPr_el.remove(jc)
        else:
            jc_val = _ALIGN_TO_JC.get(val, val)
            if jc is None:
                jc = OxmlElement("w:jc")
                pPr_el.append(jc)
            jc.set(qn("w:val"), jc_val)

    # indentation → <w:ind> attributes
    ind_changes = {
        k: v for k, v in changes.items()
        if k in ("indent_left", "indent_right", "indent_first_line")
    }
    if ind_changes:
        ind = pPr_el.find(qn("w:ind"))
        if ind is None:
            ind = OxmlElement("w:ind")
            pPr_el.append(ind)
        _IND_ATTR = {
            "indent_left": qn("w:left"),
            "indent_right": qn("w:right"),
            "indent_first_line": qn("w:firstLine"),
        }
        for field, attr in _IND_ATTR.items():
            if field in ind_changes:
                v = ind_changes[field]
                if v is None:
                    ind.attrib.pop(attr, None)
                else:
                    ind.set(attr, str(int(v)))

    # spacing → <w:spacing> attributes
    spacing_changes = {
        k: v for k, v in changes.items()
        if k in ("space_before", "space_after")
    }
    if spacing_changes:
        spacing = pPr_el.find(qn("w:spacing"))
        if spacing is None:
            spacing = OxmlElement("w:spacing")
            pPr_el.append(spacing)
        _SPACING_ATTR = {
            "space_before": qn("w:before"),
            "space_after": qn("w:after"),
        }
        for field, attr in _SPACING_ATTR.items():
            if field in spacing_changes:
                v = spacing_changes[field]
                if v is None:
                    spacing.attrib.pop(attr, None)
                else:
                    spacing.set(attr, str(int(v)))

    return etree.tostring(pPr_el, encoding="unicode")


def _apply_rPr_changes(raw_rPr: str, changes: dict) -> str | None:
    """将语义字段变更写入 _raw_rPr XML 字符串，返回更新后的字符串。

    Applies *changes* (a dict of semantic field → new value) to the
    ``_raw_rPr`` XML string.  Returns the updated XML string, or ``None``
    if the input XML cannot be parsed.
    """
    try:
        rPr_el = parse_xml(raw_rPr)
    except Exception:
        return None

    # font_ascii / font_east_asia → <w:rFonts> attributes
    if "font_ascii" in changes or "font_east_asia" in changes:
        rFonts = rPr_el.find(qn("w:rFonts"))
        if rFonts is None:
            rFonts = OxmlElement("w:rFonts")
            rPr_el.insert(0, rFonts)
        if "font_ascii" in changes:
            val = changes["font_ascii"]
            if val:
                rFonts.set(qn("w:ascii"), val)
                rFonts.set(qn("w:hAnsi"), val)
            else:
                rFonts.attrib.pop(qn("w:ascii"), None)
                rFonts.attrib.pop(qn("w:hAnsi"), None)
        if "font_east_asia" in changes:
            val = changes["font_east_asia"]
            if val:
                rFonts.set(qn("w:eastAsia"), val)
            else:
                rFonts.attrib.pop(qn("w:eastAsia"), None)

    # size → <w:sz> and <w:szCs> (value is already in half-points)
    if "size" in changes:
        val = changes["size"]
        for w_tag in ("w:sz", "w:szCs"):
            el = rPr_el.find(qn(w_tag))
            if val is None:
                if el is not None:
                    rPr_el.remove(el)
            else:
                if el is None:
                    el = OxmlElement(w_tag)
                    rPr_el.append(el)
                el.set(qn("w:val"), str(int(val)))

    # bold → <w:b> (presence = True, absence = False)
    if "bold" in changes:
        el = rPr_el.find(qn("w:b"))
        if changes["bold"]:
            if el is None:
                rPr_el.append(OxmlElement("w:b"))
        else:
            if el is not None:
                rPr_el.remove(el)

    # italic → <w:i>
    if "italic" in changes:
        el = rPr_el.find(qn("w:i"))
        if changes["italic"]:
            if el is None:
                rPr_el.append(OxmlElement("w:i"))
        else:
            if el is not None:
                rPr_el.remove(el)

    # color → <w:color w:val="RRGGBB"/>
    if "color" in changes:
        val = changes["color"]
        el = rPr_el.find(qn("w:color"))
        if val is None:
            if el is not None:
                rPr_el.remove(el)
        else:
            if el is None:
                el = OxmlElement("w:color")
                rPr_el.append(el)
            el.set(qn("w:val"), val.lstrip("#"))

    return etree.tostring(rPr_el, encoding="unicode")
