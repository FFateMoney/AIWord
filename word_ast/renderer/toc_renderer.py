from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from .paragraph_renderer import render_paragraph


def render_toc(doc, block: dict, styles: dict | None = None):
    """Render a ``TOC`` block as a native Word TOC field wrapped in an SDT.

    The generated structure uses ``<w:sdt>`` with a ``docPartGallery`` of
    *Table of Contents* so that Word recognises it as a real TOC.  The field
    instruction (e.g. ``TOC \\o "1-3" \\h \\z \\u``) is written inside a
    complex field (``fldChar begin`` / ``instrText`` / ``fldChar separate`` /
    ``fldChar end``) and marked *dirty* so that Word refreshes the entries
    when the document is first opened.
    """
    body_el = doc.element.body

    # --- SDT wrapper ---
    sdt = OxmlElement("w:sdt")

    sdtPr = OxmlElement("w:sdtPr")
    docPartObj = OxmlElement("w:docPartObj")
    gallery = OxmlElement("w:docPartGallery")
    gallery.set(qn("w:val"), "Table of Contents")
    docPartObj.append(gallery)
    unique = OxmlElement("w:docPartUnique")
    docPartObj.append(unique)
    sdtPr.append(docPartObj)
    sdt.append(sdtPr)

    sdtContent = OxmlElement("w:sdtContent")

    # --- Optional title paragraph ---
    title = block.get("title")
    if title:
        render_paragraph(doc, title, styles)
        # Move the paragraph that was just appended to the document body
        # into sdtContent.  python-docx inserts before <w:sectPr>.
        sectPr = body_el.find(qn("w:sectPr"))
        if sectPr is not None:
            new_p = sectPr.getprevious()
        else:
            new_p = body_el[-1]
        if new_p is not None and new_p.tag == qn("w:p"):
            body_el.remove(new_p)
            sdtContent.append(new_p)

    # --- TOC field ---
    instruction = block.get("instruction", 'TOC \\o "1-3" \\h \\z \\u')

    # Paragraph: field-begin + instrText + field-separate
    p_field = OxmlElement("w:p")

    r_begin = OxmlElement("w:r")
    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    fld_begin.set(qn("w:dirty"), "true")
    r_begin.append(fld_begin)
    p_field.append(r_begin)

    r_instr = OxmlElement("w:r")
    instrText = OxmlElement("w:instrText")
    instrText.set(qn("xml:space"), "preserve")
    instrText.text = f" {instruction} "
    r_instr.append(instrText)
    p_field.append(r_instr)

    r_sep = OxmlElement("w:r")
    fld_sep = OxmlElement("w:fldChar")
    fld_sep.set(qn("w:fldCharType"), "separate")
    r_sep.append(fld_sep)
    p_field.append(r_sep)

    sdtContent.append(p_field)

    # Paragraph: field-end
    p_end = OxmlElement("w:p")
    r_end = OxmlElement("w:r")
    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    r_end.append(fld_end)
    p_end.append(r_end)
    sdtContent.append(p_end)

    sdt.append(sdtContent)

    # Insert SDT into document body (before sectPr)
    sectPr = body_el.find(qn("w:sectPr"))
    if sectPr is not None:
        sectPr.addprevious(sdt)
    else:
        body_el.append(sdt)
