"""Translate Typst's tight first/last text edges to Word paragraph line boxes."""
from functools import lru_cache
import json
from typing import Any

from docx.enum.text import WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import typst


@lru_cache(maxsize=128)
def cap_height(font: str, size: float) -> float | None:
    source = (f"#set text(font: {json.dumps(font)}, size: {size}pt)\n"
              "#context metadata(measure([H]).height / 1pt)")
    try:
        return float(json.loads(typst.query(source.encode(), "metadata", field="value", one=True)))
    except Exception:
        return None


def _line_excess(paragraph: Any) -> float:
    size = font = line = rule = None
    fmt = paragraph.paragraph_format
    line, rule = fmt.line_spacing, fmt.line_spacing_rule
    style = paragraph.style
    while style is not None:
        size = size or style.font.size
        font = font or style.font.name
        if line is None:
            line = style.paragraph_format.line_spacing
            rule = style.paragraph_format.line_spacing_rule
        style = style.base_style
    if not size or not font or line is None or rule != WD_LINE_SPACING.EXACTLY:
        return 0.0
    height = cap_height(font, size.pt)
    return max(0.0, line.pt - height) if height is not None else 0.0


def native_cell_padding(cell: Any, inset: float) -> None:
    paragraphs = cell.paragraphs
    if not paragraphs:
        return
    top = max(0.0, inset - _line_excess(paragraphs[0]) / 2)
    bottom = max(0.0, inset - _line_excess(paragraphs[-1]) / 2)
    if top == inset and bottom == inset:
        return
    # Preserve the logical table inset in tblCellMar for source round trips.
    # Word's line box already supplies part of the visible vertical inset;
    # only the remainder belongs in the per-cell top/bottom margins.
    properties = cell._tc.get_or_add_tcPr()
    margins = properties.find(qn("w:tcMar"))
    if margins is None:
        margins = OxmlElement("w:tcMar")
        properties.insert_element_before(margins, "w:textDirection", "w:tcFitText", "w:vAlign", "w:hideMark")
    for edge, value in (("top", top), ("bottom", bottom)):
        for previous in margins.findall(qn("w:" + edge)):
            margins.remove(previous)
        element = OxmlElement("w:" + edge)
        element.set(qn("w:w"), str(round(value * 20)))
        element.set(qn("w:type"), "dxa")
        margins.append(element)
