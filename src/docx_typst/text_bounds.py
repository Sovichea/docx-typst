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


@lru_cache(maxsize=128)
def full_line_height(font: str, size: float) -> float | None:
    source = (f"#set text(font: {json.dumps(font)}, size: {size}pt)\n"
              '#context metadata(measure(text(top-edge: "ascender", bottom-edge: "descender")[H]).height / 1pt)')
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
    for run in paragraph.runs:
        if run.text.strip():
            candidate_size = run.font.size or size
            if candidate_size and (size is None or candidate_size >= size):
                size = candidate_size
                font = run.font.name or font
    if not size or not font or line is None:
        return 0.0
    from .docx_reader import split_font_face
    font, _ = split_font_face(font)
    height = cap_height(font, size.pt)
    if rule == WD_LINE_SPACING.EXACTLY:
        line_height = line.pt
    elif rule == WD_LINE_SPACING.MULTIPLE:
        natural = full_line_height(font, size.pt)
        line_height = natural * line if natural is not None else 0.0
    else:
        return 0.0
    return max(0.0, line_height - height) if height is not None else 0.0


def native_cell_padding(cell: Any, inset: float | dict[str, float]) -> None:
    paragraphs = cell.paragraphs
    if not paragraphs:
        return
    values = {edge: inset.get(edge, 0) if isinstance(inset, dict) else inset
              for edge in ("top", "bottom", "left", "right")}
    top_excess = _line_excess(paragraphs[0]) / 2
    bottom_excess = _line_excess(paragraphs[-1]) / 2
    top = max(0.0, values["top"] - top_excess)
    bottom = max(0.0, values["bottom"] - bottom_excess)
    if isinstance(inset, dict):
        # When an authored edge is zero, compensate the remaining line-box
        # padding at the other edge rather than enlarging the whole row.
        bottom = max(0.0, bottom - max(0.0, top_excess - values["top"]))
        top = max(0.0, top - max(0.0, bottom_excess - values["bottom"]))
    if not isinstance(inset, dict) and top == inset and bottom == inset:
        return
    values.update(top=top, bottom=bottom)
    # Preserve the logical table inset in tblCellMar for source round trips.
    # Word's line box already supplies part of the visible vertical inset;
    # only the remainder belongs in the per-cell top/bottom margins.
    properties = cell._tc.get_or_add_tcPr()
    margins = properties.find(qn("w:tcMar"))
    if margins is None:
        margins = OxmlElement("w:tcMar")
        properties.insert_element_before(margins, "w:textDirection", "w:tcFitText", "w:vAlign", "w:hideMark")
    for edge, value in (values.items() if isinstance(inset, dict) else (("top", top), ("bottom", bottom))):
        for previous in margins.findall(qn("w:" + edge)):
            margins.remove(previous)
        element = OxmlElement("w:" + edge)
        element.set(qn("w:w"), str(round(value * 20)))
        element.set(qn("w:type"), "dxa")
        margins.append(element)


@lru_cache(maxsize=256)
def text_width(font: str, size: float, weight: str, text: str) -> float | None:
    source = (f"#set text(font: {json.dumps(font)}, size: {size}pt, weight: {json.dumps(weight)})\n"
              f"#context metadata(measure({json.dumps(text)}).width / 1pt)")
    try:
        return float(json.loads(typst.query(source.encode(), "metadata", field="value", one=True)))
    except Exception:
        return None
