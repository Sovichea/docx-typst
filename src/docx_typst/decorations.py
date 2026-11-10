"""Explicit decorative rules represented by ordinary Word paragraph borders."""
from typing import Any

from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt

from .model import DecorativeRule


def write_rule(writer: Any, container: Any, rule: DecorativeRule) -> None:
    paragraph = container.add_paragraph()
    fmt = paragraph.paragraph_format
    fmt.space_before = Pt(rule.word_before if rule.word_before is not None else rule.before)
    fmt.space_after = Pt(rule.after)
    fmt.line_spacing = Pt(1)
    fmt.left_indent = Pt(rule.height / 2)
    fmt.right_indent = Pt(max(0, writer._content_width().pt - rule.width + rule.height / 2))
    fmt.keep_with_next = True
    borders = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    for key, value in {"val": "single", "sz": str(round(rule.height * 8)),
                       "space": "0", "color": rule.fill.lstrip("#")}.items():
        bottom.set(qn("w:" + key), value)
    borders.append(bottom)
    paragraph._p.get_or_add_pPr().append(borders)
    # Bookmarks survive Word saves; no drawing object or unsupported shape XML.
    marker = OxmlElement("w:bookmarkStart")
    writer._rule_index = getattr(writer, "_rule_index", 0) + 1
    marker.set(qn("w:id"), str(100000 + writer._rule_index))
    marker.set(qn("w:name"), f"dxt_rule_{round(rule.width * 20)}_{writer._rule_index}_{round(rule.before * 20)}")
    end = OxmlElement("w:bookmarkEnd")
    end.set(qn("w:id"), marker.get(qn("w:id")))
    paragraph._p.extend([marker, end])
    paragraph.add_run().font.size = Pt(1)


def read_rule(paragraph: Any) -> DecorativeRule | None:
    markers = paragraph.xpath('./w:bookmarkStart[starts-with(@w:name,"dxt_rule_")]')
    borders = paragraph.xpath('./w:pPr/w:pBdr/w:bottom')
    if not markers or not borders:
        return None
    border = borders[0]
    after = paragraph.xpath('./w:pPr/w:spacing/@w:after')
    parts = markers[0].get(qn("w:name")).removeprefix("dxt_rule_").split("_")
    before = paragraph.xpath('./w:pPr/w:spacing/@w:before')
    return DecorativeRule(
        before=int(parts[2]) / 20 if len(parts) > 2 else 0,
        word_before=int(before[0]) / 20 if before else 0,
        width=int(markers[0].get(qn("w:name")).removeprefix("dxt_rule_").split("_")[0]) / 20,
        height=int(border.get(qn("w:sz"))) / 8,
        fill="#" + border.get(qn("w:color")),
        after=int(after[0]) / 20 if after else 0,
    )
