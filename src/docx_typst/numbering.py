"""Independent Word numbering instances for independent Typst enums."""
from typing import Any

from docx.oxml import OxmlElement
from docx.oxml.ns import qn


def apply_numbering(writer: Any, paragraph: Any, block: Any) -> None:
    if block.list_kind != "number":
        return
    if not hasattr(writer, "_enum_ids"):
        writer._enum_ids = {}
        writer._enum_current = None
    root = writer.doc.part.numbering_part.element
    if not hasattr(writer, "_enum_abstract"):
        writer._enum_abstract = max((int(n.get(qn("w:abstractNumId")))
                                    for n in root.findall(qn("w:abstractNum"))), default=-1) + 1
        abstract = OxmlElement("w:abstractNum")
        abstract.set(qn("w:abstractNumId"), str(writer._enum_abstract))
        for level in range(9):
            node = OxmlElement("w:lvl")
            node.set(qn("w:ilvl"), str(level))
            for tag, value in (("start", "1"), ("numFmt", "decimal"),
                               ("lvlText", f"%{level + 1}."), ("lvlJc", "left")):
                child = OxmlElement("w:" + tag)
                child.set(qn("w:val"), value)
                node.append(child)
            abstract.append(node)
        root.insert(0, abstract)
    number = writer._enum_ids.get(block.number_id) if block.number_id is not None else writer._enum_current
    if number is None or (block.list_restart and block.number_id is None):
        instance = root.add_num(writer._enum_abstract)
        instance.add_lvlOverride(block.list_level).add_startOverride(1)
        number = instance.numId
        if block.number_id is not None:
            writer._enum_ids[block.number_id] = number
        writer._enum_current = number
    properties = paragraph._p.get_or_add_pPr().get_or_add_numPr()
    properties.get_or_add_ilvl().val = block.list_level
    properties.get_or_add_numId().val = number
