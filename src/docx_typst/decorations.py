"""Native Word vector rules, kept editable and round-trippable."""
from typing import Any

from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt
from lxml import etree

from .model import DecorativeRule
from .positioned import WPS


def write_rule(writer: Any, container: Any, rule: DecorativeRule) -> None:
    paragraph = container.add_paragraph()
    fmt = paragraph.paragraph_format
    fmt.space_before = Pt(0)
    fmt.space_after = Pt(rule.after)
    fmt.line_spacing = Pt(rule.height)
    fmt.keep_with_next = True
    inline = OxmlElement("wp:inline")
    extent = OxmlElement("wp:extent")
    extent.set("cx", str(round(rule.width * 12700)))
    extent.set("cy", str(round(rule.height * 12700)))
    properties = OxmlElement("wp:docPr")
    properties.set("id", str(writer.doc.part.next_id))
    properties.set("name", "dxt-heading-rule")
    properties.set("descr", str(rule.after))
    graphic = OxmlElement("a:graphic")
    data = OxmlElement("a:graphicData")
    data.set("uri", WPS)
    shape = etree.SubElement(data, f"{{{WPS}}}wsp", nsmap={"wps": WPS})
    etree.SubElement(shape, f"{{{WPS}}}cNvSpPr")
    geometry = etree.SubElement(shape, f"{{{WPS}}}spPr")
    preset = OxmlElement("a:prstGeom")
    preset.set("prst", "rect")
    preset.append(OxmlElement("a:avLst"))
    geometry.append(preset)
    fill = OxmlElement("a:solidFill")
    color = OxmlElement("a:srgbClr")
    color.set("val", rule.fill.lstrip("#"))
    fill.append(color)
    geometry.append(fill)
    line = OxmlElement("a:ln")
    line.append(OxmlElement("a:noFill"))
    geometry.append(line)
    etree.SubElement(shape, f"{{{WPS}}}bodyPr")
    graphic.append(data)
    inline.extend([extent, properties, graphic])
    drawing = OxmlElement("w:drawing")
    drawing.append(inline)
    paragraph.add_run()._r.append(drawing)


def read_rule(paragraph: Any) -> DecorativeRule | None:
    nodes = paragraph.xpath('.//wp:docPr[@name="dxt-heading-rule"]')
    if not nodes:
        return None
    inline = nodes[0].getparent()
    extent = inline.find(qn("wp:extent"))
    color = inline.xpath(".//a:solidFill/a:srgbClr")[0].get("val")
    return DecorativeRule(width=int(extent.get("cx")) / 12700,
                          height=int(extent.get("cy")) / 12700,
                          fill="#" + color, after=float(nodes[0].get("descr", "10")))
