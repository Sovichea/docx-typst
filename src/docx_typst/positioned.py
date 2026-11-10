"""Bounded page canvases with editable Word text and vector artwork."""
from __future__ import annotations

from copy import deepcopy
from pathlib import Path
from typing import Any

from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.opc.part import Part
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.shared import Pt
from lxml import etree

from .model import CanvasBlock, ImageBlock, Paragraph, PositionedBlock
from .svg_compat import office_svg

ASVG = "http://schemas.microsoft.com/office/drawing/2016/SVG/main"
SOURCE_NS = "https://docx-typst.dev/source/v1"
WPS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"


def embed_svg(part: Any, inline: Any, path: Path) -> None:
    package = part.package
    svg = Part(package.next_partname("/word/media/vector%d.svg"),
               "image/svg+xml", office_svg(path.read_bytes()), package)
    rid = part.relate_to(svg, RT.IMAGE)
    blip = inline.xpath(".//a:blip")[0]
    extensions = OxmlElement("a:extLst")
    extension = OxmlElement("a:ext")
    extension.set("uri", "{96DAC541-7B7A-43D3-8B79-37D633B846F1}")
    svg_blip = etree.SubElement(extension, f"{{{ASVG}}}svgBlip", nsmap={"asvg": ASVG})
    svg_blip.set(qn("r:embed"), rid)
    extensions.append(extension)
    blip.append(extensions)


def _anchor(inline: Any, x: float, y: float, layer: int) -> Any:
    anchor = OxmlElement("wp:anchor")
    for key, value in {"distT": "0", "distB": "0", "distL": "0", "distR": "0",
                       "simplePos": "0", "relativeHeight": str(251658240 + layer * 1024), "behindDoc": "0",
                       "locked": "0", "layoutInCell": "1", "allowOverlap": "1"}.items():
        anchor.set(key, value)
    simple = OxmlElement("wp:simplePos")
    simple.set("x", "0")
    simple.set("y", "0")
    anchor.append(simple)
    for axis, offset in (("H", x), ("V", y)):
        position = OxmlElement("wp:position" + axis)
        position.set("relativeFrom", "page")
        value = OxmlElement("wp:posOffset")
        value.text = str(round(offset * 12700))
        position.append(value)
        anchor.append(position)
    anchor.append(deepcopy(inline.find(qn("wp:extent"))))
    effect = inline.find(qn("wp:effectExtent"))
    if effect is not None:
        anchor.append(deepcopy(effect))
    anchor.append(OxmlElement("wp:wrapNone"))
    for tag in ("wp:docPr", "wp:cNvGraphicFramePr", "a:graphic"):
        child = inline.find(qn(tag))
        if child is not None:
            anchor.append(deepcopy(child))
    return anchor


def write_canvas(writer: Any, container: Any, canvas: CanvasBlock) -> None:
    from .source_preservation import record_canvas

    paragraph = container.add_paragraph()
    fmt = paragraph.paragraph_format
    fmt.space_before = fmt.space_after = Pt(0)
    fmt.line_spacing = Pt(1)
    writer._pending_space = 0
    paragraph._p.set(f"{{{SOURCE_NS}}}canvas", canvas.identifier)
    bookmark = f"dxt_canvas_{len(writer.source_records)}"
    writer._bookmark(paragraph, bookmark, start=True)
    positions: list[dict[str, Any]] = []
    layer = 1
    for item in canvas.blocks:
        position = item if isinstance(item, PositionedBlock) else PositionedBlock(
            blocks=[item], width=canvas.width, height=canvas.height)
        images = [b for b in position.blocks if isinstance(b, ImageBlock)]
        texts = [b for b in position.blocks if isinstance(b, Paragraph)]
        for image in images:
            before = len(paragraph._p.xpath(".//wp:inline"))
            writer._image(paragraph, image.image)
            inlines = paragraph._p.xpath(".//wp:inline")
            if len(inlines) > before:
                inline = inlines[-1]
                inline.getparent().replace(inline, _anchor(inline, position.x, position.y, layer))
                layer += 1
        if not texts:
            continue
        identity = f"docx-typst-{canvas.identifier}-{len(positions)}"
        inline = OxmlElement("wp:inline")
        extent = OxmlElement("wp:extent")
        extent.set("cx", str(round(position.width * 12700)))
        extent.set("cy", str(round(position.height * 12700)))
        inline.append(extent)
        properties = OxmlElement("wp:docPr")
        properties.set("id", str(writer.doc.part.next_id))
        properties.set("name", identity)
        inline.append(properties)
        graphic = OxmlElement("a:graphic")
        data = OxmlElement("a:graphicData")
        data.set("uri", WPS)
        graphic.append(data)
        inline.append(graphic)
        shape = etree.SubElement(data, f"{{{WPS}}}wsp", nsmap={"wps": WPS})
        nonvisual = etree.SubElement(shape, f"{{{WPS}}}cNvSpPr")
        nonvisual.set("txBox", "1")
        geometry = etree.SubElement(shape, f"{{{WPS}}}spPr")
        transform = OxmlElement("a:xfrm")
        offset = OxmlElement("a:off")
        offset.set("x", "0"); offset.set("y", "0")
        dimensions = OxmlElement("a:ext")
        dimensions.set("cx", extent.get("cx")); dimensions.set("cy", extent.get("cy"))
        transform.extend([offset, dimensions])
        geometry.append(transform)
        preset = OxmlElement("a:prstGeom"); preset.set("prst", "rect")
        preset.append(OxmlElement("a:avLst")); geometry.append(preset)
        geometry.append(OxmlElement("a:noFill"))
        line = OxmlElement("a:ln"); line.append(OxmlElement("a:noFill")); geometry.append(line)
        textbox = etree.SubElement(shape, f"{{{WPS}}}txbx")
        content = OxmlElement("w:txbxContent")
        textbox.append(content)
        # Use a temporary python-docx container so existing style and run
        # handling remains the single implementation for editable text.
        from docx.blkcntnr import BlockItemContainer
        temporary = OxmlElement("w:tc")
        box_container = BlockItemContainer(temporary, paragraph._parent)
        for text in texts:
            text = deepcopy(text)
            if position.style:
                text.style = position.style
            text.align = position.align
            text.space_before = text.space_after = 0
            writer._paragraph(box_container, text)
        for child in list(temporary):
            content.append(child)
        body_properties = etree.SubElement(shape, f"{{{WPS}}}bodyPr")
        for key, value in {"lIns": "0", "tIns": "0", "rIns": "0", "bIns": "0",
                           "anchor": "t", "wrap": "square", "vertOverflow": "overflow"}.items():
            body_properties.set(key, value)
        body_properties.append(OxmlElement("a:noAutofit"))
        drawing = OxmlElement("w:drawing")
        drawing.append(_anchor(inline, position.x, position.y, layer))
        paragraph.add_run()._r.append(drawing)
        positions.append({"id": identity, "source": position.source,
                          "text": _box_text(content), "style": position.style})
        layer += 1
    writer._bookmark(paragraph, bookmark, start=False)
    record_canvas(writer, canvas, positions, bookmark)
    writer.counts["canvases"] = writer.counts.get("canvases", 0) + 1


def _box_text(content: Any) -> str:
    paragraphs = []
    for paragraph in content.findall(qn("w:p")):
        pieces = []
        for node in paragraph.iter():
            if node.tag == qn("w:t"):
                pieces.append(node.text or "")
            elif node.tag == qn("w:br"):
                pieces.append("\n")
        paragraphs.append("".join(pieces))
    return "\n".join(paragraphs)
