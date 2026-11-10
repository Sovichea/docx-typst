"""Normalize nested SVG images to vector groups supported by Microsoft Word."""
from __future__ import annotations

import base64
import re

from lxml import etree

SVG = "http://www.w3.org/2000/svg"
XLINK = "http://www.w3.org/1999/xlink"


def office_svg(data: bytes) -> bytes:
    parser = etree.XMLParser(resolve_entities=False, no_network=True)
    root = etree.fromstring(data, parser)
    serial = 0
    # Typst emits an SVG source image as a data URI. Word silently omits these
    # images, even though the surrounding SVG and its native paths render.
    while True:
        images = [node for node in root.iter(f"{{{SVG}}}image")
                  if (node.get(f"{{{XLINK}}}href") or node.get("href") or "")
                  .startswith("data:image/svg+xml;base64,")]
        if not images:
            break
        for image in images:
            serial += 1
            href = image.get(f"{{{XLINK}}}href") or image.get("href")
            nested = etree.fromstring(base64.b64decode(href.split(",", 1)[1]), parser)
            viewbox = [float(v) for v in re.split(r"[\s,]+", nested.get("viewBox", "").strip()) if v]
            if len(viewbox) != 4 or viewbox[2] <= 0 or viewbox[3] <= 0:
                raise ValueError("embedded SVG requires a positive viewBox")
            x, y, width, height = (float(image.get(key, default)) for key, default in
                                   (("x", "0"), ("y", "0"), ("width", "0"), ("height", "0")))
            sx, sy = width / viewbox[2], height / viewbox[3]
            aspect = image.get("preserveAspectRatio", "xMidYMid meet")
            if aspect != "none":
                scale = max(sx, sy) if "slice" in aspect else min(sx, sy)
                align = aspect.split()[0]
                x += (width - viewbox[2] * scale) * (0 if "xMin" in align else 1 if "xMax" in align else 0.5)
                y += (height - viewbox[3] * scale) * (0 if "YMin" in align else 1 if "YMax" in align else 0.5)
                sx = sy = scale
            # Rename resource IDs so separate source SVGs can reuse gradient
            # names without colliding with each other or the outer drawing.
            identifiers = {node.get("id"): f"embedded{serial}-{node.get('id')}"
                           for node in nested.iter() if node.get("id")}
            for node in nested.iter():
                for key, value in list(node.attrib.items()):
                    if key == "id":
                        node.set(key, identifiers[value])
                    else:
                        for old, new in identifiers.items():
                            value = value.replace(f"url(#{old})", f"url(#{new})")
                            if value == "#" + old:
                                value = "#" + new
                        node.set(key, value)
            outer = etree.Element(f"{{{SVG}}}g")
            for key, value in image.attrib.items():
                if key not in {"x", "y", "width", "height", "preserveAspectRatio", "href", f"{{{XLINK}}}href"}:
                    outer.set(key, value)
            group = etree.SubElement(outer, f"{{{SVG}}}g")
            group.set("transform", f"translate({x} {y}) scale({sx} {sy}) translate({-viewbox[0]} {-viewbox[1]})")
            for key, value in nested.attrib.items():
                if key not in {"viewBox", "width", "height", "x", "y"}:
                    group.set(key, value)
            for child in list(nested):
                group.append(child)
            image.getparent().replace(image, outer)
    return etree.tostring(root)
