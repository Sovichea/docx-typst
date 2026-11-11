"""Embed editable script fonts so Office does not silently substitute them."""
from pathlib import Path
import os
from uuid import uuid4

from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.packuri import PackURI
from docx.opc.part import Part
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn
from fontTools.ttLib import TTFont
from lxml import etree


def embed_script_fonts(document, families: set[str]) -> None:
    if not families:
        return
    directories = [Path(os.environ.get("WINDIR", "C:/Windows")) / "Fonts",
                   Path(os.environ.get("LOCALAPPDATA", "")) / "Microsoft/Windows/Fonts"]
    table = document.part.part_related_by(RT.FONT_TABLE)
    root = parse_xml(table.blob)
    for family in sorted(families):
        compact = family.replace(" ", "").lower()
        candidates = [p for folder in directories if folder.is_dir()
                      for p in folder.glob("*.ttf") if compact in p.stem.replace(" ", "").lower()]
        for path in candidates:
            try:
                with TTFont(path, lazy=True) as font:
                    if font["name"].getDebugName(1) != family:
                        continue
                    if font["name"].getDebugName(2) not in {"Regular", "Normal"}:
                        continue
                    if font["OS/2"].fsType & (0x0002 | 0x0200):
                        continue
                key = uuid4()
                blob = bytearray(path.read_bytes())
                mask = key.bytes[::-1]
                for i in range(32):
                    blob[i] ^= mask[i % 16]
                part = Part(PackURI(f"/word/fonts/{key.hex}.odttf"),
                            "application/vnd.openxmlformats-officedocument.obfuscatedFont",
                            bytes(blob), table.package)
                relationship = table.relate_to(part, RT.FONT)
                node = next((n for n in root if n.get(qn("w:name")) == family), None)
                if node is None:
                    node = OxmlElement("w:font")
                    node.set(qn("w:name"), family)
                    root.append(node)
                for existing in list(node.findall(qn("w:embedRegular"))):
                    node.remove(existing)
                embedded = OxmlElement("w:embedRegular")
                embedded.set(qn("r:id"), relationship)
                embedded.set(qn("w:fontKey"), "{" + str(key).upper() + "}")
                node.append(embedded)
                break
            except (OSError, KeyError, ValueError):
                continue
    table._blob = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
