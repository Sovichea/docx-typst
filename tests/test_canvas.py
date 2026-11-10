from __future__ import annotations

import base64
from pathlib import Path
import shutil
from zipfile import ZipFile

from docx import Document
from docx.oxml.ns import qn
from lxml import etree
from PIL import Image, ImageChops
import typst

from docx_typst import docx_to_typst, typst_to_docx
from docx_typst.positioned import SOURCE_NS
from docx_typst.svg_compat import office_svg

ROOT = Path(__file__).resolve().parents[1]
ART = '''#docx-static(width: 120pt, height: 80pt, format: "svg")[
  // Keep this native drawing, including its original coordinates.
  #place(top + left, dx: 90pt, dy: 60pt)[
    #rect(width: 20pt, height: 10pt, fill: rgb("207F77"), stroke: none)
  ]
]'''


def project(tmp_path: Path, body: str) -> Path:
    root = tmp_path / "project"
    root.mkdir()
    shutil.copyfile(ROOT / "src/docx_typst/template.typ", root / "template.typ")
    (root / "main.typ").write_text(
        '#import "template.typ": *\n'
        '#let docx-styles = ("Cover": (font: "Arial", size: 18pt, after: 0pt))\n'
        '#show: docx-document.with(page: (width: 120pt, height: 80pt, margin: 0pt))\n'
        + body, encoding="utf-8")
    return root


def test_svg_export_retains_bounds_transparency_and_source(tmp_path: Path) -> None:
    root = project(tmp_path, ART)
    original = (root / "main.typ").read_bytes()
    output = tmp_path / "art.docx"
    report = typst_to_docx(root / "main.typ", output)
    assert not report["diagnostics"]
    assert (root / "main.typ").read_bytes() == original
    svg = next((root / "assets/static").glob("*.svg"))
    node = etree.fromstring(svg.read_bytes())
    assert node.get("viewBox") == "0 0 120 80"
    png = Image.open(svg.with_suffix(".png")).convert("RGBA")
    assert png.getpixel((0, 0))[3] == 0
    with ZipFile(output) as package:
        assert any(name.endswith(".svg") for name in package.namelist())
        assert any(name.endswith(".png") for name in package.namelist())
    restored = tmp_path / "restored"
    docx_to_typst(output, restored)
    assert ART in (restored / "main.typ").read_text(encoding="utf-8")
    typst.compile(str(restored / "main.typ"), output=str(tmp_path / "restored.pdf"))


def test_canvas_keeps_editable_text_coordinates_and_native_art(tmp_path: Path) -> None:
    canvas = '#docx-canvas(id: "cover", width: 120pt, height: 80pt)[\n' + ART + '''
  #docx-position(dx: 10pt, dy: 15pt, width: 100pt, height: 60pt, style: "Cover")[
    #text(fill: rgb("58686D"))[Original title \\ Second line]
  ]
]'''
    root = project(tmp_path, canvas)
    output = tmp_path / "canvas.docx"
    typst_to_docx(root / "main.typ", output)
    document = Document(output)
    anchors = document.element.body.xpath(".//wp:anchor")
    assert len(anchors) == 2
    assert int(anchors[0].get("relativeHeight")) >= 251658240
    assert int(anchors[1].get("relativeHeight")) > int(anchors[0].get("relativeHeight"))
    assert anchors[1].find(qn("wp:positionH")).find(qn("wp:posOffset")).text == "127000"
    assert anchors[1].find(qn("wp:positionV")).find(qn("wp:posOffset")).text == "190500"
    text = document.element.body.xpath(".//w:txbxContent")[0]
    assert text.findall(".//" + qn("w:br"))
    assert text.findall(".//" + qn("w:color"))[0].get(qn("w:val")) == "58686D"
    restored = tmp_path / "restored"
    docx_to_typst(output, restored)
    assert canvas in (restored / "main.typ").read_text(encoding="utf-8")
    typst.compile(str(root / "main.typ"), output=str(tmp_path / "before.png"), format="png")
    typst.compile(str(restored / "main.typ"), output=str(tmp_path / "after.png"), format="png")
    assert ImageChops.difference(Image.open(tmp_path / "before.png"),
                                Image.open(tmp_path / "after.png")).getbbox() is None

    # Simulate Word stripping the unknown paragraph attribute during Save.
    paragraph = document.paragraphs[0]._p
    del paragraph.attrib[f"{{{SOURCE_NS}}}canvas"]
    text.findall(".//" + qn("w:t"))[0].text = "Revised # title "
    document.save(tmp_path / "edited.docx")
    edited = tmp_path / "edited"
    docx_to_typst(tmp_path / "edited.docx", edited)
    result = (edited / "main.typ").read_text(encoding="utf-8")
    assert '#docx-canvas(id: "cover"' in result
    assert ART in result
    assert r"Revised \# title" in result
    assert "Original title" not in result
    typst.compile(str(edited / "main.typ"), output=str(tmp_path / "edited.pdf"))

    # Even deleting every editable text box must retain the artwork source.
    shape_run = anchors[1].getparent().getparent()
    shape_run.getparent().remove(shape_run)
    document.save(tmp_path / "deleted.docx")
    deleted = tmp_path / "deleted"
    docx_to_typst(tmp_path / "deleted.docx", deleted)
    result = (deleted / "main.typ").read_text(encoding="utf-8")
    assert ART in result
    assert '#docx-canvas(id: "cover"' in result
    assert "#docx-position(" not in result


def test_static_restores_imports_and_nested_local_dependencies(tmp_path: Path) -> None:
    root = project(tmp_path, '#docx-static(width: 120pt, format: "svg")[#art]')
    (root / "art").mkdir()
    (root / "art/picture.svg").write_text(
        '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 10 10">'
        '<circle cx="5" cy="5" r="4" fill="red"/></svg>', encoding="utf-8")
    (root / "art/drawing.typ").write_text(
        '#let art = image("picture.svg", width: 100pt)', encoding="utf-8")
    main = root / "main.typ"
    main.write_text('#import "art/drawing.typ": art\n' + main.read_text(encoding="utf-8"), encoding="utf-8")
    typst_to_docx(main, tmp_path / "drawing.docx")
    restored = tmp_path / "restored"
    docx_to_typst(tmp_path / "drawing.docx", restored)
    assert (restored / "art/picture.svg").read_bytes() == (root / "art/picture.svg").read_bytes()
    assert (restored / "art/drawing.typ").read_bytes() == (root / "art/drawing.typ").read_bytes()
    typst.compile(str(restored / "main.typ"), output=str(tmp_path / "drawing.pdf"))


def test_office_svg_inlines_nested_vectors_and_renames_resources() -> None:
    nested = b'''<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 10 10">
<defs><linearGradient id="y"/></defs><circle cx="5" cy="5" r="4" fill="url(#y)"/></svg>'''
    encoded = base64.b64encode(nested).decode()
    image = f'<image width="20" height="30" preserveAspectRatio="none" href="data:image/svg+xml;base64,{encoded}"/>'
    svg = f'<svg xmlns="http://www.w3.org/2000/svg">{image}{image}</svg>'
    result = etree.fromstring(office_svg(svg.encode()))
    assert not result.findall(".//{http://www.w3.org/2000/svg}image")
    circles = result.findall(".//{http://www.w3.org/2000/svg}circle")
    assert len(circles) == 2
    assert circles[0].get("fill") != circles[1].get("fill")
    groups = result.findall(".//{http://www.w3.org/2000/svg}g")
    assert any("scale(2.0 3.0)" in node.get("transform", "") for node in groups)


def test_svg_keeps_native_intrinsic_diagram_size_without_clipping(tmp_path: Path) -> None:
    root = project(tmp_path, '''#docx-static(width: 120pt, format: "svg")[
  #block(width: 180pt, height: 60pt)[
    #rect(width: 180pt, height: 60pt, fill: rgb("207F77"), stroke: none)
  ]
]''')
    output = tmp_path / "wide.docx"
    report = typst_to_docx(root / "main.typ", output)
    assert not report["diagnostics"]
    svg = next((root / "assets/static").glob("*.svg"))
    assert etree.fromstring(svg.read_bytes()).get("viewBox") == "0 0 180 60"
    # docx-static leaves its body unchanged in Typst, so Word must use the same
    # intrinsic frame rather than shrinking it to the export width hint.
    shape = Document(output).inline_shapes[0]
    assert shape.width.pt == 180
    assert shape.height.pt == 60
