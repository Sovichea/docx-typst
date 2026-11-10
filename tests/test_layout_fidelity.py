from pathlib import Path
import shutil

from docx import Document
from docx.oxml.ns import qn
from PIL import Image

from docx_typst import typst_to_docx

ROOT = Path(__file__).resolve().parents[1]


def project(tmp_path: Path, body: str) -> Path:
    shutil.copyfile(ROOT / "src/docx_typst/template.typ", tmp_path / "template.typ")
    (tmp_path / "main.typ").write_text(
        '#import "template.typ": *\n'
        '#show: docx-document.with(page: (paper: "a4", margin: 65pt))\n' + body,
        encoding="utf-8")
    return tmp_path / "main.typ"


def test_native_table_inherits_stroke_and_aligns_its_outer_border(tmp_path: Path) -> None:
    source = project(tmp_path, '''
#set table(inset: 6pt, stroke: 0.5pt + rgb("DCE3DF"))
#table(columns: (100pt, 200pt), [Header], [Value], [A], [B])
''')
    typst_to_docx(source, tmp_path / "out.docx")
    table = Document(tmp_path / "out.docx").tables[0]
    properties = table._tbl.tblPr
    assert properties.find(qn("w:tblInd")).get(qn("w:w")) == "120"
    assert properties.find(qn("w:tblCellMar")).find(qn("w:left")).get(qn("w:w")) == "120"
    borders = properties.find(qn("w:tblBorders"))
    assert len(borders) == 6
    assert all(border.get(qn("w:color")) == "DCE3DF" for border in borders)
    assert all(border.get(qn("w:sz")) == "4" for border in borders)


def test_borderless_table_overrides_word_style_borders(tmp_path: Path) -> None:
    source = project(tmp_path, '''
#docx-table(columns: 1, style: "Table Grid", stroke: none, cells: ([No border],))
''')
    typst_to_docx(source, tmp_path / "out.docx")
    borders = Document(tmp_path / "out.docx").tables[0]._tbl.tblPr.find(qn("w:tblBorders"))
    assert len(borders) == 6
    assert all(border.get(qn("w:val")) == "nil" for border in borders)


def test_table_rules_remain_scoped_and_alignment_survives_wrappers(tmp_path: Path) -> None:
    source = project(tmp_path, '''
#set table(stroke: 0.5pt + rgb("DCE3DF"))
#box[
  #set table(stroke: 2pt + rgb("112233"))
  #table(columns: 1, [Red])
]
#align(center)[#table(columns: (100pt,), [Gray])]
''')
    typst_to_docx(source, tmp_path / "out.docx")
    tables = Document(tmp_path / "out.docx").tables
    red = tables[0]._tbl.tblPr.find(qn("w:tblBorders"))[0]
    gray = tables[1]._tbl.tblPr.find(qn("w:tblBorders"))[0]
    assert red.get(qn("w:color")) == "112233"
    assert red.get(qn("w:sz")) == "16"
    assert gray.get(qn("w:color")) == "DCE3DF"
    assert tables[1]._tbl.tblPr.find(qn("w:jc")).get(qn("w:val")) == "center"
    assert tables[1]._tbl.tblPr.find(qn("w:tblInd")) is None


def test_image_keeps_explicit_spacing_before_it(tmp_path: Path) -> None:
    Image.new("RGB", (20, 10), "white").save(tmp_path / "picture.png")
    source = project(tmp_path, '''
Before
#v(12pt)
#align(right)[#image("picture.png", width: 50pt)]
After
''')
    typst_to_docx(source, tmp_path / "out.docx")
    document = Document(tmp_path / "out.docx")
    paragraph = next(p for p in document.paragraphs if p._p.xpath(".//wp:inline"))
    assert paragraph.paragraph_format.space_before.pt == 12
    assert paragraph._p.pPr.find(qn("w:jc")).get(qn("w:val")) == "right"
