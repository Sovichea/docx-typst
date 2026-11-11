from pathlib import Path
import shutil

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from PIL import Image
import typst

from docx_typst import docx_to_typst, typst_to_docx

ROOT = Path(__file__).resolve().parents[1]


def test_asymmetric_borders_and_cell_insets_survive_reconstruction(tmp_path: Path) -> None:
    shutil.copyfile(ROOT / "src/docx_typst/template.typ", tmp_path / "template.typ")
    source = tmp_path / "source.typ"
    source.write_text('''#import "template.typ": *
#show: docx-document.with(page: (paper: "a4", margin: 65pt))
#docx-table(columns: 1, rows: 1, stroke: none, inset: 0pt,
  borders: (top: none, bottom: none, left: 3pt + rgb("F7EE16"), right: none, insideH: none, insideV: none),
  cell-insets: ((0, 0, (top: 6pt, bottom: 6pt, left: 18pt, right: 18pt)),),
  cells: ([Text],))
''', encoding="utf-8")
    typst_to_docx(source, tmp_path / "first.docx")
    docx_to_typst(tmp_path / "first.docx", tmp_path / "restored", restore_source=False)
    restored = tmp_path / "restored/main.typ"
    assert "borders:" in restored.read_text(encoding="utf-8")
    typst.compile(str(restored))
    typst_to_docx(restored, tmp_path / "second.docx")
    first = Document(tmp_path / "first.docx").tables[0]
    second = Document(tmp_path / "second.docx").tables[0]
    for edge in ("top", "bottom", "left", "right", "insideH", "insideV"):
        a = first._tbl.tblPr.find(qn("w:tblBorders")).find(qn("w:" + edge))
        b = second._tbl.tblPr.find(qn("w:tblBorders")).find(qn("w:" + edge))
        assert dict(a.attrib) == dict(b.attrib)
    margins = lambda table: {n.tag: dict(n.attrib) for n in table.cell(0, 0)._tc.tcPr.find(qn("w:tcMar"))}
    assert margins(first) == margins(second)


def test_inline_image_retains_inherited_right_alignment(tmp_path: Path) -> None:
    Image.new("RGB", (40, 20), "red").save(tmp_path / "image.png")
    document = Document()
    document.styles["Normal"].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    document.add_paragraph().add_run().add_picture(str(tmp_path / "image.png"))
    document.save(tmp_path / "source.docx")
    docx_to_typst(tmp_path / "source.docx", tmp_path / "restored")
    source = (tmp_path / "restored/main.typ").read_text(encoding="utf-8")
    assert 'align: "right"' in source
    typst_to_docx(tmp_path / "restored/main.typ", tmp_path / "second.docx")
    assert next(p for p in Document(tmp_path / "second.docx").paragraphs if p._p.xpath(".//w:drawing")).alignment == WD_ALIGN_PARAGRAPH.RIGHT


def test_vertical_gap_and_multiline_caption_do_not_leak_code(tmp_path: Path) -> None:
    shutil.copyfile(ROOT / "src/docx_typst/template.typ", tmp_path / "template.typ")
    source = tmp_path / "source.typ"
    source.write_text('''#import "template.typ": *
#show: docx-document.with(page: (paper: "a4", margin: 65pt))
The hospital's records remain authoritative.
#v(0.2cm) #text(
  size: 9pt, style: "italic",
)[Exceptions are escalated; the platform cannot clear billing.]
''', encoding="utf-8")
    typst_to_docx(source, tmp_path / "source.docx")
    text = " ".join(Document(tmp_path / "source.docx").element.body.xpath(".//w:t/text()"))
    assert "Exceptions are escalated" in text
    assert "#text" not in text and "v(0.2cm)" not in text
