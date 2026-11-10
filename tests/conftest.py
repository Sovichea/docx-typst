from __future__ import annotations

from pathlib import Path
import sys

import pytest
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
from PIL import Image


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))


@pytest.fixture()
def source_docx(tmp_path: Path) -> Path:
    image_path = tmp_path / "pixel.png"
    Image.new("RGBA", (24, 16), (40, 120, 200, 255)).save(image_path)

    doc = Document()
    section = doc.sections[0]
    section.page_width = Inches(8.5)
    section.page_height = Inches(11)
    section.left_margin = Inches(1)
    section.right_margin = Inches(1)
    section.top_margin = Inches(1)
    section.bottom_margin = Inches(1)

    header = section.header.paragraphs[0]
    header.text = "Confidential draft"
    header.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    footer = section.footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    field_run = footer.add_run()
    field = OxmlElement("w:fldSimple")
    field.set(qn("w:instr"), "PAGE")
    field_run._r.addnext(field)

    doc.add_heading("Quarterly report", level=1)
    paragraph = doc.add_paragraph()
    paragraph.add_run("Bold").bold = True
    paragraph.add_run(" and ")
    italic = paragraph.add_run("italic")
    italic.italic = True
    paragraph.add_run(" and ")
    colored = paragraph.add_run("colored underlined")
    colored.underline = True
    colored.font.color.rgb = RGBColor(0xC0, 0x20, 0x20)
    colored.font.size = Pt(13)
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

    paragraph = doc.add_paragraph("A hyperlink: ")
    rel_id = doc.part.relate_to(
        "https://example.com",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        is_external=True,
    )
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), rel_id)
    run = OxmlElement("w:r")
    text = OxmlElement("w:t")
    text.text = "example"
    run.append(text)
    hyperlink.append(run)
    paragraph._p.append(hyperlink)

    doc.add_paragraph("First bullet", style="List Bullet")
    doc.add_paragraph("Second bullet", style="List Bullet")
    doc.add_paragraph("First number", style="List Number")
    doc.add_paragraph("Second number", style="List Number")

    paragraph = doc.add_paragraph()
    paragraph.add_run("Inline picture: ")
    paragraph.add_run().add_picture(str(image_path), width=Pt(48))

    table = doc.add_table(rows=3, cols=2)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.cell(0, 0).text = "Name"
    table.cell(0, 1).text = "Value"
    table.cell(1, 0).text = "Alpha"
    table.cell(1, 1).text = "1"
    table.cell(2, 0).merge(table.cell(2, 1))
    table.cell(2, 0).text = "Merged across both columns"

    doc.add_page_break()
    doc.add_heading("Second page", level=2)
    doc.add_paragraph("After the page break.")

    properties = doc.core_properties
    properties.title = "Quarterly report"
    properties.author = "Smoke Test"
    path = tmp_path / "source.docx"
    doc.save(path)
    return path


@pytest.fixture()
def converted_project(tmp_path: Path, source_docx: Path):
    from docx_typst import docx_to_typst

    project = tmp_path / "project"
    report = docx_to_typst(source_docx, project)
    return source_docx, project, report
