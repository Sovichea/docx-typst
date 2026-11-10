from __future__ import annotations

from pathlib import Path
import shutil
import subprocess

from docx import Document
from docx.enum.section import WD_ORIENT, WD_SECTION
from docx.shared import Inches


def _typst_compile(project: Path) -> None:
    typst = shutil.which("typst")
    if not typst:
        return
    subprocess.run(
        [typst, "compile", "main.typ", "preview.pdf"],
        cwd=project,
        check=True,
        capture_output=True,
        text=True,
        timeout=120,
    )


def test_table_row_count_survives_multiple_merged_rows(tmp_path: Path) -> None:
    from docx_typst import docx_to_typst, typst_to_docx

    source = tmp_path / "merged.docx"
    document = Document()
    table = document.add_table(rows=4, cols=2)
    table.style = "Table Grid"
    for row in range(4):
        table.cell(row, 0).text = f"A{row}"
        table.cell(row, 1).text = f"B{row}"
    table.cell(2, 0).merge(table.cell(2, 1))
    table.cell(2, 0).text = "Merged two"
    table.cell(3, 0).merge(table.cell(3, 1))
    table.cell(3, 0).text = "Merged three"
    document.save(source)

    project = tmp_path / "project"
    docx_to_typst(source, project)
    main = (project / "main.typ").read_text(encoding="utf-8")
    assert "rows: 4," in main
    _typst_compile(project)

    output = tmp_path / "roundtrip.docx"
    typst_to_docx(project / "main.typ", output)
    result = Document(output)
    assert len(result.tables[0].rows) == 4
    assert result.tables[0].cell(2, 0)._tc is result.tables[0].cell(2, 1)._tc
    assert result.tables[0].cell(3, 0).text == "Merged three"


def test_section_breaks_and_landscape_page_setup(tmp_path: Path) -> None:
    from docx_typst import docx_to_typst, typst_to_docx

    source = tmp_path / "sections.docx"
    document = Document()
    document.add_paragraph("First section")
    section = document.add_section(WD_SECTION.NEW_PAGE)
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width = Inches(11)
    section.page_height = Inches(8.5)
    section.left_margin = Inches(0.5)
    section.header.paragraphs[0].text = "Second header"
    document.add_paragraph("Second section")
    document.add_section(WD_SECTION.NEW_PAGE)
    document.add_paragraph("Third section")
    document.save(source)

    project = tmp_path / "project"
    docx_to_typst(source, project)
    main = (project / "main.typ").read_text(encoding="utf-8")
    assert "#set page(" in main
    assert "width:" in main
    assert main.count("Second header") >= 2
    _typst_compile(project)

    output = tmp_path / "roundtrip.docx"
    typst_to_docx(project / "main.typ", output)
    result = Document(output)
    assert len(result.sections) == 3
    assert round(result.sections[1].page_width.pt) == 792
    assert round(result.sections[1].left_margin.pt, 1) == 36.0
    assert result.sections[1].header.paragraphs[0].text == "Second header"
    assert result.sections[2].header.paragraphs[0].text == "Second header"
