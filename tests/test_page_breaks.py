from __future__ import annotations

from pathlib import Path
import shutil

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

from docx_typst import typst_to_docx


ROOT = Path(__file__).resolve().parents[1]


def _project(tmp_path: Path, body: str) -> Path:
    project = tmp_path / "project"
    project.mkdir()
    shutil.copyfile(ROOT / "src/docx_typst/template.typ", project / "template.typ")
    (project / "main.typ").write_text(
        "\n".join(
            [
                '#import "template.typ": *',
                '#show: docx-document.with(page: (paper: "a4", margin: 2cm))',
                body,
                "",
            ]
        ),
        encoding="utf-8",
    )
    return project


def test_typst_page_break_is_ignored_before_heading(tmp_path: Path) -> None:
    project = _project(
        tmp_path,
        """First page text.

#pagebreak()
= Second Page Heading

Body text.
""",
    )
    output = tmp_path / "breaks.docx"
    typst_to_docx(project / "main.typ", output)
    document = Document(output)

    assert not document.element.body.xpath(".//w:br[@w:type='page']")
    heading = next(p for p in document.paragraphs if p.text == "Second Page Heading")
    assert heading.paragraph_format.page_break_before is not True
    assert heading.alignment == WD_ALIGN_PARAGRAPH.LEFT


def test_page_settings_inherit_margins_across_sections(tmp_path: Path) -> None:
    project = _project(
        tmp_path,
        """Portrait page.

#pagebreak()
#set page(paper: "a4", flipped: true)
Landscape page.

#pagebreak()
#set page(paper: "a4")
Portrait page again.
""",
    )
    output = tmp_path / "margins.docx"
    typst_to_docx(project / "main.typ", output)
    document = Document(output)

    assert len(document.sections) == 3
    landscape = document.sections[1]
    assert landscape.page_width > landscape.page_height
    first = document.sections[0]
    for section in document.sections:
        assert abs(section.left_margin.pt - first.left_margin.pt) < 0.1
        assert abs(section.right_margin.pt - first.right_margin.pt) < 0.1
        assert abs(section.top_margin.pt - first.top_margin.pt) < 0.1
        assert abs(section.bottom_margin.pt - first.bottom_margin.pt) < 0.1


def test_typst_page_break_is_ignored_before_table(tmp_path: Path) -> None:
    project = _project(
        tmp_path,
        """First page text.

#pagebreak()
#table(columns: 2, [A], [B])
""",
    )
    output = tmp_path / "break-table.docx"
    typst_to_docx(project / "main.typ", output)
    document = Document(output)
    assert len(document.tables) == 1
    assert not document.element.body.xpath(".//w:br[@w:type='page']")
