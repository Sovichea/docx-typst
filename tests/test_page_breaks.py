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


def test_page_break_moves_to_next_heading_without_blank_page(tmp_path: Path) -> None:
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

    assert len(document.element.body.xpath(".//w:br[@w:type='page']")) == 1
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


def test_page_break_before_table_keeps_table_after_break(tmp_path: Path) -> None:
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
    assert len(document.element.body.xpath(".//w:br[@w:type='page']")) == 1


def test_section_break_has_separate_paragraph_after_justified_text(tmp_path: Path) -> None:
    project = _project(tmp_path, '''
#set par(justify: true)
A justified caption before the orientation change.
#set page(paper: "a4", flipped: true)
Landscape content.
''')
    output = tmp_path / "section-marker.docx"
    typst_to_docx(project / "main.typ", output)
    document = Document(output)
    caption = next(p for p in document.paragraphs if p.text.startswith("A justified"))
    assert caption._p.find("./" + qn("w:pPr") + "/" + qn("w:sectPr")) is None
    marker = caption._p.getnext()
    assert marker.tag == qn("w:p")
    assert marker.find("./" + qn("w:pPr") + "/" + qn("w:sectPr")) is not None
    assert not marker.xpath(".//w:t")
    assert marker.find("./" + qn("w:pPr") + "/" + qn("w:jc")).get(qn("w:val")) == "left"
