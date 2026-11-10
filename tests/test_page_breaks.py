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
    marker = heading._p.getprevious()
    assert marker.xpath("./w:r/w:br[@w:type='page']")
    assert not marker.xpath(".//w:t")
    assert marker.find("./" + qn("w:pPr") + "/" + qn("w:spacing")).get(qn("w:line")) == "240"


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


def test_inline_column_break_occupies_separate_paragraph(tmp_path: Path) -> None:
    project = _project(tmp_path, '#docx-para[Before#colbreak()After]')
    output = tmp_path / "column-break.docx"
    typst_to_docx(project / "main.typ", output)
    paragraphs = Document(output).paragraphs
    before = next(p for p in paragraphs if p.text == "Before")
    marker = before._p.getnext()
    assert marker.xpath("./w:r/w:br[@w:type='column']")
    assert not marker.xpath(".//w:t")
    assert marker.getnext().xpath(".//w:t/text()") == ["After"]


def test_page_start_heading_metrics_leave_multiline_spacing_unchanged(tmp_path: Path) -> None:
    project = _project(tmp_path, '''
#let docx-styles = ("Heading 1": (
  font: "Arial", size: 17pt, line-spacing-pt: 25.5pt,
  word-page-before: 0pt, word-page-line-spacing-pt: 16.3pt))
= Short heading
Body.
#pagebreak()
= This deliberately long heading extends well beyond the available page width and must retain the original multiline baseline spacing in the Word export
''')
    source = project / "main.typ"
    content = source.read_text(encoding="utf-8")
    start = content.index("#let docx-styles")
    end = content.index("= Short heading", start)
    styles = content[start:end]
    content = content[:start] + content[end:]
    content = content.replace("#show: docx-document", styles + "#show: docx-document", 1)
    source.write_text(content, encoding="utf-8")
    output = tmp_path / "heading-metrics.docx"
    typst_to_docx(project / "main.typ", output)
    headings = [p for p in Document(output).paragraphs if p.style.name == "Heading 1"]
    assert abs(headings[0].paragraph_format.line_spacing.pt - 16.3) < 0.01
    assert headings[0].paragraph_format.space_before.pt == 0
    assert headings[1].paragraph_format.line_spacing is None
    assert headings[1].style.paragraph_format.line_spacing.pt == 25.5
