from __future__ import annotations

from pathlib import Path

from docx import Document
from docx.oxml.ns import qn


def test_roundtrip_preserves_text_and_structure(converted_project, tmp_path: Path) -> None:
    from docx_typst import typst_to_docx

    _source, project, _report = converted_project
    output = tmp_path / "roundtrip.docx"
    report = typst_to_docx(project / "main.typ", output)
    assert output.is_file()
    assert report["status"] == "completed"

    document = Document(output)
    texts = [paragraph.text for paragraph in document.paragraphs]
    assert "Quarterly report" in texts
    assert "Bold and italic and colored underlined" in texts
    assert "A hyperlink: example" in texts
    assert "First bullet" in texts
    assert "Merged across both columns" in document.tables[0].cell(2, 0).text
    assert document.sections[0].header.paragraphs[0].text == "Confidential draft"
    assert document.sections[0].page_width.pt == 612.0
    assert len(document.inline_shapes) == 1


def test_roundtrip_preserves_inline_formatting(converted_project, tmp_path: Path) -> None:
    from docx_typst import typst_to_docx

    _source, project, _report = converted_project
    output = tmp_path / "roundtrip.docx"
    typst_to_docx(project / "main.typ", output)
    document = Document(output)
    assert document.inline_shapes
    paragraph = next(item for item in document.paragraphs if item.text.startswith("Bold and"))
    runs = {run.text: run for run in paragraph.runs}
    assert runs["Bold"].bold is True
    assert runs["italic"].italic is True
    assert runs["colored underlined"].underline is True
    assert str(runs["colored underlined"].font.color.rgb) == "C02020"


def test_roundtrip_preserves_hyperlinks_fields_and_merges(converted_project, tmp_path: Path) -> None:
    from docx_typst import typst_to_docx

    _source, project, _report = converted_project
    output = tmp_path / "roundtrip.docx"
    typst_to_docx(project / "main.typ", output)
    document = Document(output)
    hyperlink_paragraph = next(item for item in document.paragraphs if item.text.startswith("A hyperlink"))
    hyperlinks = hyperlink_paragraph._p.findall(qn("w:hyperlink"))
    assert hyperlinks
    run_properties = hyperlinks[0].find(qn("w:r")).find(qn("w:rPr"))
    assert run_properties.find(qn("w:u")) is not None
    color = run_properties.find(qn("w:color"))
    assert color is not None
    assert color.get(qn("w:val")) == "0563C1"
    assert document.sections[0].footer._element.findall(".//" + qn("w:fldSimple"))
    table = document.tables[0]
    assert table.style.name == "Table Grid"
    assert table.cell(2, 0)._tc is table.cell(2, 1)._tc


def test_edited_typst_source_changes_exported_text(converted_project, tmp_path: Path) -> None:
    from docx_typst import typst_to_docx

    _source, project, _report = converted_project
    main = project / "main.typ"
    main.write_text(
        main.read_text(encoding="utf-8").replace(
            "Quarterly report", "Revised report for export"
        ),
        encoding="utf-8",
    )
    output = tmp_path / "edited.docx"
    typst_to_docx(main, output)
    document = Document(output)
    assert "Revised report for export" in [paragraph.text for paragraph in document.paragraphs]


def test_native_static_typst_subset(tmp_path: Path) -> None:
    from docx_typst import typst_to_docx

    main = tmp_path / "main.typ"
    main.write_text(
        """
#set page(paper: "a4", margin: 2cm)
#set document(title: "Native document", author: "Tester")

= Native heading

A *bold* and _italic_ paragraph with `code`.

- First bullet
- Second bullet

+ First number
+ Second number

#table(columns: 2, [A], [B], [C], [D])
""".strip(),
        encoding="utf-8",
    )
    output = tmp_path / "native.docx"
    typst_to_docx(main, output)
    document = Document(output)
    properties = document.core_properties
    assert properties.title == "Native document"
    assert properties.author == "Tester"
    assert round(document.sections[0].page_width.pt, 1) == 595.3
    texts = [paragraph.text for paragraph in document.paragraphs]
    assert "Native heading" in texts
    assert "A bold and italic paragraph with code." in texts
    assert document.paragraphs[0].style.name == "Heading 1"
    assert document.tables[0].cell(1, 1).text == "D"
