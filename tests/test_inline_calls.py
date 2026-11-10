from __future__ import annotations

from pathlib import Path
import shutil

from docx import Document

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


def test_inline_calls_at_block_start_keep_their_formatting(tmp_path: Path) -> None:
    project = _project(
        tmp_path,
        """#link("https://example.com")[Example site]

#text(font: "DM Sans", size: 11pt)[Styled text]
""",
    )
    output = tmp_path / "inline.docx"
    report = typst_to_docx(project / "main.typ", output)
    document = Document(output)

    link = document.paragraphs[0]
    assert link.text == "Example site"
    assert link._p.xpath(".//w:hyperlink")
    rels = [r for r in document.part.rels.values() if "hyperlink" in r.reltype]
    assert [r.target_ref for r in rels] == ["https://example.com"]
    assert document.paragraphs[1].text == "Styled text"
    assert not any(item["source_type"] == "link" for item in report["diagnostics"])
