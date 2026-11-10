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


def test_docx_static_renders_body_to_png(tmp_path: Path) -> None:
    project = _project(
        tmp_path,
        """#docx-static(width: 240pt)[
  #rect(width: 240pt, height: 80pt, fill: rgb("#207F77"))
]""",
    )
    output = tmp_path / "out.docx"
    report = typst_to_docx(project / "main.typ", output)
    assert report["counts"]["images"] == 1
    assets = list((project / "assets" / "static").glob("*.png"))
    assert len(assets) == 1
    assert assets[0].stat().st_size > 0
    assert len(Document(output).inline_shapes) == 1


def test_docx_panel_becomes_shaded_table(tmp_path: Path) -> None:
    project = _project(
        tmp_path,
        """#docx-panel(fill: rgb("#FFF9E7"), inset: 10pt)[
  Panel text with *bold* content.
]""",
    )
    output = tmp_path / "panel.docx"
    typst_to_docx(project / "main.typ", output)
    document = Document(output)
    assert len(document.tables) == 1
    table = document.tables[0]
    assert table.cell(0, 0).text == "Panel text with bold content."
    shading = table.cell(0, 0)._tc.tcPr.find(
        "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}shd"
    )
    assert shading is not None
    assert shading.get(
        "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}fill"
    ) == "FFF9E7"
