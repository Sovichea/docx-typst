from __future__ import annotations

import json
from pathlib import Path
import shutil
import subprocess

from docx import Document


def test_forward_project_contains_editable_constructs(converted_project) -> None:
    _source, project, report = converted_project
    main = (project / "main.typ").read_text(encoding="utf-8")
    assert '#import "template.typ": *' in main
    assert '#let docx-styles = (' in main
    assert '#docx-heading(' in main
    assert "*Bold*" in main
    assert "_italic_" in main
    assert "#link(\"https://example.com\")" in main
    assert '#docx-item(styles: docx-styles, kind: "bullet"' in main
    assert '#docx-item(styles: docx-styles, kind: "number"' in main
    assert "#docx-table(" in main
    assert "#docx-inline-image(" in main
    assert "#pagebreak()" in main
    assert report["status"] == "completed_with_warnings"
    assert report["counts"]["headings"] == 2
    assert report["counts"]["tables"] == 1
    assert report["counts"]["images"] == 1


def test_forward_writes_manifest_and_assets(converted_project) -> None:
    _source, project, _report = converted_project
    manifest = json.loads((project / "asset_manifest.json").read_text(encoding="utf-8"))
    assert manifest["source_file"] == "source.docx"
    assert manifest["page"]["width_pt"] == 612.0
    assert manifest["assets"]
    for asset in manifest["assets"]:
        assert (project / asset["output_path"]).is_file()
    assert "Heading 1" in manifest["styles"]
    report = json.loads((project / "conversion_report.json").read_text(encoding="utf-8"))
    assert any(item["issue"] == "table_borders" for item in report["diagnostics"])


def test_forward_output_compiles_with_typst(converted_project) -> None:
    typst = shutil.which("typst")
    if not typst:
        return
    _source, project, _report = converted_project
    subprocess.run(
        [typst, "compile", "main.typ", "preview.pdf"],
        cwd=project,
        check=True,
        capture_output=True,
        text=True,
        timeout=120,
    )
    assert (project / "preview.pdf").is_file()


def test_forward_preserves_source_text(converted_project) -> None:
    source, _project, _report = converted_project
    original = Document(source).element.body
    text = "".join(original.itertext())
    assert "Quarterly report" in text
    assert "Merged across both columns" in text


def test_no_media_project_stays_compilable(tmp_path: Path, source_docx: Path) -> None:
    from docx_typst import docx_to_typst

    project = tmp_path / "no-media"
    docx_to_typst(source_docx, project, extract_media=False)
    main = (project / "main.typ").read_text(encoding="utf-8")
    assert "#docx-unsupported-image(" in main
    typst = shutil.which("typst")
    if typst:
        subprocess.run(
            [typst, "compile", "main.typ", "preview.pdf"],
            cwd=project,
            check=True,
            capture_output=True,
            text=True,
            timeout=120,
        )
