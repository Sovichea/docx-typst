from __future__ import annotations

import json
from pathlib import Path


def test_validate_matches_generated_project(converted_project, tmp_path: Path) -> None:
    from docx_typst.validate import validate

    source, project, _report = converted_project
    report = validate(source, project, tmp_path / "validation.json")
    assert report["status"] == "match"
    assert report["text_similarity"] == 1.0
    assert report["differences"] == []
    saved = json.loads((tmp_path / "validation.json").read_text(encoding="utf-8"))
    assert saved["source"]["counts"]["tables"] == 1


def test_validate_detects_replaced_content(converted_project, tmp_path: Path) -> None:
    from docx_typst.validate import validate

    source, project, _report = converted_project
    (project / "main.typ").write_text(
        '#import "template.typ": *\n\n#set page(paper: "a4")\n\nA completely different document.\n',
        encoding="utf-8",
    )
    report = validate(source, project, tmp_path / "validation.json")
    assert report["status"] == "differences"
    assert report["text_similarity"] < 0.98
    assert any(item["kind"] == "text" for item in report["differences"])
