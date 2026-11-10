from __future__ import annotations

import os
from pathlib import Path
import shutil
import subprocess
import sys

from docx import Document


ROOT = Path(__file__).resolve().parents[1]


def _env() -> dict[str, str]:
    environment = os.environ.copy()
    environment["PYTHONPATH"] = str(ROOT / "src")
    return environment


def _run(*args: str, cwd: Path | None = None) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        [sys.executable, "-m", "docx_typst.cli", *args],
        cwd=cwd or ROOT,
        env=_env(),
        check=True,
        capture_output=True,
        text=True,
        timeout=120,
    )


def test_init_creates_compilable_starter(tmp_path: Path) -> None:
    project = tmp_path / "starter"
    _run("init", str(project))
    assert (project / "docx-typst.typ").is_file()
    assert (project / "main.typ").is_file()
    assert (project / "assets" / "docx-typst-mark.svg").is_file()
    typst = shutil.which("typst")
    if typst:
        subprocess.run(
            [typst, "compile", "main.typ", "starter.pdf"],
            cwd=project,
            check=True,
            capture_output=True,
            text=True,
            timeout=120,
        )
        assert (project / "starter.pdf").is_file()
    output = tmp_path / "starter.docx"
    _run("to-docx", str(project / "main.typ"), "--output", str(output))
    assert output.is_file()
    assert Document(output).core_properties.title == "A new document"


def test_module_writes_only_the_helper(tmp_path: Path) -> None:
    output = tmp_path / "helper.typ"
    _run("module", "--output", str(output))
    text = output.read_text(encoding="utf-8")
    assert "#let docx-document(" in text
    assert "#let docx-para(" in text


def test_cli_converts_both_directions(tmp_path: Path, source_docx: Path) -> None:
    project = tmp_path / "project"
    _run("to-typst", str(source_docx), "--output", str(project))
    assert (project / "main.typ").is_file()
    output = tmp_path / "roundtrip.docx"
    _run("to-docx", str(project / "main.typ"), "--output", str(output))
    document = Document(output)
    assert "Quarterly report" in [paragraph.text for paragraph in document.paragraphs]


def test_init_refuses_to_overwrite(tmp_path: Path) -> None:
    project = tmp_path / "starter"
    _run("init", str(project))
    result = subprocess.run(
        [sys.executable, "-m", "docx_typst.cli", "init", str(project)],
        cwd=ROOT,
        env=_env(),
        capture_output=True,
        text=True,
    )
    assert result.returncode != 0
    assert "Refusing to overwrite" in result.stderr
    _run("init", str(project), "--force")
