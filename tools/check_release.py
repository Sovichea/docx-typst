"""Check release archives against the checkout without importing the package."""
import ast
from email.parser import BytesParser
from pathlib import Path
import re
import sys
import tarfile
import zipfile

ROOT = Path(__file__).resolve().parents[1]


def main():
    directory = Path(sys.argv[1] if len(sys.argv) > 1 else "dist")
    version = re.search(r'^version = "([^"]+)"',
                        (ROOT / "pyproject.toml").read_text(), re.M).group(1)
    module = ast.parse((ROOT / "src/docx_typst/__init__.py").read_text())
    runtime = next(ast.literal_eval(node.value) for node in module.body
                   if isinstance(node, ast.Assign)
                   and any(isinstance(t, ast.Name) and t.id == "__version__" for t in node.targets))
    assert version == runtime, "project/runtime versions disagree"
    expected = {f"docx_typst-{version}-py3-none-any.whl", f"docx_typst-{version}.tar.gz"}
    actual = {p.name for p in directory.iterdir() if p.is_file()}
    assert actual == expected, f"expected only {expected}, found {actual}"
    with zipfile.ZipFile(directory / f"docx_typst-{version}-py3-none-any.whl") as wheel:
        for source in (ROOT / "src/docx_typst").rglob("*"):
            if source.is_file() and source.suffix in {".py", ".typ", ".svg"}:
                name = source.relative_to(ROOT / "src").as_posix()
                assert wheel.read(name) == source.read_bytes(), f"missing/stale wheel resource: {name}"
        metadata = BytesParser().parsebytes(wheel.read(f"docx_typst-{version}.dist-info/METADATA"))
        assert metadata["Name"] == "docx-typst" and metadata["Version"] == version
        assert metadata["License-Expression"] == "MIT"
        assert wheel.read(f"docx_typst-{version}.dist-info/licenses/LICENSE") == (ROOT / "LICENSE").read_bytes()
        entries = wheel.read(f"docx_typst-{version}.dist-info/entry_points.txt").decode()
        for command in ("docx-typst", "docx-to-typst", "typst-to-docx", "docx-typst-validate"):
            assert command + " = " in entries, f"missing CLI: {command}"
    with tarfile.open(directory / f"docx_typst-{version}.tar.gz") as archive:
        required = [ROOT / name for name in ("README.md", "LICENSE", "CHANGELOG.md", "pyproject.toml")]
        for folder in ("docs", "skills", "tools", "tests", "src/docx_typst"):
            required.extend(p for p in (ROOT / folder).rglob("*")
                            if p.is_file() and p.suffix in {".md", ".yaml", ".py", ".typ", ".svg"})
        for source in required:
            name = f"docx_typst-{version}/" + source.relative_to(ROOT).as_posix()
            assert archive.extractfile(name).read() == source.read_bytes(), f"missing/stale source file: {name}"
        for name in archive.getnames():
            assert not {".git", ".work", ".venv", "__pycache__"}.intersection(Path(name).parts), name
    print(f"Release {version}: versions, CLI entry points, wheel resources, and sdist contents verified")


if __name__ == "__main__":
    main()
