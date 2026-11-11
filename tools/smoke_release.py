"""Exercise a wheel installation from outside the source tree (no test extras)."""
from importlib import metadata
from pathlib import Path
import os
import shutil
import subprocess
import sys
from tempfile import TemporaryDirectory

import docx_typst
from docx import Document
import typst


def main():
    root = Path(__file__).resolve().parents[1]
    assert not Path(docx_typst.__file__).resolve().is_relative_to(root / "src"), "test an installed wheel, not the checkout"
    assert metadata.version("docx-typst") == docx_typst.__version__
    environment = os.environ.copy()
    environment.pop("PYTHONPATH", None)
    with TemporaryDirectory(prefix="docx-typst-installed-") as temporary:
        directory = Path(temporary)

        def run(command, *args):
            executable = shutil.which(command, path=str(Path(sys.executable).parent))
            assert executable, f"missing installed entry point: {command}"
            subprocess.run([executable, *map(str, args)], check=True, cwd=directory,
                           env=environment, capture_output=True, text=True, timeout=120)

        for command in ("docx-typst", "docx-to-typst", "typst-to-docx", "docx-typst-validate"):
            run(command, "--help")
        project = directory / "project"
        run("docx-typst", "init", project)
        assert (project / "assets/docx-typst-mark.svg").is_file()
        typst.compile(str(project / "main.typ"))
        run("docx-typst", "module", "--output", directory / "helper.typ")
        source = project / "main.typ"
        source.write_text('''#import "docx-typst.typ": *
#let prices = (100, 200)
#let total = prices.sum()
#docx-eval("total", [*Total #total*])
#docx-eval("table", table(columns: (80pt, 1fr), [Subtotal], [#total]))
''', encoding="utf-8")
        authored = source.read_bytes()
        output = directory / "converted.docx"
        run("typst-to-docx", source, "--output", output, "--strict")
        doc = Document(output)
        assert doc.paragraphs[0].text == "Total 300"
        assert doc.tables[0].cell(0, 1).text == "300"
        assert source.read_bytes() == authored
        restored = directory / "restored"
        run("docx-to-typst", output, "--output", restored)
        assert (restored / "main.typ").read_bytes() == authored
        assert (restored / "docx-typst.typ").read_bytes() == (project / "docx-typst.typ").read_bytes()
        typst.compile(str(restored / "main.typ"))
        # Also exercise the Word reader rather than only source restoration.
        run("docx-to-typst", output, "--output", directory / "reconstructed", "--reconstruct")
    print(f"Installed wheel {docx_typst.__version__}: CLIs, starter, calculations, and roundtrip passed")


if __name__ == "__main__":
    main()
