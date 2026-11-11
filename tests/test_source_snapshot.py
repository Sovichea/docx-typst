from pathlib import Path
import shutil

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from PIL import Image
import pytest
import typst

from docx_typst import docx_to_typst, typst_to_docx

ROOT = Path(__file__).resolve().parents[1]


@pytest.fixture
def exported(tmp_path: Path):
    shutil.copyfile(ROOT / "src/docx_typst/template.typ", tmp_path / "template.typ")
    Image.new("RGB", (40, 20), "red").save(tmp_path / "picture.png")
    source = tmp_path / "original.typ"
    source.write_bytes(b'''#import "template.typ": *

// Preserve native comments, definitions, and structured call sites.

#let caption(body) = text(size: 9pt, style: "italic", body)

#let docx-styles = ("Normal": (font: "Arial", size: 11pt))

#show: docx-document.with(page: (paper: "a4", margin: 65pt))

= Native heading

Original body text.

#v(0.2cm)

#caption[Native caption.]

#docx-image("picture.png", width: 40pt, align: "right")

#docx-table(columns: 1, rows: 1, stroke: 0.5pt + rgb("DCE3DF"), cells: ([Cell],))

''')
    output = tmp_path / "out.docx"
    typst_to_docx(source, output)
    return source, output


def test_unchanged_export_restores_exact_source_and_dependencies(exported, tmp_path: Path) -> None:
    source, output = exported
    original = source.read_bytes()
    # The source file on disk is not used to reconstruct the exported document.
    source.write_text("A different file now lives here.", encoding="utf-8")
    report = docx_to_typst(output, tmp_path / "restored")
    assert report["source_restoration"] == "exact"
    assert (tmp_path / "restored/main.typ").read_bytes() == original
    assert (tmp_path / "restored/picture.png").read_bytes() == (tmp_path / "picture.png").read_bytes()
    typst.compile(str(tmp_path / "restored/main.typ"))


@pytest.mark.parametrize("edit", ["text", "border", "image", "image-content", "format"])
def test_word_edits_are_not_replaced_by_stale_source(exported, tmp_path: Path, edit: str) -> None:
    source, output = exported
    document = Document(output)
    if edit == "text":
        next(p for p in document.paragraphs if p.text == "Original body text.").text = "Edited in Word."
    elif edit == "border":
        document.tables[0]._tbl.tblPr.find(qn("w:tblBorders")).find(qn("w:top")).set(qn("w:color"), "FF0000")
    elif edit == "image-content":
        from io import BytesIO
        blob = BytesIO()
        Image.new("RGB", (40, 20), "blue").save(blob, format="PNG")
        rel_id = document.inline_shapes[0]._inline.xpath(".//a:blip")[0].get(qn("r:embed"))
        document.part.rels[rel_id].target_part._blob = blob.getvalue()
    elif edit == "image":
        next(p for p in document.paragraphs if p._p.xpath(".//w:drawing")).alignment = WD_ALIGN_PARAGRAPH.LEFT
    else:
        document.styles["Normal"].font.bold = True
    document.save(output)
    report = docx_to_typst(output, tmp_path / "restored")
    assert report["source_restoration"] == "reconstructed"
    assert any(d["issue"] in {"native_source_changed", "native_source_asset_changed"} for d in report["diagnostics"])
    restored = (tmp_path / "restored/main.typ").read_text(encoding="utf-8")
    assert restored != source.read_text(encoding="utf-8")
    if edit == "text":
        assert "Edited in Word." in restored
        assert "Original body text." not in restored
    typst.compile(str(tmp_path / "restored/main.typ"))


def test_reconstruction_can_be_requested_explicitly(exported, tmp_path: Path) -> None:
    _, output = exported
    report = docx_to_typst(output, tmp_path / "restored", restore_source=False)
    assert report["source_restoration"] == "reconstructed"
    assert "#docx-heading" in (tmp_path / "restored/main.typ").read_text(encoding="utf-8")


@pytest.mark.parametrize("entrypoint", ["docx_typst.cli", "docx_typst.docx_to_typst"])
def test_cli_reconstruct_bypasses_source_snapshot(exported, tmp_path: Path, monkeypatch, entrypoint: str) -> None:
    import importlib
    import json
    _, output = exported
    target = tmp_path / "cli-restored"
    argv = ["converter"] + (["to-typst"] if entrypoint.endswith(".cli") else [])
    argv += [str(output), "--output", str(target), "--reconstruct"]
    monkeypatch.setattr("sys.argv", argv)
    importlib.import_module(entrypoint).main()
    report = json.loads((target / "conversion_report.json").read_text(encoding="utf-8"))
    assert report["source_restoration"] == "reconstructed"


def test_snapshot_collects_assets_relative_to_helper_definitions(tmp_path: Path) -> None:
    from docx_typst.source_preservation import _dependencies
    import base64
    helper = tmp_path / "helpers"
    (helper / "assets").mkdir(parents=True)
    (helper / "template.typ").write_text('#let picture(path) = image(path)', encoding="utf-8")
    (helper / "assets/logo.svg").write_text('<svg xmlns="http://www.w3.org/2000/svg"/>', encoding="utf-8")
    source = '#import "helpers/template.typ": picture\n#picture("assets/logo.svg")'
    files = _dependencies(tmp_path, source)
    assert base64.b64decode(files["helpers/assets/logo.svg"]) == (helper / "assets/logo.svg").read_bytes()
