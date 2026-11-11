from pathlib import Path
import base64
import json
import random
import shutil

from docx import Document
from docx.opc.packuri import PackURI
from PIL import Image
import pytest

from docx_typst import docx_to_typst, typst_to_docx
from docx_typst.source_preservation import REL
from docx_typst.source_assets import unpack_files

ROOT = Path(__file__).resolve().parents[1]


@pytest.fixture
def exported(tmp_path: Path):
    shutil.copyfile(ROOT / "src/docx_typst/template.typ", tmp_path / "template.typ")
    Image.frombytes("RGB", (128, 128), random.Random(42).randbytes(128 * 128 * 3)).save(tmp_path / "picture.png")
    shutil.copyfile(tmp_path / "picture.png", tmp_path / "alias.png")
    (tmp_path / "assets").mkdir()
    shutil.copyfile(tmp_path / "picture.png", tmp_path / "assets/logo.png")
    (tmp_path / "extra.bin").write_bytes(b"An original source dependency that Word does not display.")
    shutil.copyfile(tmp_path / "extra.bin", tmp_path / "extra-alias.bin")
    source = tmp_path / "original.typ"
    source.write_text('''#import "template.typ": *
#let source-assets = ("alias.png", "assets/logo.png", "extra.bin", "extra-alias.bin")
#show: docx-document.with(page: (paper: "a4", margin: 65pt))
#docx-image("picture.png", width: 40pt)
''', encoding="utf-8")
    target = tmp_path / "out.docx"
    typst_to_docx(source, target)
    return source, target


def test_media_and_private_assets_are_stored_once(exported, tmp_path: Path) -> None:
    source, target = exported
    document = Document(target)
    part = document.part.part_related_by(REL)
    data = json.loads(part.blob)
    assert data["version"] == 2
    refs = [data["files"][n] for n in ("picture.png", "alias.png", "assets/logo.png")]
    assert refs[0] == refs[1] == refs[2]
    assert "relationship" not in refs[0]
    assert all(not str(r.target_part.partname).startswith("/word/media/") for r in part.rels.values())
    assert len([p for p in document.part.package.iter_parts() if str(p.partname).startswith("/word/media/")]) == 1
    private = [p for p in document.part.package.iter_parts() if str(p.partname).startswith("/docxTypst/assets/")]
    assert len(private) == 1
    assert data["files"]["extra.bin"] == data["files"]["extra-alias.bin"]
    report = docx_to_typst(target, tmp_path / "restored")
    assert report["source_restoration"] == "exact"
    assert (tmp_path / "restored/main.typ").read_bytes() == source.read_bytes()
    for relative in data["files"]:
        assert (tmp_path / "restored" / relative).read_bytes() == (tmp_path / relative).read_bytes()


def test_asset_hash_lookup_survives_word_part_and_relationship_renames(exported, tmp_path: Path) -> None:
    _, target = exported
    document = Document(target)
    part = document.part.part_related_by(REL)
    data = json.loads(part.blob)
    image = next(p for p in document.part.package.iter_parts() if str(p.partname).startswith("/word/media/"))
    image._partname = PackURI("/word/media/renamed.png")
    for value in data["files"].values():
        if isinstance(value, dict):
            value["relationship"] = "old-relationship-name"
    part._blob = json.dumps(data).encode()
    document.save(target)
    assert docx_to_typst(target, tmp_path / "restored")["source_restoration"] == "exact"


def test_version_one_inline_dependencies_remain_readable(exported, tmp_path: Path) -> None:
    source, target = exported
    document = Document(target)
    part = document.part.part_related_by(REL)
    data = json.loads(part.blob)
    files, issues = unpack_files(part, data["files"])
    assert not issues
    data["files"] = {name: base64.b64encode(blob).decode() for name, blob in files.items()}
    data["version"] = 1
    part._blob = json.dumps(data).encode()
    document.save(target)
    assert docx_to_typst(target, tmp_path / "restored")["source_restoration"] == "exact"
    assert (tmp_path / "restored/main.typ").read_bytes() == source.read_bytes()


def test_missing_original_asset_blocks_exact_restoration(exported, tmp_path: Path) -> None:
    _, target = exported
    document = Document(target)
    part = document.part.part_related_by(REL)
    data = json.loads(part.blob)
    record = data["files"]["extra.bin"]
    record["sha256"] = "0" * 64
    record["relationship"] = "missing"
    part._blob = json.dumps(data).encode()
    document.save(target)
    report = docx_to_typst(target, tmp_path / "restored")
    assert report["source_restoration"] == "reconstructed"
    assert any(d["issue"] == "native_source_asset_changed" for d in report["diagnostics"])


def test_recompaction_removes_word_created_source_media_duplicates(exported, tmp_path: Path) -> None:
    from docx.opc.part import Part
    from docx_typst.source_assets import ASSET_REL, pack_files
    _, target = exported
    document = Document(target)
    part = document.part.part_related_by(REL)
    data = json.loads(part.blob)
    image = next(p for p in document.part.package.iter_parts() if str(p.partname).startswith("/word/media/"))
    duplicate = Part(PackURI("/word/media/source-copy.png"), image.content_type, image.blob, part.package)
    data["files"]["picture.png"]["relationship"] = part.relate_to(duplicate, ASSET_REL)
    part._blob = json.dumps(data).encode()
    document.save(target)
    assert len([p for p in document.part.package.iter_parts() if str(p.partname).startswith("/word/media/")]) == 2
    data["files"] = pack_files(part, data["files"])
    part._blob = json.dumps(data).encode()
    document.save(target)
    saved = Document(target)
    assert len([p for p in saved.part.package.iter_parts() if str(p.partname).startswith("/word/media/")]) == 1
    assert docx_to_typst(target, tmp_path / "restored")["source_restoration"] == "exact"
