#!/usr/bin/env python3
"""Convert a supported Word document to an editable Typst project."""

from __future__ import annotations

import argparse
import json
from pathlib import Path
import shutil
from typing import Any

from . import __version__
from .docx_media import MediaMixin
from .docx_paragraphs import ParagraphMixin
from .docx_reader import BaseDocxReader
from .docx_tables import TableMixin
from .typst_render import render_document


class DocxReader(ParagraphMixin, MediaMixin, TableMixin, BaseDocxReader):
    """Forward reader composed of paragraph, media, and table mixins."""


def convert(
    source: str | Path,
    output: str | Path,
    *,
    template: str | Path | None = None,
    extract_media: bool = True,
    keep_empty: bool = True,
    strict: bool = False,
) -> dict[str, Any]:
    source_path = Path(source)
    output_path = Path(output)
    output_path.mkdir(parents=True, exist_ok=True)
    (output_path / "assets" / "media").mkdir(parents=True, exist_ok=True)
    reader = DocxReader(
        source_path,
        output_path,
        extract_media=extract_media,
        keep_empty=keep_empty,
    )
    document = reader.read()
    main_text = render_document(document)
    imports = reader.native_sources.get("imports", "")
    if imports:
        main_text = main_text.replace('#import "template.typ": *',
                                      '#import "template.typ": *\n' + imports, 1)
    (output_path / "main.typ").write_text(main_text, encoding="utf-8")

    template_source = (
        Path(template) if template else Path(__file__).with_name("template.typ")
    )
    template_target = output_path / "template.typ"
    shutil.copyfile(template_source, template_target)

    manifest = {
        "source_file": source_path.name,
        "generator": f"docx-typst {__version__}",
        "default_font": document.default_font,
        "default_size_pt": document.default_size,
        "page": document.page.as_dict(),
        "sections": document.section_count,
        "assets": [
            {
                "original_name": asset.original_name,
                "output_path": asset.output_path,
                "derived_path": asset.derived_path,
                "sha256": asset.sha256,
                "bytes": asset.byte_count,
                "content_type": asset.content_type,
                "source_part": asset.source_part,
                "emitted": asset.emitted,
                "warning": asset.warning,
            }
            for asset in document.assets
        ],
        "styles": {
            name: spec.spec() for name, spec in sorted(document.styles.items())
        },
    }
    (output_path / "asset_manifest.json").write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8"
    )

    status = "completed"
    if any(
        item.outcome in {"approximated", "unsupported", "skipped"}
        for item in document.diagnostics
    ):
        status = "completed_with_warnings"
    report = {
        "source_file": source_path.name,
        "generator": f"docx-typst {__version__}",
        "status": status,
        "counts": dict(reader.counts),
        "page": document.page.as_dict(),
        "diagnostics": [item.as_dict() for item in document.diagnostics],
        "styles": sorted(document.styles),
    }
    (output_path / "conversion_report.json").write_text(
        json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    if strict and any(item.outcome == "unsupported" for item in document.diagnostics):
        raise RuntimeError(
            "strict conversion found unsupported DOCX content; see conversion_report.json"
        )
    return report


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Convert a DOCX document to an editable Typst project."
    )
    parser.add_argument("input_docx", type=Path, help="source Word document")
    parser.add_argument(
        "--output",
        type=Path,
        default=Path("./typst-output"),
        help="Typst project output directory",
    )
    parser.add_argument(
        "--template",
        type=Path,
        help="optional Typst template to copy instead of the bundled one",
    )
    parser.add_argument("--no-media", action="store_true", help="do not extract images")
    parser.add_argument(
        "--drop-empty",
        action="store_true",
        help="omit empty paragraphs instead of preserving vertical space",
    )
    parser.add_argument(
        "--strict", action="store_true", help="fail when unsupported content is found"
    )
    args = parser.parse_args()
    report = convert(
        args.input_docx,
        args.output,
        template=args.template,
        extract_media=not args.no_media,
        keep_empty=not args.drop_empty,
        strict=args.strict,
    )
    counts = report["counts"]
    print(
        "Saved",
        Path(args.output).resolve(),
        f"paragraphs={counts.get('paragraphs', 0)}",
        f"tables={counts.get('tables', 0)}",
        f"images={counts.get('images', 0)}",
        f"warnings={len(report['diagnostics'])}",
    )


if __name__ == "__main__":
    main()
