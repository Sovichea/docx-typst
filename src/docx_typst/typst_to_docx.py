#!/usr/bin/env python3
"""Convert the documented static Typst subset into editable DOCX."""

from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any

from . import __version__
from .docx_render import DocxWriter
from .static_render import materialize_static_blocks
from .typst_parser import parse_typst_text


def convert(
    main_typ: str | Path,
    output: str | Path,
    *,
    reference: str | Path | None = None,
    strict: bool = False,
) -> dict[str, Any]:
    source = Path(main_typ)
    if not source.is_file():
        raise FileNotFoundError(source)
    output_path = Path(output)
    text = source.read_text(encoding="utf-8")
    text, static_warnings = materialize_static_blocks(source, text, source.parent)
    document = parse_typst_text(source, text)
    for warning in static_warnings:
        document.warn(
            str(warning.get("issue", "docx_static")),
            str(warning.get("outcome", "approximated")),
            str(warning.get("location", source.name)),
            reason=str(warning.get("reason", "")),
        )
    writer = DocxWriter(
        document,
        output_path,
        base=source.parent,
        reference=Path(reference) if reference else None,
    )
    writer.write()

    status = "completed"
    if any(
        item.outcome in {"approximated", "unsupported", "skipped"}
        for item in document.diagnostics
    ):
        status = "completed_with_warnings"
    counts = dict(writer.counts)
    counts["styles"] = len(document.styles)
    counts["sections"] = counts.get("sections", 0) or 1
    counts["unsupported"] = sum(
        1 for item in document.diagnostics if item.outcome == "unsupported"
    )
    report = {
        "source_typ": source.name,
        "output_file": output_path.name,
        "generator": f"docx-typst {__version__}",
        "status": status,
        "counts": counts,
        "page": document.page.as_dict(),
        "diagnostics": [item.as_dict() for item in document.diagnostics],
        "styles": sorted(document.styles),
    }
    report_path = output_path.parent / "conversion_report.json"
    report_path.write_text(
        json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    if strict and any(item.outcome == "unsupported" for item in document.diagnostics):
        raise RuntimeError(
            "strict conversion found unsupported Typst content; see conversion_report.json"
        )
    return report


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Convert a supported Typst document to editable DOCX."
    )
    parser.add_argument("main_typ", type=Path, help="source Typst document")
    parser.add_argument(
        "--output",
        type=Path,
        default=Path("converted.docx"),
        help="output DOCX path",
    )
    parser.add_argument(
        "--reference",
        type=Path,
        help="optional reference DOCX whose styles and theme are reused",
    )
    parser.add_argument(
        "--strict", action="store_true", help="fail when unsupported content is found"
    )
    args = parser.parse_args()
    report = convert(
        args.main_typ,
        args.output,
        reference=args.reference,
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
