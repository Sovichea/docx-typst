#!/usr/bin/env python3
"""Compare a source DOCX with a generated Typst project."""

from __future__ import annotations

import argparse
from collections import Counter
from difflib import SequenceMatcher
import json
from pathlib import Path
import re
from typing import Any

from .docx_reader import BaseDocxReader
from .docx_media import MediaMixin
from .docx_paragraphs import ParagraphMixin
from .docx_tables import TableMixin
from .model import CodeBlock, ImageBlock, Paragraph, TableBlock
from .typst_parser import parse_typst


class ValidationReader(ParagraphMixin, MediaMixin, TableMixin, BaseDocxReader):
    pass


def _normalize(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def _counts(document: Any) -> dict[str, int]:
    counts: Counter[str] = Counter()
    counts["styles"] = len(document.styles)
    counts["sections"] = getattr(document, "section_count", 1)

    def walk(blocks: list[Any]) -> None:
        for block in blocks:
            if isinstance(block, Paragraph):
                counts["paragraphs"] += 1
                if block.heading_level is not None:
                    counts["headings"] += 1
                if block.list_kind:
                    counts["list_items"] += 1
                counts["images"] += sum(
                    1 for run in block.runs if run.kind == "image"
                )
            elif isinstance(block, TableBlock):
                counts["tables"] += 1
                for row in block.rows:
                    for cell in row:
                        walk(cell.blocks)
            elif isinstance(block, ImageBlock):
                counts["images"] += 1
            elif isinstance(block, CodeBlock):
                counts["code_blocks"] += 1

    walk(document.blocks)
    return dict(counts)


def validate(source_docx: str | Path, project: str | Path, out: str | Path | None = None) -> dict[str, Any]:
    source_path = Path(source_docx)
    project_path = Path(project)
    main_typ = project_path / "main.typ"
    if not main_typ.is_file():
        raise FileNotFoundError(main_typ)

    source_document = ValidationReader(
        source_path, project_path, extract_media=False, keep_empty=True
    ).read()
    typst_document = parse_typst(main_typ)

    source_text = _normalize(source_document.plain_text())
    typst_text = _normalize(typst_document.plain_text())
    similarity = SequenceMatcher(None, source_text, typst_text).ratio()

    source_counts = _counts(source_document)
    typst_counts = _counts(typst_document)

    differences: list[dict[str, Any]] = []
    for key in ("paragraphs", "headings", "tables", "list_items", "images"):
        if source_counts.get(key, 0) != typst_counts.get(key, 0):
            differences.append(
                {
                    "kind": "count",
                    "name": key,
                    "source": source_counts.get(key, 0),
                    "typst": typst_counts.get(key, 0),
                }
            )
    for name, source_value, typst_value in (
        ("page_width_pt", source_document.page.width_pt, typst_document.page.width_pt),
        ("page_height_pt", source_document.page.height_pt, typst_document.page.height_pt),
        ("margin_top_pt", source_document.page.margin_top, typst_document.page.margin_top),
        ("margin_right_pt", source_document.page.margin_right, typst_document.page.margin_right),
        ("margin_bottom_pt", source_document.page.margin_bottom, typst_document.page.margin_bottom),
        ("margin_left_pt", source_document.page.margin_left, typst_document.page.margin_left),
    ):
        if abs(float(source_value) - float(typst_value)) > 1.0:
            differences.append(
                {
                    "kind": "page",
                    "name": name,
                    "source": round(float(source_value), 2),
                    "typst": round(float(typst_value), 2),
                }
            )
    missing_styles = sorted(set(source_document.styles) - set(typst_document.styles))
    if missing_styles:
        differences.append({"kind": "styles", "name": "missing", "values": missing_styles})
    if similarity < 0.98:
        differences.append(
            {
                "kind": "text",
                "name": "similarity",
                "source": 1.0,
                "typst": round(similarity, 4),
            }
        )

    report = {
        "source_file": source_path.name,
        "project": project_path.name,
        "status": "match" if not differences else "differences",
        "text_similarity": round(similarity, 4),
        "source": {
            "counts": source_counts,
            "styles": sorted(source_document.styles),
            "page": source_document.page.as_dict(),
        },
        "typst": {
            "counts": typst_counts,
            "styles": sorted(typst_document.styles),
            "page": typst_document.page.as_dict(),
        },
        "differences": differences,
    }
    destination = Path(out) if out else project_path / "validation.json"
    destination.parent.mkdir(parents=True, exist_ok=True)
    destination.write_text(
        json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    return report


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Compare a source DOCX document with a generated Typst project."
    )
    parser.add_argument("source_docx", type=Path)
    parser.add_argument("project", type=Path)
    parser.add_argument("--out", type=Path)
    parser.add_argument("--strict", action="store_true")
    args = parser.parse_args()
    report = validate(args.source_docx, args.project, args.out)
    print(
        json.dumps(
            {
                "status": report["status"],
                "text_similarity": report["text_similarity"],
                "differences": len(report["differences"]),
            },
            indent=2,
        )
    )
    if args.strict and report["status"] != "match":
        raise SystemExit(1)


if __name__ == "__main__":
    main()
