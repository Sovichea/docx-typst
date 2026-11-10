"""Render ``docx-static`` bodies to images before DOCX export."""

from __future__ import annotations

import hashlib
from pathlib import Path
import re
from typing import Any

from PIL import Image, ImageChops
from lxml import etree
import typst

from .typst_syntax import (
    TypstCall,
    UnitValue,
    parse_call_at,
)


def _import_preamble(text: str) -> str:
    """Copy top-level imports and simple ``#let`` definitions."""
    lines = text.splitlines()
    picked: list[str] = []
    i = 0
    while i < len(lines):
        line = lines[i]
        stripped = line.strip().lstrip("\ufeff")
        if stripped.startswith("#import") or stripped.startswith("#let "):
            picked.append(line)
            # Continue across balanced parentheses.
            balance = line.count("(") - line.count(")")
            while balance > 0 and i + 1 < len(lines):
                i += 1
                picked.append(lines[i])
                balance += lines[i].count("(") - lines[i].count(")")
            i += 1
            continue
        if not stripped or stripped.startswith("//"):
            i += 1
            continue
        break
    return "\n".join(picked)


def _width_pt(value: Any, default: float = 465.0) -> float:
    if isinstance(value, UnitValue):
        return value.to_pt()
    if isinstance(value, (int, float)) and value > 10:
        return float(value)
    return default


def _render_body(
    project_root: Path,
    preamble: str,
    body: str,
    width_pt: float,
    target: Path,
    height_pt: float | None = None,
) -> str | None:
    if target.suffix == ".svg":
        height = f"{height_pt:.6f}pt" if height_pt is not None else "auto"
        wrapper = project_root / f".docx-static-{target.stem}.typ"
        if height_pt is None:
            # A diagram may have an intrinsic frame wider than its requested
            # export hint. Measure and retain that complete native frame;
            # clipping at the requested width silently drops drawing content.
            drawing = [
                "#set page(width: auto, height: auto, margin: 0pt, fill: none)",
                "#context {",
                "let artwork = [", body, "]",
                "let bounds = measure(artwork)",
                "block(width: bounds.width, height: bounds.height, clip: false, artwork)",
                "}",
            ]
        else:
            drawing = [
                f"#set page(width: {width_pt:.6f}pt, height: {height}, margin: 0pt, fill: none)",
                f"#block(width: {width_pt:.6f}pt, height: {height}, clip: true)[",
                body, "]",
            ]
        wrapper.write_text("\n".join([
            preamble, *drawing,
        ]), encoding="utf-8")
        try:
            # Compile the same bounded frame twice: native vectors plus the
            # compatibility bitmap Word requires alongside its SVG extension.
            typst.compile(str(wrapper), output=str(target), root=str(project_root), format="svg")
            typst.compile(str(wrapper), output=str(target.with_suffix(".png")),
                          root=str(project_root), format="png", ppi=200)
            return None
        except Exception as exc:
            return str(exc)[:500]
        finally:
            wrapper.unlink(missing_ok=True)
    page_width = max(width_pt, 1200.0)
    wrapper = project_root / f".docx-static-{target.stem}.typ"
    wrapper.write_text(
        "\n".join(
            [
                preamble,
                f'#set page(width: {page_width:.2f}pt, height: 2000pt, margin: 0pt, fill: white)',
                f"#box(width: {width_pt:.2f}pt)[",
                body,
                "]",
                "",
            ]
        ),
        encoding="utf-8",
    )
    try:
        typst.compile(
            str(wrapper),
            output=str(target),
            root=str(project_root),
            format="png",
            ppi=200,
        )
        if not target.is_file():
            return "Typst did not produce an image"
        image = Image.open(target).convert("RGB")
        background = Image.new("RGB", image.size, (255, 255, 255))
        bbox = ImageChops.difference(image, background).getbbox()
        if bbox:
            left, top, right, bottom = bbox
            image.crop(
                (
                    max(0, left - 2),
                    max(0, top - 2),
                    min(image.width, right + 2),
                    min(image.height, bottom + 2),
                )
            ).save(target)
        return None
    except Exception as exc:
        return str(exc)[:500]
    finally:
        wrapper.unlink(missing_ok=True)


def materialize_static_blocks(
    source: Path, text: str, project_root: Path
) -> tuple[str, list[dict[str, str]]]:
    """Materialize export-only assets without changing the authored source."""
    preamble = _import_preamble(text)
    pattern = re.compile(r"#docx-static\s*\(")
    insertions: list[tuple[int, str]] = []
    warnings: list[dict[str, str]] = []
    count = 0
    for match in pattern.finditer(text):
        call = parse_call_at(text, match.start() + 1)
        if call is None or call.body is None:
            continue
        if isinstance(call.arg("path"), str):
            continue
        count += 1
        width = _width_pt(call.arg("width"))
        fmt = call.arg("format", "png")
        if fmt not in {"png", "svg"}:
            raise ValueError(f"unsupported docx-static format: {fmt}")
        height_value = call.arg("height")
        height = _width_pt(height_value) if isinstance(height_value, UnitValue) else None
        digest = hashlib.sha256((preamble + call.raw).encode("utf-8")).hexdigest()[:10]
        relative = f"assets/static/docx-static-{count}-{digest}.{fmt}"
        target = project_root / relative
        target.parent.mkdir(parents=True, exist_ok=True)
        error = _render_body(project_root, preamble, call.body.text, width, target, height)
        if error:
            line_number = text[: match.start()].count("\n") + 1
            warnings.append(
                {
                    "issue": "docx_static_render",
                    "outcome": "approximated",
                    "location": f"line {line_number}",
                    "reason": error,
                }
            )
            continue
        native_size = ""
        if fmt == "svg":
            viewbox = etree.fromstring(target.read_bytes()).get("viewBox", "").split()
            if len(viewbox) == 4:
                native_size = f"native-width: {viewbox[2]}pt, native-height: {viewbox[3]}pt, "
        insertions.append((match.end(), f'path: "{relative}", ' + native_size))
    for position, insertion in reversed(insertions):
        text = text[:position] + insertion + text[position:]
    return text, warnings
