"""Convert parsed Typst values into document-model values."""

from __future__ import annotations

from typing import Any

from .model import ImageRef, PageSetup, StyleDef
from .typst_syntax import (
    Identifier,
    RawExpr,
    TypstCall,
    UnitValue,
    color_to_rgb,
    length_to_pt,
)


PAPER_SIZES = {
    "a4": (595.28, 841.89),
    "a5": (419.53, 595.28),
    "a3": (841.89, 1190.55),
    "us-letter": (612.0, 792.0),
    "us-legal": (612.0, 1008.0),
    "us-tabloid": (792.0, 1224.0),
}


def string_value(value: Any) -> str | None:
    if value is None:
        return None
    if isinstance(value, str):
        return value
    if isinstance(value, (Identifier, RawExpr)):
        return str(value)
    return None


def style_from_spec(name: str, spec: dict[str, Any]) -> StyleDef:
    style = StyleDef(name=name)
    style.based_on = string_value(spec.get("based-on"))
    style.next_style = string_value(spec.get("next-style"))
    font = spec.get("font")
    if isinstance(font, str):
        style.font = font
    east_asia = spec.get("east-asia-font")
    if isinstance(east_asia, str):
        style.east_asia_font = east_asia
    word_font = spec.get("word-font")
    if isinstance(word_font, str):
        style.word_font = word_font
    weight = spec.get("weight")
    if isinstance(weight, str):
        style.weight = weight
    tracking = spec.get("tracking")
    if tracking is not None:
        style.tracking = length_to_pt(tracking)
    leading = spec.get("leading")
    if leading is not None:
        style.leading = length_to_pt(leading)
    size = spec.get("size")
    if size is not None:
        style.size = length_to_pt(size)
    for key in ("bold", "italic", "underline", "strike"):
        value = spec.get(key)
        if isinstance(value, bool):
            setattr(style, key, value)
    for key, attr_name in (("small-caps", "small_caps"), ("all-caps", "all_caps")):
        value = spec.get(key)
        if isinstance(value, bool):
            setattr(style, attr_name, value)
    fill = color_to_rgb(spec.get("fill"))
    if fill is not None:
        style.color = "#" + "".join(f"{value:02X}" for value in fill)
    highlight = color_to_rgb(spec.get("highlight"))
    if highlight is not None:
        style.highlight = "#" + "".join(f"{value:02X}" for value in highlight)
    align = spec.get("align")
    if isinstance(align, str):
        style.align = align
    justify = spec.get("justify")
    if isinstance(justify, bool):
        style.justify = justify
    for key, attr_name in (
        ("before", "space_before"),
        ("after", "space_after"),
        ("word-before", "word_before"),
        ("word-after", "word_after"),
        ("line-spacing-pt", "line_spacing_pt"),
        ("left-indent", "left_indent"),
        ("right-indent", "right_indent"),
        ("first-line-indent", "first_line_indent"),
    ):
        value = spec.get(key)
        if value is not None:
            setattr(style, attr_name, length_to_pt(value))
    value = spec.get("line-spacing")
    if value is not None:
        style.line_spacing = float(value)
    for key, attr_name in (
        ("keep-with-next", "keep_with_next"),
        ("keep-lines", "keep_lines"),
        ("page-break-before", "page_break_before"),
    ):
        value = spec.get(key)
        if isinstance(value, bool):
            setattr(style, attr_name, value)
    value = spec.get("outline-level")
    if isinstance(value, int):
        style.outline_level = value
    return style


def page_setup_from_dict(
    values: dict[str, Any], base: PageSetup | None = None
) -> PageSetup:
    page = PageSetup(**base.as_dict()) if base is not None else PageSetup()
    paper = values.get("paper")
    if isinstance(paper, str) and paper.lower() in PAPER_SIZES:
        page.width_pt, page.height_pt = PAPER_SIZES[paper.lower()]
    width = values.get("width")
    height = values.get("height")
    if width is not None:
        page.width_pt = length_to_pt(width, page.width_pt)
    if height is not None:
        page.height_pt = length_to_pt(height, page.height_pt)
    if values.get("flipped") is True:
        page.width_pt, page.height_pt = page.height_pt, page.width_pt
    margin = values.get("margin")
    if isinstance(margin, UnitValue):
        page.margin_top = page.margin_right = margin.to_pt()
        page.margin_bottom = page.margin_left = margin.to_pt()
    elif isinstance(margin, dict):
        x_value = margin.get("x")
        y_value = margin.get("y")
        rest = margin.get("rest")
        page.margin_top = length_to_pt(
            margin.get("top", y_value if y_value is not None else rest),
            page.margin_top,
        )
        page.margin_bottom = length_to_pt(
            margin.get("bottom", y_value if y_value is not None else rest),
            page.margin_bottom,
        )
        page.margin_left = length_to_pt(
            margin.get("left", x_value if x_value is not None else rest),
            page.margin_left,
        )
        page.margin_right = length_to_pt(
            margin.get("right", x_value if x_value is not None else rest),
            page.margin_right,
        )
    columns = values.get("columns")
    if isinstance(columns, int):
        page.columns = max(1, columns)
    page.orientation = "landscape" if page.width_pt > page.height_pt else "portrait"
    return page


def image_from_call(call: TypstCall) -> ImageRef | None:
    path = call.positional[0] if call.positional else call.arg("path")
    if not isinstance(path, str):
        return None
    width = call.arg("width")
    height = call.arg("height")
    alt = call.arg("alt")
    width_pt = None
    width_percent = None
    if isinstance(width, UnitValue):
        width_pt = width.to_pt()
    elif isinstance(width, (int, float)):
        if 0 < width <= 1:
            width_percent = float(width)
        elif width > 1:
            width_pt = float(width)
    return ImageRef(
        path=path,
        width_pt=width_pt,
        height_pt=length_to_pt(height) if height is not None else None,
        width_percent=width_percent,
        alt=alt if isinstance(alt, str) else None,
    )
