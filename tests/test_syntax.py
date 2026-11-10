from __future__ import annotations

from docx_typst.model import Document
from docx_typst.typst_inline import InlineParser
from docx_typst.typst_syntax import (
    UnitValue,
    balanced,
    color_to_rgb,
    length_to_pt,
    parse_call_at,
    parse_value,
    split_top_level,
    typst_escape,
    typst_string,
)
from docx_typst.typst_values import page_setup_from_dict, style_from_spec


def test_value_parser_reads_generated_style_values() -> None:
    value = parse_value(
        '(font: "Calibri", size: 11pt, fill: rgb("#C02020"), '
        "bold: true, before: 12pt, line-spacing: 1.15)"
    )
    assert value["font"] == "Calibri"
    assert value["size"] == UnitValue(11.0, "pt")
    assert value["fill"] == "#C02020"
    assert value["bold"] is True
    assert length_to_pt(value["before"]) == 12.0


def test_value_parser_reads_arrays_and_nested_dicts() -> None:
    value = parse_value("((0, 1, 2, 3), (1, 0, 1, 2))")
    assert value == [[0, 1, 2, 3], [1, 0, 1, 2]]
    assert parse_value("(:)") == {}
    assert parse_value('("a",)') == ["a"]


def test_balanced_and_split_respect_nested_calls_and_strings() -> None:
    text = 'outer(inner("a,b"), [x, y], z)'
    inner, end = balanced(text, 5, "(", ")")
    assert inner == 'inner("a,b"), [x, y], z'
    assert end == len(text)
    assert split_top_level(inner) == ['inner("a,b")', "[x, y]", "z"]


def test_parse_call_at_includes_content_body() -> None:
    call = parse_call_at('#docx-para(style: "Normal")[Hello *world*]', 1)
    assert call is not None
    assert call.name == "docx-para"
    assert call.arg("style") == "Normal"
    assert call.body is not None
    assert call.body.text == "Hello *world*"


def test_colors_and_lengths() -> None:
    assert color_to_rgb(parse_value('rgb("#C02020")')) == (192, 32, 32)
    assert color_to_rgb(parse_value("rgb(255, 0, 0)")) == (255, 0, 0)
    assert length_to_pt(parse_value("2.54cm")) == 72.0
    assert length_to_pt(parse_value("1in")) == 72.0


def test_typst_escaping_roundtrips_common_markup() -> None:
    escaped = typst_escape("C# and *stars* [brackets] <label> @ref")
    assert "\\#" in escaped
    assert "\\*" in escaped
    assert "\\[" in escaped
    assert "\\<" in escaped
    assert "\\@" in escaped
    assert typst_string('a "quote"') == '"a \\"quote\\""'


def test_inline_parser_applies_nested_formatting_and_collapses_spaces() -> None:
    document = Document()
    runs = InlineParser(document).parse("  *bold _and italic_*  ")
    assert "".join(run.text for run in runs) == "bold and italic"
    assert all(run.bold for run in runs)
    assert any(run.italic for run in runs)


def test_inline_parser_reads_links_code_and_fields() -> None:
    document = Document()
    runs = InlineParser(document).parse(
        '#link("https://example.com")[site] `code` #docx-page-number()'
    )
    linked = next(run for run in runs if run.text == "site")
    code = next(run for run in runs if run.text == "code")
    field = next(run for run in runs if run.kind == "field")
    assert linked.link == "https://example.com"
    assert code.code is True
    assert field.field == "PAGE"


def test_style_and_page_conversion() -> None:
    style = style_from_spec(
        "Heading 1",
        {
            "font": "Calibri",
            "size": UnitValue(18.0, "pt"),
            "fill": "#2F5496",
            "before": UnitValue(18.0, "pt"),
            "keep-with-next": True,
        },
    )
    assert style.name == "Heading 1"
    assert style.font == "Calibri"
    assert style.size == 18.0
    assert style.color == "#2F5496"
    assert style.space_before == 18.0
    assert style.keep_with_next is True

    page = page_setup_from_dict(
        {
            "paper": "a4",
            "margin": {"top": UnitValue(2.0, "cm"), "x": UnitValue(1.5, "cm")},
            "columns": 2,
        }
    )
    assert round(page.width_pt, 1) == 595.3
    assert round(page.margin_top, 1) == 56.7
    assert round(page.margin_left, 1) == 42.5
    assert page.columns == 2
