"""Render the shared document model as readable Typst source."""

from __future__ import annotations

import re
from typing import Any, Iterable

from . import __version__
from .model import (
    Block,
    CodeBlock,
    Document,
    HorizontalRule,
    ImageBlock,
    PageBreak,
    PageSetup,
    Paragraph,
    Run,
    SectionBreak,
    StyleDef,
    TableBlock,
    TableCell,
    UnsupportedBlock,
)
from .typst_syntax import format_length, format_number, typst_escape, typst_string


STANDARD_SIZES = {
    "a4": (595.28, 841.89),
    "a5": (419.53, 595.28),
    "a3": (841.89, 1190.55),
    "us-letter": (612.0, 792.0),
    "us-legal": (612.0, 1008.0),
}


# ---------------------------------------------------------------------------
# Style and page rendering
# ---------------------------------------------------------------------------


def style_entry(name: str, spec: StyleDef) -> str:
    fields: list[str] = []
    if spec.based_on:
        fields.append(f"based-on: {typst_string(spec.based_on)}")
    if spec.next_style:
        fields.append(f"next-style: {typst_string(spec.next_style)}")
    if spec.font:
        fields.append(f"font: {typst_string(spec.font)}")
    if spec.east_asia_font:
        fields.append(f"east-asia-font: {typst_string(spec.east_asia_font)}")
    if spec.size is not None:
        fields.append(f"size: {format_length(spec.size)}")
    if spec.bold is not None:
        fields.append(f"bold: {_bool(spec.bold)}")
    if spec.italic is not None:
        fields.append(f"italic: {_bool(spec.italic)}")
    if spec.underline is not None:
        fields.append(f"underline: {_bool(spec.underline)}")
    if spec.strike is not None:
        fields.append(f"strike: {_bool(spec.strike)}")
    if spec.small_caps is not None:
        fields.append(f"small-caps: {_bool(spec.small_caps)}")
    if spec.all_caps is not None:
        fields.append(f"all-caps: {_bool(spec.all_caps)}")
    if spec.color:
        fields.append(f"fill: rgb({typst_string(spec.color)})")
    if spec.highlight:
        fields.append(f"highlight: rgb({typst_string(spec.highlight)})")
    if spec.align:
        fields.append(f"align: {typst_string(spec.align)}")
    if spec.justify is not None:
        fields.append(f"justify: {_bool(spec.justify)}")
    if spec.space_before is not None:
        fields.append(f"before: {format_length(spec.space_before)}")
    if spec.space_after is not None:
        fields.append(f"after: {format_length(spec.space_after)}")
    if spec.line_spacing is not None:
        fields.append(f"line-spacing: {format_number(spec.line_spacing)}")
    if spec.line_spacing_pt is not None:
        fields.append(f"line-spacing-pt: {format_length(spec.line_spacing_pt)}")
    if spec.left_indent is not None:
        fields.append(f"left-indent: {format_length(spec.left_indent)}")
    if spec.right_indent is not None:
        fields.append(f"right-indent: {format_length(spec.right_indent)}")
    if spec.first_line_indent is not None:
        fields.append(f"first-line-indent: {format_length(spec.first_line_indent)}")
    if spec.keep_with_next is not None:
        fields.append(f"keep-with-next: {_bool(spec.keep_with_next)}")
    if spec.keep_lines is not None:
        fields.append(f"keep-lines: {_bool(spec.keep_lines)}")
    if spec.page_break_before is not None:
        fields.append(f"page-break-before: {_bool(spec.page_break_before)}")
    if spec.outline_level is not None:
        fields.append(f"outline-level: {spec.outline_level}")
    content = ", ".join(fields)
    return f"  {typst_string(name)}: ({content}),"


def render_styles(document: Document) -> str:
    lines = ["#let docx-styles = ("]
    names = list(document.styles)
    names.sort(key=lambda name: (name != "Normal", _style_sort_key(name), name.lower()))
    for name in names:
        lines.append(style_entry(name, document.styles[name]))
    lines.append(")")
    return "\n".join(lines)


def _style_sort_key(name: str) -> tuple[int, int, str]:
    match = re.match(r"^Heading\s+(\d+)$", name)
    if match:
        return (1, int(match.group(1)), "")
    return (2 if name != "Normal" else 0, 0, name.lower())


def page_dict(page: PageSetup, *, include_header_footer: bool = False) -> str:
    fields: list[str] = []
    size_name = _standard_size(page.width_pt, page.height_pt)
    if size_name:
        fields.append(f"paper: {typst_string(size_name)}")
    else:
        fields.append(f"width: {format_length(page.width_pt)}")
        fields.append(f"height: {format_length(page.height_pt)}")
    margins = (
        page.margin_top,
        page.margin_right,
        page.margin_bottom,
        page.margin_left,
    )
    if max(margins) - min(margins) < 0.01:
        fields.append(f"margin: {format_length(page.margin_top)}")
    else:
        fields.append(
            "margin: ("
            f"top: {format_length(page.margin_top)}, "
            f"right: {format_length(page.margin_right)}, "
            f"bottom: {format_length(page.margin_bottom)}, "
            f"left: {format_length(page.margin_left)})"
        )
    if page.columns and page.columns > 1:
        fields.append(f"columns: {page.columns}")
    return "(" + ", ".join(fields) + ")"


def _standard_size(width: float, height: float) -> str | None:
    # Typst's paper names are portrait; landscape pages must carry explicit
    # dimensions so the orientation survives the roundtrip.
    if width > height:
        return None
    for name, (w, h) in STANDARD_SIZES.items():
        if abs(width - w) < 1.0 and abs(height - h) < 1.0:
            return name
    return None


# ---------------------------------------------------------------------------
# Inline rendering
# ---------------------------------------------------------------------------


def render_runs(runs: Iterable[Run]) -> str:
    parts: list[str] = []
    for run in runs:
        rendered = render_run(run)
        if rendered:
            parts.append(rendered)
    return "".join(parts)


def render_run(run: Run) -> str:
    if run.kind == "break":
        return "#linebreak()"
    if run.kind == "page-break":
        return "#pagebreak()"
    if run.kind == "column-break":
        return "#colbreak()"
    if run.kind == "tab":
        return "#h(1.5em)"
    if run.kind == "image":
        return _render_inline_image(run)
    if run.kind == "field":
        if run.field == "PAGE":
            return "#docx-page-number()"
        if run.field in {"NUMPAGES", "SECTIONPAGES"}:
            return "#docx-page-count()"
        return f"#docx-field({typst_string(run.field or '')})"
    if run.kind == "footnote":
        text = ""
        if run.footnote is not None:
            text = "\n\n".join(_render_block_text(block) for block in run.footnote.blocks).strip()
        return f"#footnote[{text}]"
    if run.kind == "math":
        return f"#docx-math({typst_string(run.text)})"
    if not run.text:
        return ""
    content = _render_text(run.text)
    if run.code:
        content = f"#raw({typst_string(run.text)})"
    if run.style_name and run.style_name.lower() in {"strong", "emphasis"}:
        content = f"*{content}*" if run.style_name.lower() == "strong" else f"_{content}_"
    if run.superscript:
        content = f"#super[{content}]"
    if run.subscript:
        content = f"#sub[{content}]"
    if run.underline:
        content = f"#underline[{content}]"
    if run.strike:
        content = f"#strike[{content}]"
    if run.small_caps:
        content = f"#smallcaps[{content}]"
    if run.all_caps:
        content = f"#docx-allcaps[{content}]"
    text_args: list[str] = []
    if run.font:
        text_args.append(f"font: {typst_string(run.font)}")
    if run.size is not None:
        text_args.append(f"size: {format_length(run.size)}")
    if run.color:
        text_args.append(f"fill: rgb({typst_string(run.color)})")
    if text_args:
        content = f"#text({', '.join(text_args)})[{content}]"
    if run.highlight:
        content = f"#highlight(fill: rgb({typst_string(run.highlight)}))[{content}]"
    if run.bold and run.italic:
        content = f"*_{content}_*"
    elif run.bold:
        content = f"*{content}*"
    elif run.italic:
        content = f"_{content}_"
    if run.link:
        content = f"#link({typst_string(run.link)})[{content}]"
    if run.anchor:
        content = f"#docx-anchor-link({typst_string(run.anchor)})[{content}]"
    return content


def _render_text(text: str) -> str:
    lines = text.split("\n")
    rendered: list[str] = []
    for index, line in enumerate(lines):
        if index:
            rendered.append("#linebreak()")
        rendered.append(_render_spaces(line))
    return "".join(rendered)


def _render_spaces(line: str) -> str:
    if not line:
        return ""
    parts: list[str] = []
    position = 0
    for match in re.finditer(r" +", line):
        before = line[position : match.start()]
        if before:
            parts.append(typst_escape(before))
        spaces = match.group()
        if len(spaces) == 1 and match.start() > 0 and match.end() < len(line):
            parts.append(" ")
        else:
            amount = len(spaces) * 0.25
            parts.append(f"#h({format_number(amount)}em)")
        position = match.end()
    tail = line[position:]
    if tail:
        parts.append(typst_escape(tail))
    return "".join(parts)


def _render_inline_image(run: Run) -> str:
    image = run.image
    if image is None or not image.path:
        label = (image.alt if image is not None else None) or "image"
        return f"#docx-unsupported-image({typst_string(label)})"
    args = [typst_string(image.path)]
    if image.width_pt:
        args.append(f"width: {format_length(image.width_pt)}")
    if image.height_pt:
        args.append(f"height: {format_length(image.height_pt)}")
    if image.alt:
        args.append(f"alt: {typst_string(image.alt)}")
    return f"#docx-inline-image({', '.join(args)})"


# ---------------------------------------------------------------------------
# Block rendering
# ---------------------------------------------------------------------------


def render_block(block: Block, indent: int = 0) -> str:
    if isinstance(block, Paragraph):
        return _indent(render_paragraph(block), indent)
    if isinstance(block, TableBlock):
        return _indent(render_table(block), indent)
    if isinstance(block, ImageBlock):
        return _indent(render_image(block), indent)
    if isinstance(block, CodeBlock):
        args = [typst_string(block.text)]
        if block.language:
            args.append(f"language: {typst_string(block.language)}")
        if block.style:
            args.append(f"style: {typst_string(block.style)}")
        return _indent(f"#docx-code({', '.join(args)})", indent)
    if isinstance(block, PageBreak):
        return _indent("#pagebreak()", indent)
    if isinstance(block, HorizontalRule):
        return _indent("#line(length: 100%)", indent)
    if isinstance(block, UnsupportedBlock):
        return _indent(
            f"#docx-unsupported({typst_string(block.kind)}, detail: {typst_string(block.text or block.detail)})",
            indent,
        )
    if isinstance(block, SectionBreak):
        return _indent(render_section(block), indent)
    return ""


def render_paragraph(paragraph: Paragraph) -> str:
    body = render_runs(paragraph.runs)
    if paragraph.heading_level is not None:
        args = [
            "styles: docx-styles",
            f"level: {paragraph.heading_level}",
            f"style: {typst_string(paragraph.style or f'Heading {paragraph.heading_level}')}",
        ]
        args.extend(_paragraph_args(paragraph))
        return f"#docx-heading({', '.join(args)})[{body}]"
    if paragraph.list_kind:
        args = [
            "styles: docx-styles",
            f"kind: {typst_string(paragraph.list_kind)}",
            f"level: {paragraph.list_level}",
            f"style: {typst_string(paragraph.list_style or ('List Bullet' if paragraph.list_kind == 'bullet' else 'List Number'))}",
        ]
        if paragraph.marker is not None:
            args.append(f"marker: {typst_string(paragraph.marker)}")
        if paragraph.number_id is not None:
            args.append(f"num-id: {paragraph.number_id}")
        args.extend(_paragraph_args(paragraph, for_list=True))
        return f"#docx-item({', '.join(args)})[{body}]"
    args = ["styles: docx-styles", f"style: {typst_string(paragraph.style or 'Normal')}"]
    args.extend(_paragraph_args(paragraph))
    return f"#docx-para({', '.join(args)})[{body}]"


def _paragraph_args(paragraph: Paragraph, *, for_list: bool = False) -> list[str]:
    args: list[str] = []
    if paragraph.align and not for_list:
        args.append(f"align: {typst_string(paragraph.align)}")
    if paragraph.space_before is not None:
        args.append(f"before: {format_length(paragraph.space_before)}")
    if paragraph.space_after is not None:
        args.append(f"after: {format_length(paragraph.space_after)}")
    if paragraph.line_spacing is not None:
        args.append(f"line-spacing: {format_number(paragraph.line_spacing)}")
    if paragraph.line_spacing_pt is not None:
        args.append(f"line-spacing-pt: {format_length(paragraph.line_spacing_pt)}")
    if for_list:
        if paragraph.indent_left is not None:
            args.append(f"left-indent: {format_length(paragraph.indent_left)}")
        if paragraph.first_line_indent is not None:
            args.append(f"first-line-indent: {format_length(paragraph.first_line_indent)}")
        return args
    if paragraph.indent_left is not None:
        args.append(f"left-indent: {format_length(paragraph.indent_left)}")
    if paragraph.indent_right is not None:
        args.append(f"right-indent: {format_length(paragraph.indent_right)}")
    if paragraph.first_line_indent is not None:
        args.append(f"first-line-indent: {format_length(paragraph.first_line_indent)}")
    if paragraph.keep_with_next is not None:
        args.append(f"keep-with-next: {_bool(paragraph.keep_with_next)}")
    if paragraph.keep_lines is not None:
        args.append(f"keep-lines: {_bool(paragraph.keep_lines)}")
    if paragraph.page_break_before:
        args.append("page-break-before: true")
    if paragraph.bookmarks:
        names = ", ".join(typst_string(name) for name in paragraph.bookmarks)
        args.append(f"bookmarks: ({names},)")
    return args


def render_image(block: ImageBlock) -> str:
    image = block.image
    if not image.path:
        return f"#docx-unsupported-image({typst_string(image.alt or 'image')})"
    args = [typst_string(image.path)]
    if image.width_pt:
        args.append(f"width: {format_length(image.width_pt)}")
    if image.height_pt:
        args.append(f"height: {format_length(image.height_pt)}")
    if image.alt:
        args.append(f"alt: {typst_string(image.alt)}")
    if block.align and block.align != "center":
        args.append(f"align: {typst_string(block.align)}")
    if block.caption:
        caption = "\n\n".join(render_block(item).strip() for item in block.caption).strip()
        args.append(f"caption: [{caption}]")
    return f"#docx-image({', '.join(args)})"


def render_table(table: TableBlock) -> str:
    if not table.rows:
        return ""
    column_count = len(table.columns) or max(
        (cell.col + cell.colspan for row in table.rows for cell in row), default=1
    )
    column_count = max(1, column_count)
    lines = ["#docx-table("]
    lines.append(f"  columns: {column_count},")
    lines.append(f"  rows: {len(table.rows)},")
    if table.columns and any(width is not None for width in table.columns):
        widths = [format_length(width) if width is not None else "1fr" for width in table.columns]
        if len(widths) < column_count:
            widths.extend(["1fr"] * (column_count - len(widths)))
        lines.append(f"  widths: ({', '.join(widths)}),")
    if table.header_rows:
        lines.append(f"  header-rows: {table.header_rows},")
    merges = [
        f"({cell.row}, {cell.col}, {cell.rowspan}, {cell.colspan})"
        for row in table.rows
        for cell in row
        if cell.rowspan > 1 or cell.colspan > 1
    ]
    if merges:
        lines.append(f"  merges: ({', '.join(merges)},),")
    fills = [
        f"({cell.row}, {cell.col}, {typst_string(cell.fill)})"
        for row in table.rows
        for cell in row
        if cell.fill
    ]
    if fills:
        lines.append(f"  fills: ({', '.join(fills)},),")
    valigns = [
        f"({cell.row}, {cell.col}, {typst_string(cell.valign)})"
        for row in table.rows
        for cell in row
        if cell.valign
    ]
    if valigns:
        lines.append(f"  valigns: ({', '.join(valigns)},),")
    if table.style:
        lines.append(f"  style: {typst_string(table.style)},")
    if table.align:
        lines.append(f"  align: {typst_string(table.align)},")
    if table.stroke:
        lines.append(f"  stroke: {table.stroke},")
    lines.append("  cells: (")
    for row in table.rows:
        for cell in row:
            content = _render_cell(cell)
            lines.append(f"    [{content}],")
    lines.append("  ),")
    lines.append(")")
    return "\n".join(lines)


def _render_cell(cell: TableCell) -> str:
    rendered = [render_block(block).strip() for block in cell.blocks]
    return "\n\n".join(item for item in rendered if item)


def render_section(
    section: SectionBreak,
    header_blocks: list[Block] | None = None,
    footer_blocks: list[Block] | None = None,
) -> str:
    page = page_dict(section.page)
    inner = page[1:-1].strip() if page.startswith("(") and page.endswith(")") else page
    header_content = section.header if header_blocks is None else header_blocks
    footer_content = section.footer if footer_blocks is None else footer_blocks
    header = "\n\n".join(render_block(block).strip() for block in header_content).strip()
    footer = "\n\n".join(render_block(block).strip() for block in footer_content).strip()
    fields = [inner] if inner else []
    if header:
        fields.append(f"header: [\n{_indent(header, 2)}\n]")
    if footer:
        fields.append(f"footer: [\n{_indent(footer, 2)}\n]")
    return "#pagebreak()\n#set page(" + ", ".join(fields) + ")"


def render_header_footer(blocks: list[Block]) -> str:
    return "\n\n".join(render_block(block).strip() for block in blocks if render_block(block).strip()).strip()


# ---------------------------------------------------------------------------
# Document rendering
# ---------------------------------------------------------------------------


def render_document(document: Document) -> str:
    metadata = document.metadata
    lines: list[str] = [
        f"// Generated by docx-typst {__version__} from {document.source or 'a Word document'}.",
        "// Edit this file, then run: typst-to-docx main.typ --output document.docx",
        "",
        '#import "template.typ": *',
        "",
        render_styles(document),
        "",
        "#show: docx-document.with(",
    ]
    meta = []
    if metadata.title:
        meta.append(f"  title: {typst_string(metadata.title)},")
    if metadata.author:
        meta.append(f"  author: {typst_string(metadata.author)},")
    if metadata.subject:
        meta.append(f"  subject: {typst_string(metadata.subject)},")
    if metadata.keywords:
        meta.append(f"  keywords: {typst_string(metadata.keywords)},")
    lines.extend(meta)
    lines.append(f"  page: {page_dict(document.page)},")
    header = render_header_footer(document.header)
    footer = render_header_footer(document.footer)
    if header:
        lines.append("  header: [")
        lines.append(_indent(header, 4))
        lines.append("  ],")
    if footer:
        lines.append("  footer: [")
        lines.append(_indent(footer, 4))
        lines.append("  ],")
    lines.append(")")
    lines.append("")
    current_header = document.header
    current_footer = document.footer
    rendered_blocks: list[str] = []
    for block in document.blocks:
        if isinstance(block, SectionBreak):
            if block.header:
                current_header = block.header
            elif not block.header_linked:
                current_header = []
            if block.footer:
                current_footer = block.footer
            elif not block.footer_linked:
                current_footer = []
            rendered_blocks.append(
                render_section(block, current_header, current_footer)
            )
        else:
            rendered_blocks.append(render_block(block))
    body = "\n\n".join(item for item in rendered_blocks if item.strip())
    lines.append(body.strip())
    lines.append("")
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _render_block_text(block: Block) -> str:
    if isinstance(block, Paragraph):
        return render_runs(block.runs)
    if isinstance(block, CodeBlock):
        return typst_escape(block.text)
    return ""


def _indent(text: str, spaces: int) -> str:
    if not text:
        return text
    prefix = " " * spaces
    return "\n".join(prefix + line if line else line for line in text.split("\n"))


def _bool(value: bool) -> str:
    return "true" if value else "false"
