"""Call-to-block conversion for the reverse Typst parser."""

from __future__ import annotations

from typing import Any

from .model import (
    CanvasBlock,
    PositionedBlock,
    Block,
    CodeBlock,
    HorizontalRule,
    ImageBlock,
    ImageRef,
    PageBreak,
    Paragraph,
    Run,
    Spacer,
    TableBlock,
    TableCell,
    UnsupportedBlock,
)
from .typst_syntax import Content, TypstCall, UnitValue, color_to_rgb, length_to_pt
from .typst_values import image_from_call, string_value, stroke_value


def _relative_widths(
    columns_value: Any, total: float = 465.0, gutter: float = 0.0
) -> list[float] | None:
    if not isinstance(columns_value, list):
        return None
    values: list[tuple[str, float]] = []
    fixed = 0.0
    weights = 0.0
    for item in columns_value:
        if isinstance(item, UnitValue) and item.unit == "fr":
            values.append(("fr", item.value))
            weights += item.value
        else:
            width = length_to_pt(item, 0.0)
            values.append(("pt", width))
            fixed += width
    remaining = max(0.0, total - fixed - gutter * max(0, len(values) - 1))
    result: list[float] = []
    for kind, value in values:
        if kind == "pt":
            result.append(value)
        else:
            result.append(remaining * value / weights if weights else remaining)
    return result


class BlockBuilderMixin:
    def _call_block(self, call: TypstCall, location: str) -> Block | list[Block] | None:
        name = call.name
        if name == "docx-canvas" and call.body is not None:
            return CanvasBlock(
                blocks=self._parse_blocks(call.body.text, location + ":canvas"),
                width=length_to_pt(call.arg("width"), 595.2756),
                height=length_to_pt(call.arg("height"), 841.8898),
                identifier=string_value(call.arg("id")) or "canvas",
                source="#" + call.raw,
            )
        if name == "docx-position" and call.body is not None:
            return PositionedBlock(
                blocks=self._parse_blocks(call.body.text, location + ":position"),
                x=length_to_pt(call.arg("dx")), y=length_to_pt(call.arg("dy")),
                width=length_to_pt(call.arg("width"), 100.0),
                height=length_to_pt(call.arg("height"), 100.0),
                style=string_value(call.arg("style")),
                align=string_value(call.arg("align")) or "left",
                source="#" + call.raw,
            )
        if name == "docx-para":
            return self._paragraph(call, location)
        if name == "docx-heading":
            level = int(call.arg("level", 1) or 1)
            paragraph = self._paragraph(call, location)
            paragraph.heading_level = level
            paragraph.style = string_value(call.arg("style")) or f"Heading {level}"
            paragraph.outline_level = level - 1
            return paragraph
        if name == "docx-item":
            paragraph = self._paragraph(call, location)
            paragraph.list_kind = string_value(call.arg("kind")) or "bullet"
            paragraph.list_level = int(call.arg("level", 0) or 0)
            paragraph.list_style = string_value(call.arg("style")) or (
                "List Bullet" if paragraph.list_kind == "bullet" else "List Number"
            )
            paragraph.style = paragraph.list_style
            paragraph.marker = string_value(call.arg("marker"))
            number_id = call.arg("num-id")
            paragraph.number_id = number_id if isinstance(number_id, int) else None
            return paragraph
        if name == "docx-image":
            return self._image_block(call, location)
        if name == "docx-inline-image":
            image = image_from_call(call)
            if image is not None:
                return ImageBlock(image=image, align=image.align or "left")
            return None
        if name == "docx-static":
            return self._static_block(call, location)
        if name == "docx-panel":
            return self._panel_block(call, location)
        if name == "docx-table":
            return self._table(call, location)
        if name == "docx-code":
            text = call.positional[0] if call.positional else ""
            return CodeBlock(
                text=string_value(text) or "",
                language=string_value(call.arg("language")),
            )
        if name == "docx-unsupported":
            kind = string_value(call.positional[0] if call.positional else "content") or "content"
            return UnsupportedBlock(kind=kind, text=string_value(call.arg("detail")) or "")
        if name == "pagebreak":
            return PageBreak()
        if name in {"v", "vertical-space"}:
            value = call.positional[0] if call.positional else 0
            height = length_to_pt(value)
            if height <= 0 or location.startswith("header") or location.startswith("footer"):
                return None
            return Spacer(height_pt=height)
        if name in {"line", "hline"}:
            return HorizontalRule()
        if name == "table":
            return self._native_table(call, location)
        if name == "grid":
            return self._native_grid(call, location)
        if name == "figure":
            return self._figure(call, location)
        if name == "image":
            image = image_from_call(call)
            return ImageBlock(image=image) if image is not None else None
        if name in {"block", "align", "pad", "box", "place", "quote"}:
            if call.body is None:
                return None
            parsed = self._parse_blocks(call.body.text, location)
            if name == "align" and call.positional:
                alignment = string_value(call.positional[0])
                if alignment in {"left", "center", "right"}:
                    for block in parsed:
                        if isinstance(block, (TableBlock, ImageBlock, Paragraph)):
                            block.align = alignment
            if name == "quote":
                for block in parsed:
                    if isinstance(block, Paragraph):
                        block.quoted = True
                        block.style = block.style or "Quote"
            return parsed
        if name in {"list", "enum"}:
            return self._native_list(
                call, location, "bullet" if name == "list" else "number"
            )
        if name in {"docx-page-number", "docx-page-count", "docx-field"}:
            field = (
                "PAGE"
                if name == "docx-page-number"
                else "NUMPAGES"
                if name == "docx-page-count"
                else string_value(call.positional[0] if call.positional else "")
            )
            return Paragraph(style="Normal", runs=[Run(kind="field", field=field)])
        if name in {"docx-page-number-padded", "docx-page-count-padded"}:
            field = (
                "PAGE_PADDED"
                if name == "docx-page-number-padded"
                else "NUMPAGES_PADDED"
            )
            return Paragraph(style="Normal", runs=[Run(kind="field", field=field)])
        if name == "docx-unsupported-image":
            return UnsupportedBlock(
                kind="image", text=string_value(call.arg("alt")) or ""
            )
        if name == "docx-page-start":
            value = call.positional[0] if call.positional else 1
            try:
                self._pending_page_number_start = int(value)
            except (TypeError, ValueError):
                self._pending_page_number_start = 1
            return None
        if name == "docx-anchor":
            return None
        if name in {
            "text",
            "link",
            "strong",
            "emph",
            "underline",
            "strike",
            "smallcaps",
            "sub",
            "super",
            "upper",
            "highlight",
            "raw",
        }:
            return Paragraph(style="Normal", runs=self.inline.parse("#" + call.raw))
        if call.body is not None:
            self.document.warn(
                "typst_function",
                "approximated",
                location,
                source_type=name,
                fallback="function body flattened into paragraphs",
                reason="the static subset does not evaluate user-defined functions",
            )
            return self._parse_blocks(call.body.text, location)
        return None

    def _paragraph(self, call: TypstCall, location: str) -> Paragraph:
        body = call.body.text if call.body is not None else ""
        style = string_value(call.arg("style")) or "Normal"
        paragraph = Paragraph(style=style, runs=self.inline.parse(body))
        align = string_value(call.arg("align"))
        if align:
            paragraph.align = align
        for key, attr_name in (
            ("before", "space_before"),
            ("after", "space_after"),
            ("left-indent", "indent_left"),
            ("right-indent", "indent_right"),
            ("first-line-indent", "first_line_indent"),
            ("line-spacing-pt", "line_spacing_pt"),
        ):
            value = call.arg(key)
            if value is not None:
                setattr(paragraph, attr_name, length_to_pt(value))
        value = call.arg("line-spacing")
        if value is not None:
            paragraph.line_spacing = float(value)
        for key, attr_name in (
            ("keep-with-next", "keep_with_next"),
            ("keep-lines", "keep_lines"),
            ("page-break-before", "page_break_before"),
        ):
            value = call.arg(key)
            if isinstance(value, bool):
                setattr(paragraph, attr_name, value)
        bookmarks = call.arg("bookmarks")
        if isinstance(bookmarks, list):
            paragraph.bookmarks = [
                str(item) for item in bookmarks if isinstance(item, str)
            ]
        return paragraph

    def _image_block(self, call: TypstCall, location: str) -> ImageBlock:
        image = image_from_call(call)
        block = ImageBlock(image=image or ImageRef(path=""))
        align = string_value(call.arg("align"))
        if align:
            block.align = align
        caption = call.arg("caption")
        if isinstance(caption, Content):
            block.caption = self._parse_blocks(caption.text, location + ":caption")
        return block

    def _static_block(self, call: TypstCall, location: str) -> Block | list[Block] | None:
        path = string_value(call.arg("path"))
        if path:
            width = call.arg("width")
            height = call.arg("height")
            return ImageBlock(
                image=ImageRef(
                    path=path,
                    width_pt=length_to_pt(width) if width is not None else None,
                    height_pt=length_to_pt(height) if height is not None else None,
                    typst_source="#" + call.raw,
                )
            )
        if call.body is not None:
            return self._parse_blocks(call.body.text, location)
        return None

    def _panel_block(self, call: TypstCall, location: str) -> TableBlock | None:
        if call.body is None:
            return None
        fill = call.arg("fill")
        color = color_to_rgb(fill) if fill is not None else (255, 249, 231)
        fill_hex = "#" + "".join(f"{value:02X}" for value in color) if color else "#FFF9E7"
        inset = length_to_pt(call.arg("inset"), 10.0)
        blocks = self._parse_blocks(call.body.text, location)
        return TableBlock(
            rows=[[TableCell(blocks=blocks, fill=fill_hex)]],
            columns=[465.0],
            style="Normal Table",
            stroke="none",
            cell_margins_pt=inset,
        )

    def _table(self, call: TypstCall, location: str) -> TableBlock | None:
        columns_value = call.arg("columns", 1)
        if isinstance(columns_value, list):
            ncols = len(columns_value)
            widths = [length_to_pt(item) for item in columns_value]
        else:
            ncols = int(columns_value or 1)
            widths = []
        widths_value = call.arg("widths")
        if isinstance(widths_value, list):
            widths = [length_to_pt(item) for item in widths_value]
        cells_value = call.arg("cells")
        cells: list[Content] = []
        if isinstance(cells_value, list):
            cells = [item for item in cells_value if isinstance(item, Content)]
        if not cells:
            cells = [item for item in call.positional if isinstance(item, Content)]
        if not cells:
            return None
        merge_map: dict[tuple[int, int], tuple[int, int]] = {}
        covered: set[tuple[int, int]] = set()
        merges = call.arg("merges")
        if isinstance(merges, list):
            for merge in merges:
                if isinstance(merge, list) and len(merge) >= 4:
                    row, col, rowspan, colspan = (int(item) for item in merge[:4])
                    merge_map[(row, col)] = (rowspan, colspan)
                    for r in range(row, row + rowspan):
                        for c in range(col, col + colspan):
                            if (r, c) != (row, col):
                                covered.add((r, c))
        fill_map: dict[tuple[int, int], str] = {}
        fills = call.arg("fills")
        if isinstance(fills, list):
            for item in fills:
                if isinstance(item, list) and len(item) >= 3 and isinstance(item[2], str):
                    fill_map[(int(item[0]), int(item[1]))] = item[2]
        valign_map: dict[tuple[int, int], str] = {}
        valigns = call.arg("valigns")
        if isinstance(valigns, list):
            for item in valigns:
                if isinstance(item, list) and len(item) >= 3 and isinstance(item[2], str):
                    valign_map[(int(item[0]), int(item[1]))] = item[2]
        header_rows = int(call.arg("header-rows", 0) or 0)
        inset = length_to_pt(call.arg("inset"), 5.0)
        rows_value = call.arg("rows")
        if isinstance(rows_value, int) and rows_value > 0:
            nrows = rows_value
        else:
            nrows = (len(cells) + ncols - 1) // ncols
            for (row, _col), (rowspan, _colspan) in merge_map.items():
                nrows = max(nrows, row + rowspan)
        rows: list[list[TableCell]] = []
        index = 0
        for r in range(nrows):
            row: list[TableCell] = []
            for c in range(ncols):
                if (r, c) in covered:
                    continue
                if index >= len(cells):
                    break
                rowspan, colspan = merge_map.get((r, c), (1, 1))
                blocks = self._parse_blocks(
                    cells[index].text, f"{location}:cell:{r}:{c}"
                )
                row.append(
                    TableCell(
                        blocks=blocks,
                        colspan=colspan,
                        rowspan=rowspan,
                        row=r,
                        col=c,
                        fill=fill_map.get((r, c)),
                        valign=valign_map.get((r, c)),
                    )
                )
                index += 1
            rows.append(row)
        return TableBlock(
            rows=rows,
            columns=widths or [None] * ncols,
            header_rows=header_rows,
            style=string_value(call.arg("style")),
            align=string_value(call.arg("align")),
            stroke=stroke_value(call.arg("stroke")) if "stroke" in call.named else '0.5pt + luma(180)',
            keep_together=bool(call.arg("keep-together", False)),
            cell_margins_pt=inset,
        )

    def _native_table(self, call: TypstCall, location: str) -> TableBlock | None:
        settings = {**self._table_settings, **call.named}
        columns_value = call.arg("columns", 1)
        ncols = len(columns_value) if isinstance(columns_value, list) else int(
            columns_value or 1
        )
        gutter = length_to_pt(call.arg("column-gutter"), 0.0)
        widths = _relative_widths(columns_value, gutter=gutter)
        cells = [item for item in call.positional if isinstance(item, Content)]
        if not cells:
            return None
        rows: list[list[TableCell]] = []
        for r in range((len(cells) + ncols - 1) // ncols):
            row: list[TableCell] = []
            for c in range(ncols):
                index = r * ncols + c
                if index >= len(cells):
                    break
                blocks = self._parse_blocks(
                    cells[index].text, f"{location}:cell:{r}:{c}"
                )
                for block in blocks:
                    if isinstance(block, Paragraph) and block.style in {None, "Normal"}:
                        block.style = "Table Text"
                row.append(
                    TableCell(
                        blocks=blocks,
                        row=r,
                        col=c,
                        fill="#FFF9E7" if r == 0 else None,
                    )
                )
            rows.append(row)
        return TableBlock(rows=rows, columns=widths or [None] * ncols,
                          cell_margins_pt=length_to_pt(settings.get("inset"), 5.0),
                          stroke=stroke_value(settings.get("stroke", "1pt + black")))

    def _native_grid(self, call: TypstCall, location: str) -> TableBlock | None:
        columns_value = call.arg("columns", 1)
        ncols = len(columns_value) if isinstance(columns_value, list) else int(
            columns_value or 1
        )
        gutter = length_to_pt(call.arg("column-gutter"), 0.0)
        widths = _relative_widths(columns_value, gutter=gutter)
        cells = [item for item in call.positional if isinstance(item, Content)]
        if not cells:
            return None
        rows: list[list[TableCell]] = []
        for r in range((len(cells) + ncols - 1) // ncols):
            row: list[TableCell] = []
            for c in range(ncols):
                index = r * ncols + c
                if index >= len(cells):
                    break
                row.append(
                    TableCell(
                        blocks=self._parse_blocks(
                            cells[index].text, f"{location}:cell:{r}:{c}"
                        ),
                        row=r,
                        col=c,
                    )
                )
            rows.append(row)
        return TableBlock(
            rows=rows,
            columns=widths or [None] * ncols,
            style="Normal Table",
            stroke="none",
            cell_margins_pt=0.0,
        )

    def _native_list(
        self, call: TypstCall, location: str, kind: str
    ) -> list[Block]:
        blocks: list[Block] = []
        for item in call.positional:
            if not isinstance(item, Content):
                continue
            for block in self._parse_blocks(item.text, location):
                if isinstance(block, Paragraph):
                    block.list_kind = kind
                    block.list_style = (
                        "List Bullet" if kind == "bullet" else "List Number"
                    )
                    block.style = block.list_style
                blocks.append(block)
        return blocks

    def _figure(self, call: TypstCall, location: str) -> ImageBlock | None:
        image: ImageRef | None = None
        for item in call.positional:
            if isinstance(item, TypstCall) and item.name == "image":
                image = image_from_call(item)
                break
        if image is None:
            return None
        block = ImageBlock(image=image)
        caption = call.arg("caption")
        if isinstance(caption, Content):
            block.caption = self._parse_blocks(caption.text, location + ":caption")
        return block
