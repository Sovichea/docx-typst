"""Call-to-block conversion for the reverse Typst parser."""

from __future__ import annotations

from typing import Any

from .model import (
    CanvasBlock,
    PositionedBlock,
    Block,
    CodeBlock,
    HorizontalRule,
    DecorativeRule,
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
        if name == "docx-settings":
            self._apply_show_args(call.named)
            return None
        if name == "docx-money":
            return Paragraph(runs=[Run(text=string_value(call.arg("symbol", "$"))),
                                   Run(kind="tab"), Run(text=string_value(call.positional[0]))],
                             right_tab=True, align="left", space_after=0)
        if name in self._panel_helpers and call.body is not None:
            return self._panel_block(TypstCall(name="docx-panel", named=self._panel_helpers[name], body=call.body), location)
        if name in self._text_helpers and call.body is not None:
            settings = self._text_helpers[name]
            runs: list[Run] = []
            self.inline._call(TypstCall(name="text", named=settings, body=call.body), runs)
            size = length_to_pt(settings.get("size"), self.document.default_size)
            from .text_bounds import cap_height
            font = settings.get("font", self.document.default_font)
            style_name = name.replace("-", " ").title()
            spec = self.document.styles.get(style_name)
            return Paragraph(style=style_name if spec else None, runs=runs,
                             line_spacing_pt=spec.line_spacing_pt if spec and spec.line_spacing_pt is not None else self._text_line_height(font, size),
                             native_cap_height=cap_height(font, size), native_flow=True)
        if name in self._table_helpers and call.body is not None:
            previous = self._table_settings.copy()
            self._table_settings.update(self._table_helpers[name])
            blocks = self._parse_blocks(call.body.text, location)
            self._table_settings = previous
            return blocks
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
            return HorizontalRule(stroke=stroke_value(call.arg("stroke")) if "stroke" in call.named else None)
        if name == "docx-rule":
            color = color_to_rgb(call.arg("fill")) or (0, 0, 0)
            return DecorativeRule(word_after=length_to_pt(call.arg("word-after")) if "word-after" in call.named else None,
                                  before=length_to_pt(call.arg("before"), 0),
                                  word_before=length_to_pt(call.arg("word-before")) if "word-before" in call.named else None,
                                  width=length_to_pt(call.arg("width"), 38),
                                  height=length_to_pt(call.arg("height"), 3),
                                  fill="#" + "".join(f"{c:02X}" for c in color),
                                  after=length_to_pt(call.arg("after"), 10))
        if name == "table":
            return self._native_table(call, location)
        if name == "grid":
            return self._native_grid(call, location)
        if name == "figure":
            return self._figure(call, location)
        if name == "image":
            image = image_from_call(call)
            return ImageBlock(image=image, space_before=self._paragraph_spacing,
                              space_after=self._paragraph_spacing) if image is not None else None
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
            if parsed and name in {"block", "box"}:
                for key, block, attribute in (("above", parsed[0], "space_before"),
                                               ("below", parsed[-1], "space_after")):
                    if key in call.named and isinstance(block, (TableBlock, ImageBlock)):
                        setattr(block, attribute, length_to_pt(call.arg(key)))
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
        if name == "docx-space":
            value = call.arg("word", call.positional[0] if call.positional else UnitValue(0, "pt"))
            return Spacer(height_pt=length_to_pt(value))
        if name == "docx-pagebreak":
            return PageBreak()
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
            paragraph = Paragraph(style="Normal", runs=self.inline.parse("#" + call.raw), native_flow=True)
            if name == "text" and "size" in call.named:
                from .text_bounds import cap_height
                size = length_to_pt(call.arg("size"))
                font = call.arg("font") if isinstance(call.arg("font"), str) else self.document.default_font
                paragraph.line_spacing_pt = self._text_line_height(font, size)
                paragraph.native_cap_height = cap_height(font, size)
            return paragraph
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

    def _text_line_height(self, font: Any, size: float) -> float | None:
        from .text_bounds import cap_height
        if not isinstance(font, str):
            font = self.document.default_font
        cap = cap_height(font, size)
        leading = self._paragraph_leading
        gap = leading.value * size if isinstance(leading, UnitValue) and leading.unit == "em" else length_to_pt(leading)
        return cap + gap if cap is not None else None

    def _paragraph(self, call: TypstCall, location: str) -> Paragraph:
        body = call.body.text if call.body is not None else ""
        style = string_value(call.arg("style")) or "Normal"
        paragraph = Paragraph(style=style, runs=self.inline.parse(body))
        style_spec = self.document.styles.get(style)
        if style_spec is not None:
            paragraph.align = style_spec.align
        if "native-cap-height" in call.named:
            paragraph.native_cap_height = length_to_pt(call.arg("native-cap-height"))
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
        block = ImageBlock(image=image or ImageRef(path=""), space_before=self._paragraph_spacing,
                           space_after=self._paragraph_spacing)
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
            width = call.arg("native-width", call.arg("width"))
            height = call.arg("native-height", call.arg("height"))
            return ImageBlock(
                image=ImageRef(
                    path=path,
                    width_pt=length_to_pt(width) if width is not None else None,
                    height_pt=length_to_pt(height) if height is not None else None,
                    typst_source="#" + call.raw,
                ),
                space_before=self._paragraph_spacing, space_after=self._paragraph_spacing,
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
        previous_width = self._layout_width
        page = self._current_page
        self._layout_width = (previous_width or (page.width_pt - page.margin_left - page.margin_right)) - inset * 2
        blocks = self._parse_blocks(call.body.text, location)
        panel_width = self._layout_width + inset * 2
        self._layout_width = previous_width
        return TableBlock(
            rows=[[TableCell(blocks=blocks, fill=fill_hex)]],
            columns=[panel_width],
            style="Normal Table",
            stroke="none",
            cell_margins_pt=inset,
            keep_together=True,
            space_before=self._paragraph_spacing, space_after=self._paragraph_spacing,
            native_text_bounds=True,
            borders={edge: stroke_value(value) for edge, value in call.arg("stroke", {}).items()}
            if isinstance(call.arg("stroke"), dict) else {},
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
        column_weights: list[float] = []
        if isinstance(widths_value, list):
            column_weights = [item.value if isinstance(item, UnitValue) and item.unit == "fr" else 0
                              for item in widths_value]
            widths = [None if weight else length_to_pt(item) for item, weight in zip(widths_value, column_weights)]
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
        inset_map = {}
        for item in call.arg("cell-insets", []):
            if isinstance(item, list) and len(item) == 3 and isinstance(item[2], dict):
                inset_map[(int(item[0]), int(item[1]))] = {
                    edge: length_to_pt(value) for edge, value in item[2].items()
                    if edge in {"top", "bottom", "left", "right"}}
        fill_map: dict[tuple[int, int], str] = {}
        fills = call.arg("fills")
        if isinstance(fills, list):
            for item in fills:
                if isinstance(item, list) and len(item) >= 3:
                    color = color_to_rgb(item[2])
                    if color:
                        fill_map[(int(item[0]), int(item[1]))] = '#' + ''.join(f'{v:02X}' for v in color)
        border_map = {(int(r), int(c)): {edge: stroke_value(v) for edge, v in paints.items()}
                      for r, c, paints in call.arg("cell-borders", [])}
        align_map = {(int(r), int(c)): value for r, c, value in call.arg("cell-aligns", [])}
        nowrap = {tuple(item) for item in call.arg("cell-no-wrap", [])}
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
                self._table_paragraphs(blocks)
                for paragraph in blocks:
                    if isinstance(paragraph, Paragraph) and call.arg("text-style") in self.document.styles and paragraph.style in {None, "Normal"}:
                        paragraph.style = call.arg("text-style")
                    if isinstance(paragraph, Paragraph) and (r, c) in align_map:
                        paragraph.align = align_map[(r, c)]
                row.append(
                    TableCell(
                        blocks=blocks,
                        colspan=colspan,
                        rowspan=rowspan,
                        row=r,
                        col=c,
                        fill=fill_map.get((r, c)),
                        insets=inset_map.get((r, c), {}),
                        borders=border_map.get((r, c), {}),
                        no_wrap=(r, c) in nowrap,
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
            column_weights=column_weights,
            native_text_bounds=bool(call.arg("native-text-bounds", True)),
            borders={edge: stroke_value(value) for edge, value in call.arg("borders", {}).items()},
        )

    def _native_table(self, call: TypstCall, location: str) -> TableBlock | None:
        settings = {**self._table_settings, **call.named}
        columns_value = call.arg("columns", 1)
        ncols = len(columns_value) if isinstance(columns_value, list) else int(
            columns_value or 1
        )
        gutter = length_to_pt(call.arg("column-gutter"), 0.0)
        page = self._current_page
        available = self._layout_width or (page.width_pt - page.margin_left - page.margin_right)
        widths = _relative_widths(columns_value, total=available, gutter=gutter)
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
                self._table_paragraphs(blocks)
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
                          stroke=stroke_value(settings.get("stroke", "1pt + black")),
                          native_text_bounds=True,
                          space_before=self._paragraph_spacing, space_after=self._paragraph_spacing)

    def _table_paragraphs(self, blocks: list[Block]) -> None:
        justify = self._table_par_settings.get("justify")
        if justify is None:
            return
        for block in blocks:
            if isinstance(block, Paragraph) and block.align is None:
                block.align = "justify" if justify else "left"

    def _native_grid(self, call: TypstCall, location: str) -> TableBlock | None:
        columns_value = call.arg("columns", 1)
        ncols = len(columns_value) if isinstance(columns_value, list) else int(
            columns_value or 1
        )
        gutter = length_to_pt(call.arg("column-gutter"), 0.0)
        page = self._current_page
        available = self._layout_width or (page.width_pt - page.margin_left - page.margin_right)
        widths = _relative_widths(columns_value, total=available, gutter=gutter)
        if widths is None:
            widths = [(available - gutter * (ncols - 1)) / ncols] * ncols
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
                previous_width = self._layout_width
                self._layout_width = widths[c]
                blocks = self._parse_blocks(cells[index].text, f"{location}:cell:{r}:{c}")
                self._layout_width = previous_width
                if blocks and isinstance(blocks[-1], Paragraph) and blocks[-1].native_flow:
                    from .text_bounds import cap_height
                    last = blocks[-1]
                    spec = self.document.styles.get(last.style or "Normal")
                    size = spec.size if spec and spec.size else self.document.default_size
                    font = spec.font if spec and spec.font else self.document.default_font
                    cap = cap_height(font, size)
                    line = last.line_spacing_pt or (spec.line_spacing_pt if spec else None)
                    padding = max(0, line - cap) if line and cap is not None else 0
                    row_gap = length_to_pt(call.arg("row-gutter"), 0)
                    last.space_after = max(0, row_gap - padding) if (r + 1) * ncols < len(cells) else 0
                    last.native_flow = False
                row.append(
                    TableCell(
                        blocks=blocks,
                        row=r,
                        col=c * 2 if gutter else c,
                    )
                )
                if gutter and c < ncols - 1:
                    row.append(TableCell(row=r, col=c * 2 + 1))
            rows.append(row)
        return TableBlock(
            rows=rows,
            columns=[value for index, width in enumerate(widths)
                     for value in ([width, gutter] if gutter and index < ncols - 1 else [width])],
            style="Normal Table",
            stroke="none",
            cell_margins_pt=0.0,
            space_before=self._paragraph_spacing, space_after=self._paragraph_spacing,
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
                    block.list_restart = not blocks
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
