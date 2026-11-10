"""Table reading for the forward converter."""

from __future__ import annotations

from typing import Any

from .model import Block, TableBlock, TableCell
from .ooxml import attr, first, on_off, twips, xpath


class TableMixin:
    def _read_table(self, table: Any, part: Any) -> Block:
        tbl_pr = first(table, "./w:tblPr")
        style_id = attr(first(tbl_pr, "./w:tblStyle"), "w:val")
        style_name = self.styles.style_name(style_id)
        tbl_grid = first(table, "./w:tblGrid")
        columns: list[float | None] = []
        if tbl_grid is not None:
            for grid_col in xpath(tbl_grid, "./w:gridCol"):
                columns.append(twips(attr(grid_col, "w:w")))
        rows_raw = xpath(table, "./w:tr")
        grid_cells: dict[tuple[int, int], TableCell] = {}
        occupied: set[tuple[int, int]] = set()
        row_cells: list[list[TableCell]] = []
        header_rows = 0
        for row_index, tr in enumerate(rows_raw):
            tr_pr = first(tr, "./w:trPr")
            if on_off(tr_pr, "w:tblHeader"):
                header_rows += 1
            cells: list[TableCell] = []
            col_index = 0
            for tc in xpath(tr, "./w:tc"):
                while (row_index, col_index) in occupied:
                    col_index += 1
                tc_pr = first(tc, "./w:tcPr")
                span = int(attr(first(tc_pr, "./w:gridSpan"), "w:val", "1") or 1)
                vmerge_node = first(tc_pr, "./w:vMerge")
                vmerge = attr(vmerge_node, "w:val") if vmerge_node is not None else None
                if vmerge == "continue":
                    top = None
                    for probe in range(row_index - 1, -1, -1):
                        candidate = grid_cells.get((probe, col_index))
                        if candidate is not None:
                            top = candidate
                            break
                    if top is not None:
                        top.rowspan += 1
                    for offset in range(span):
                        occupied.add((row_index, col_index + offset))
                    col_index += span
                    continue
                width = twips(attr(first(tc_pr, "./w:tcW"), "w:w"))
                fill = _cell_fill(tc_pr, self.theme)
                valign = attr(first(tc_pr, "./w:vAlign"), "w:val")
                cell_blocks = self._read_container(tc, part, section_break=False)
                cell = TableCell(
                    blocks=cell_blocks,
                    colspan=span,
                    rowspan=1,
                    fill=fill,
                    valign=valign,
                    width_pt=width,
                    row=row_index,
                    col=col_index,
                )
                cells.append(cell)
                grid_cells[(row_index, col_index)] = cell
                for offset in range(span):
                    occupied.add((row_index, col_index + offset))
                col_index += span
            row_cells.append(cells)
        if not columns:
            max_cols = max(
                (cell.col + cell.colspan for row in row_cells for cell in row), default=1
            )
            columns = [None] * max_cols
        stroke = _table_stroke(tbl_pr)
        if stroke is None:
            self.warn(
                "table_borders",
                "approximated",
                f"table:{self.counts['tables']}",
                source_type="table",
                fallback="default Typst table stroke",
                reason="complex or missing border definitions were not resolved",
            )
        self.counts["tables"] += 1
        self.counts["table_rows"] += len(row_cells)
        return TableBlock(
            rows=row_cells,
            columns=columns,
            header_rows=min(header_rows, len(row_cells)),
            style=style_name,
            align=attr(first(tbl_pr, "./w:jc"), "w:val"),
            stroke=stroke,
        )


def _cell_fill(tc_pr: Any, theme: Any) -> str | None:
    shd = first(tc_pr, "./w:shd")
    if shd is None:
        return None
    value = attr(shd, "w:fill")
    theme_name = attr(shd, "w:themeFill")
    if theme_name:
        return theme.color(
            theme_name,
            attr(shd, "w:themeFillTint"),
            attr(shd, "w:themeFillShade"),
        )
    if value and value.lower() not in {"auto", "none", "ffffff"}:
        return "#" + value.upper()
    return None


def _table_stroke(tbl_pr: Any) -> str | None:
    borders = first(tbl_pr, "./w:tblBorders")
    if borders is None:
        return None
    values = {}
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        node = first(borders, f"./w:{edge}")
        if node is None:
            continue
        val = attr(node, "w:val", "single")
        if val in {"nil", "none"}:
            continue
        size = attr(node, "w:sz")
        color = attr(node, "w:color")
        try:
            width = max(0.25, float(size or 4) / 8.0)
        except ValueError:
            width = 0.5
        values[edge] = (width, color)
    if not values:
        return None
    widths = {round(width, 2) for width, _ in values.values()}
    colors = {color for _, color in values.values() if color and color.lower() != "auto"}
    if len(widths) == 1 and len(colors) <= 1:
        width = next(iter(widths))
        color = next(iter(colors), "808080")
        return f'{_pt(width)} + rgb("#{color.upper()}")'
    return None


def _pt(value: float) -> str:
    if abs(value - round(value)) < 1e-9:
        return f"{int(round(value))}pt"
    return f"{value:.3f}".rstrip("0").rstrip(".") + "pt"
