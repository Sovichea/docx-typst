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
        keep_rows: list[bool] = []
        for row_index, tr in enumerate(rows_raw):
            tr_pr = first(tr, "./w:trPr")
            if on_off(tr_pr, "w:tblHeader"):
                header_rows += 1
            keep_rows.append(bool(on_off(tr_pr, "w:cantSplit")))
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
                    borders=_table_border_map(tc_pr, "tcBorders"),
                    no_wrap=first(tc_pr, "./w:noWrap") is not None,
                    insets={edge: twips(attr(node, "w:w")) or 0
                            for edge in ("top", "bottom", "left", "right")
                            if (node := first(tc_pr, f"./w:tcMar/w:{edge}")) is not None},
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
        border_map = _table_border_map(tbl_pr)
        if style_name:
            try:
                style = self.docx.styles[style_name]
                while style is not None:
                    inherited = _table_border_map(first(style.element, "./w:tblPr"))
                    border_map = {**inherited, **border_map}
                    style = style.base_style
            except KeyError:
                pass
        stroke = _uniform_stroke(border_map)
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
            borders=border_map if len(set(border_map.values())) > 1 else {},
            keep_together=bool(keep_rows) and all(keep_rows),
            cell_margins_pt=twips(attr(first(tbl_pr, "./w:tblCellMar/w:left"), "w:w")),
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
    if value and value.lower() not in {"auto", "none"}:
        return "#" + value.upper()
    return None


def _table_border_map(tbl_pr: Any, tag: str = "tblBorders") -> dict[str, str]:
    borders = first(tbl_pr, "./w:" + tag)
    values = {}
    if borders is None:
        return values
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        node = first(borders, f"./w:{edge}")
        if node is None:
            continue
        if attr(node, "w:val", "single") in {"nil", "none"}:
            values[edge] = "none"
            continue
        try:
            width = float(attr(node, "w:sz", "4")) / 8
        except ValueError:
            width = 0.5
        color = attr(node, "w:color", "808080")
        if color == "auto":
            color = "808080"
        values[edge] = f'{_pt(width)} + rgb("#{color.upper()}")'
    return values


def _uniform_stroke(values: dict[str, str]) -> str | None:
    if not values:
        return None
    return next(iter(values.values())) if len(set(values.values())) == 1 else "none"


def _table_stroke(tbl_pr: Any) -> str | None:
    return _uniform_stroke(_table_border_map(tbl_pr))


def _pt(value: float) -> str:
    if abs(value - round(value)) < 1e-9:
        return f"{int(round(value))}pt"
    return f"{value:.3f}".rstrip("0").rstrip(".") + "pt"
