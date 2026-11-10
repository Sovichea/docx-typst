"""Table rendering for the reverse converter."""

from __future__ import annotations

from typing import Any

from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt

from .model import Block, TableBlock
from .typst_syntax import UnitValue, color_to_rgb, parse_value, split_top_level


TABLE_ALIGN_MAP = {
    "left": WD_TABLE_ALIGNMENT.LEFT,
    "center": WD_TABLE_ALIGNMENT.CENTER,
    "right": WD_TABLE_ALIGNMENT.RIGHT,
}

VALIGN_MAP = {
    "top": WD_CELL_VERTICAL_ALIGNMENT.TOP,
    "center": WD_CELL_VERTICAL_ALIGNMENT.CENTER,
    "bottom": WD_CELL_VERTICAL_ALIGNMENT.BOTTOM,
}


class TableWriterMixin:
    def _table(self, container: Any, block: TableBlock) -> None:
        if not block.rows:
            return
        ncols = len(block.columns) or max(
            (cell.col + cell.colspan for row in block.rows for cell in row), default=1
        )
        ncols = max(1, ncols)
        nrows = len(block.rows)
        table = self._add_table(container, nrows, ncols)
        self._table_gap(container, table, block.space_before)
        if block.cell_margins_pt is not None:
            tbl_pr = table._tbl.tblPr
            margins = OxmlElement("w:tblCellMar")
            for edge in ("top", "left", "bottom", "right"):
                node = OxmlElement("w:" + edge)
                node.set(qn("w:w"), str(round(block.cell_margins_pt * 20)))
                node.set(qn("w:type"), "dxa")
                margins.append(node)
            look = tbl_pr.find(qn("w:tblLook"))
            if look is not None:
                look.addprevious(margins)
            else:
                tbl_pr.append(margins)
        style_name = block.style or "Table Grid"
        try:
            table.style = self.doc.styles[style_name]
        except KeyError:
            if block.style:
                self.model.warn(
                    "table_style",
                    "approximated",
                    "output",
                    source_type="table style",
                    fallback="default table formatting",
                    reason=f"style {block.style!r} is not available in the target document",
                )
        if block.align:
            table.alignment = TABLE_ALIGN_MAP.get(block.align)
        if block.align in {None, "left"}:
            # Word aligns left tables by their first cell's text edge. Typst
            # aligns the table frame; compensate for Word's cell inset.
            indent = OxmlElement("w:tblInd")
            indent.set(qn("w:type"), "dxa")
            indent.set(qn("w:w"), str(round((block.cell_margins_pt if block.cell_margins_pt is not None else 5.4) * 20)))
            table._tbl.tblPr.insert_element_before(indent, "w:tblBorders", "w:shd", "w:tblLayout", "w:tblCellMar", "w:tblLook")
        if block.stroke is not None:
            self._table_borders(table, block.stroke, block.borders)
        widths = list(block.columns)
        declared = sum(width or 0 for width in widths)
        available = self._available_table_width(container)
        if any(block.column_weights):
            remaining = max(0, available - declared)
            total_weight = sum(block.column_weights)
            widths = [remaining * weight / total_weight if weight else width
                      for width, weight in zip(widths, block.column_weights)]
            declared = sum(width or 0 for width in widths)
        if declared > available > 0:
            factor = available / declared
            widths = [width * factor if width else None for width in widths]
        table.autofit = block.autofit and not any(widths)
        for index, width in enumerate(widths):
            if width:
                table.columns[index].width = Pt(width)
        if any(widths):
            total = sum(width or 0 for width in widths)
            tbl_pr = table._tbl.tblPr
            tbl_w = tbl_pr.find(qn("w:tblW"))
            if tbl_w is None:
                tbl_w = OxmlElement("w:tblW")
                tbl_pr.append(tbl_w)
            tbl_w.set(qn("w:type"), "dxa")
            tbl_w.set(qn("w:w"), str(round(total * 20)))
        for row in block.rows:
            for cell in row:
                if any(widths):
                    width = sum(widths[cell.col : cell.col + cell.colspan])
                    if width:
                        table.cell(cell.row, cell.col).width = Pt(width)
                if cell.rowspan > 1 or cell.colspan > 1:
                    start = table.cell(cell.row, cell.col)
                    end = table.cell(
                        min(nrows - 1, cell.row + cell.rowspan - 1),
                        min(ncols - 1, cell.col + cell.colspan - 1),
                    )
                    start.merge(end)
        for row_index in range(min(block.header_rows, nrows)):
            tr_pr = table.rows[row_index]._tr.get_or_add_trPr()
            header = OxmlElement("w:tblHeader")
            header.set(qn("w:val"), "true")
            tr_pr.append(header)
        if block.keep_together:
            for row in table.rows:
                tr_pr = row._tr.get_or_add_trPr()
                if tr_pr.find(qn("w:cantSplit")) is not None:
                    continue
                cant_split = OxmlElement("w:cantSplit")
                try:
                    tr_pr.insert_element_before(
                        cant_split,
                        "w:trHeight",
                        "w:tblHeader",
                        "w:tblCellSpacing",
                        "w:jc",
                        "w:hidden",
                    )
                except (AttributeError, ValueError, KeyError):
                    tr_pr.append(cant_split)
        for row in block.rows:
            for cell in row:
                target = table.cell(cell.row, cell.col)
                if cell.fill:
                    self._cell_fill(target, cell.fill)
                if cell.valign:
                    target.vertical_alignment = VALIGN_MAP.get(
                        cell.valign, WD_CELL_VERTICAL_ALIGNMENT.TOP
                    )
                self._clear_cell(target)
                previous_inset = getattr(self, "_cell_inset", 0.0)
                self._cell_inset = block.cell_margins_pt or 0.0
                self._pending_space = self._pending_block_space = 0.0
                self._blocks(target, cell.blocks)
                if self._pending_space:
                    tail = target.add_paragraph()
                    tail.paragraph_format.space_before = tail.paragraph_format.space_after = Pt(0)
                    tail.paragraph_format.line_spacing = Pt(self._pending_space)
                self._cell_inset = previous_inset
                if not cell.blocks:
                    target.add_paragraph()
                if block.native_text_bounds and block.cell_margins_pt is not None:
                    from .text_bounds import native_cell_padding
                    native_cell_padding(target, block.cell_margins_pt)
        self._pending_space = 0.0
        self._pending_block_space = block.space_after or 0.0
        self.counts["tables"] = self.counts.get("tables", 0) + 1
        self.counts["table_rows"] = self.counts.get("table_rows", 0) + nrows

    def _table_borders(self, table: Any, stroke: str, overrides: dict[str, str] | None = None) -> None:
        properties = table._tbl.tblPr
        for existing in properties.findall(qn("w:tblBorders")):
            properties.remove(existing)
        borders = OxmlElement("w:tblBorders")
        for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
            edge_stroke = (overrides or {}).get(edge, stroke)
            width, color = 1.0, (0, 0, 0)
            for part in split_top_level(edge_stroke, "+"):
                value = parse_value(part)
                if isinstance(value, UnitValue):
                    width = value.to_pt()
                elif color_to_rgb(value) is not None:
                    color = color_to_rgb(value)
            border = OxmlElement("w:" + edge)
            border.set(qn("w:val"), "nil" if edge_stroke == "none" or width == 0 else "single")
            border.set(qn("w:sz"), str(round(width * 8)))
            border.set(qn("w:color"), "".join(f"{channel:02X}" for channel in color))
            borders.append(border)
        properties.insert_element_before(borders, "w:shd", "w:tblLayout", "w:tblCellMar", "w:tblLook")

    def _available_table_width(self, container: Any) -> float:
        if not hasattr(container, "_tc"):
            return self._content_width().pt
        width = getattr(container, "width", None)
        if width is None:
            return self._content_width().pt
        return max(18.0, width.pt - 2 * getattr(self, "_cell_inset", 0.0))

    def _add_table(self, container: Any, rows: int, cols: int) -> Any:
        try:
            return container.add_table(rows=rows, cols=cols)
        except TypeError:
            return container.add_table(rows, cols, self._content_width())

    def _clear_cell(self, cell: Any) -> None:
        for child in list(cell._tc):
            if child.tag in {qn("w:p"), qn("w:tbl")}:
                cell._tc.remove(child)

    def _cell_fill(self, cell: Any, color: str) -> None:
        tc_pr = cell._tc.get_or_add_tcPr()
        for existing in tc_pr.findall(qn("w:shd")):
            tc_pr.remove(existing)
        shading = OxmlElement("w:shd")
        shading.set(qn("w:val"), "clear")
        shading.set(qn("w:color"), "auto")
        shading.set(qn("w:fill"), color.lstrip("#").upper())
        anchor = None
        for tag in ("w:noWrap", "w:tcMar", "w:textDirection", "w:tcFitText", "w:vAlign", "w:hideMark"):
            anchor = tc_pr.find(qn(tag))
            if anchor is not None:
                break
        if anchor is not None:
            anchor.addprevious(shading)
        else:
            tc_pr.append(shading)
