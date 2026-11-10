"""Table rendering for the reverse converter."""

from __future__ import annotations

from typing import Any

from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt

from .model import Block, TableBlock


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
        table.autofit = block.autofit and not any(block.columns)
        for index, width in enumerate(block.columns):
            if width:
                table.columns[index].width = Pt(width)
        for row in block.rows:
            for cell in row:
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
                self._blocks(target, cell.blocks)
                if not cell.blocks:
                    target.add_paragraph()
        self.counts["tables"] = self.counts.get("tables", 0) + 1
        self.counts["table_rows"] = self.counts.get("table_rows", 0) + nrows

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
