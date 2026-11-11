"""Map Typst's weak block gaps and explicit vertical space to Word flow."""
from typing import Any

from docx.oxml.ns import qn
from docx.shared import Pt
from docx.text.paragraph import Paragraph as WordParagraph


class FlowSpacingMixin:
    def _block_gap(self, node: Any, container: Any, before: float | None) -> float:
        previous = node.getprevious()
        # Automatic block spacing disappears at the start of a page/section;
        # explicit #v spacing remains deliberate authored content.
        at_start = previous is None or previous.tag == qn("w:tcPr") or (
            previous.tag == qn("w:p") and previous.find(".//" + qn("w:sectPr")) is not None)
        after = 0.0
        if not at_start and previous.tag == qn("w:p"):
            paragraph = WordParagraph(previous, container)
            value = paragraph.paragraph_format.space_after
            style = paragraph.style
            while value is None and style is not None:
                value = style.paragraph_format.space_after
                style = style.base_style
            after = value.pt if value is not None else 0.0
            paragraph.paragraph_format.space_after = Pt(0)
        weak = max(self._pending_block_space, before or 0, after) if not at_start else 0
        gap = self._pending_space + weak
        self._pending_space = self._pending_block_space = 0.0
        return gap

    def _table_gap(self, container: Any, table: Any, before: float | None) -> None:
        gap = self._block_gap(table._tbl, container, before)
        if not gap:
            return
        previous = table._tbl.getprevious()
        if previous is not None and previous.tag == qn("w:p") and previous.find(".//" + qn("w:sectPr")) is None:
            WordParagraph(previous, container).paragraph_format.space_after = Pt(gap)
        else:
            paragraph = container.add_paragraph()
            fmt = paragraph.paragraph_format
            fmt.space_before = fmt.space_after = Pt(0)
            fmt.line_spacing = Pt(gap)
            fmt.keep_with_next = True
            table._tbl.addprevious(paragraph._p)
