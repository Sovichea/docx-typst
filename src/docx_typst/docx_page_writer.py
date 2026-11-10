"""Page, section, and header/footer rendering for the reverse converter."""

from __future__ import annotations

from typing import Any

from docx.enum.section import WD_ORIENT, WD_SECTION
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt

from .model import Block, PageSetup, SectionBreak


class PageWriterMixin:
    def _section(self, block: SectionBreak) -> None:
        section = self.doc.add_section(WD_SECTION.NEW_PAGE)
        self._page(section, block.page)
        self._header_footer_pair(
            section,
            block.header,
            block.footer,
            header_linked=block.header_linked,
            footer_linked=block.footer_linked,
        )
        self.counts["sections"] = self.counts.get("sections", 0) + 1

    def _page(self, section: Any, page: PageSetup) -> None:
        section.orientation = (
            WD_ORIENT.LANDSCAPE
            if page.width_pt > page.height_pt
            else WD_ORIENT.PORTRAIT
        )
        section.page_width = Pt(page.width_pt)
        section.page_height = Pt(page.height_pt)
        section.top_margin = Pt(page.margin_top)
        section.right_margin = Pt(page.margin_right)
        section.bottom_margin = Pt(page.margin_bottom)
        section.left_margin = Pt(page.margin_left)
        section.header_distance = Pt(
            0.0
            if page.margin_top == 0.0 and page.margin_bottom == 0.0
            else page.header_distance
        )
        section.footer_distance = Pt(
            0.0
            if page.margin_top == 0.0 and page.margin_bottom == 0.0
            else page.footer_distance
        )
        sect_pr = section._sectPr
        for existing in sect_pr.findall(qn("w:pgNumType")):
            sect_pr.remove(existing)
        if page.page_number_start is not None:
            number_type = OxmlElement("w:pgNumType")
            number_type.set(qn("w:start"), str(page.page_number_start))
            sect_pr.append(number_type)
        for existing in sect_pr.findall(qn("w:cols")):
            sect_pr.remove(existing)
        cols = OxmlElement("w:cols")
        cols.set(qn("w:num"), str(max(1, page.columns)))
        doc_grid = sect_pr.find(qn("w:docGrid"))
        if doc_grid is not None:
            doc_grid.addprevious(cols)
        else:
            sect_pr.append(cols)
        for existing in sect_pr.findall(qn("w:titlePg")):
            sect_pr.remove(existing)
        if page.title_page:
            title = OxmlElement("w:titlePg")
            doc_grid = sect_pr.find(qn("w:docGrid"))
            if doc_grid is not None:
                doc_grid.addprevious(title)
            else:
                sect_pr.append(title)

    def _header_footer_pair(
        self,
        section: Any,
        header_blocks: list[Block],
        footer_blocks: list[Block],
        *,
        header_linked: bool,
        footer_linked: bool,
    ) -> None:
        for header, blocks, linked in (
            (section.header, header_blocks, header_linked),
            (section.footer, footer_blocks, footer_linked),
        ):
            if linked:
                header.is_linked_to_previous = True
                continue
            header.is_linked_to_previous = False
            element = header._element
            for child in list(element):
                if child.tag in {qn("w:p"), qn("w:tbl")}:
                    element.remove(child)
            self._blocks(header, blocks)
            if not blocks:
                header.add_paragraph()

    def _content_width(self) -> Any:
        page = self.model.page
        return Pt(max(72.0, page.width_pt - page.margin_left - page.margin_right))

    def _bottom_border(self, paragraph: Any) -> None:
        ppr = paragraph._p.get_or_add_pPr()
        borders = OxmlElement("w:pBdr")
        bottom = OxmlElement("w:bottom")
        bottom.set(qn("w:val"), "single")
        bottom.set(qn("w:sz"), "6")
        bottom.set(qn("w:space"), "1")
        bottom.set(qn("w:color"), "808080")
        borders.append(bottom)
        ppr.append(borders)
