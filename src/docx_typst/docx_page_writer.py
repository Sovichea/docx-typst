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
        previous = self.doc.sections[-1]
        changed_width = abs(previous.page_width.pt - block.page.width_pt) > 0.1
        section = self.doc.add_section(WD_SECTION.NEW_PAGE)
        # add_section creates an empty paragraph to hold the old sectPr. A
        # Typst page switch has no such blank line; attach it to the preceding
        # body paragraph so a near-bottom caption is not pushed to a new page.
        boundary = self.doc.element.body[-2]
        preceding = boundary.getprevious()
        section_properties = boundary.find("./" + qn("w:pPr") + "/" + qn("w:sectPr"))
        if section_properties is not None and preceding is not None:
            if preceding.tag == qn("w:p"):
                preceding.get_or_add_pPr().append(section_properties)
                boundary.getparent().remove(boundary)
            else:
                from docx.text.paragraph import Paragraph
                spacer = Paragraph(boundary, self.doc)
                spacer.paragraph_format.space_before = spacer.paragraph_format.space_after = Pt(0)
                spacer.paragraph_format.line_spacing = Pt(1)
        self._page(section, block.page)
        header = block.header
        footer = block.footer
        header_linked, footer_linked = block.header_linked, block.footer_linked
        if changed_width and header_linked:
            header, header_linked = getattr(self, "_active_header", []), False
        if changed_width and footer_linked:
            footer, footer_linked = getattr(self, "_active_footer", []), False
        self._header_footer_pair(
            section,
            header,
            footer,
            header_linked=header_linked,
            footer_linked=footer_linked,
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
        if not header_linked:
            self._active_header = header_blocks
        if not footer_linked:
            self._active_footer = footer_blocks
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
            pending = self._pending_space, self._pending_block_space, self._pending_page_break
            self._pending_space = self._pending_block_space = 0.0
            self._pending_page_break = False
            self._blocks(header, blocks)
            self._pending_space, self._pending_block_space, self._pending_page_break = pending
            if not blocks:
                header.add_paragraph()

    def _content_width(self) -> Any:
        section = self.doc.sections[-1]
        return Pt(max(72.0, section.page_width.pt - section.left_margin.pt - section.right_margin.pt))

    def _bottom_border(self, paragraph: Any, stroke: str | None = None) -> None:
        ppr = paragraph._p.get_or_add_pPr()
        borders = OxmlElement("w:pBdr")
        bottom = OxmlElement("w:bottom")
        bottom.set(qn("w:val"), "single")
        bottom.set(qn("w:sz"), "6")
        bottom.set(qn("w:space"), "1")
        bottom.set(qn("w:color"), "808080")
        if stroke:
            from .typst_syntax import UnitValue, parse_value, color_to_rgb, split_top_level
            for part in split_top_level(stroke, "+"):
                value = parse_value(part)
                if isinstance(value, UnitValue):
                    bottom.set(qn("w:sz"), str(round(value.to_pt() * 8)))
                color = color_to_rgb(value)
                if color:
                    bottom.set(qn("w:color"), "".join(f"{c:02X}" for c in color))
        borders.append(bottom)
        ppr.append(borders)
