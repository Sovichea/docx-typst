"""Write the shared document model back to an editable DOCX file."""

from __future__ import annotations

from pathlib import Path
from typing import Any

from docx import Document as load_docx
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_COLOR_INDEX
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor
from lxml import etree

from .docx_page_writer import PageWriterMixin
from .docx_table_writer import TableWriterMixin
from .flow_spacing import FlowSpacingMixin
from .media_compat import rasterize_for_docx
from .ooxml import canonical_style_name
from .model import (
    Block,
    CanvasBlock,
    CodeBlock,
    Document,
    DecorativeRule,
    HorizontalRule,
    ImageBlock,
    PageBreak,
    Paragraph,
    Run,
    SectionBreak,
    Spacer,
    StyleDef,
    TableBlock,
    UnsupportedBlock,
)


ALIGN_MAP = {
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
}

HIGHLIGHT_MAP = {
    "#FFFF00": WD_COLOR_INDEX.YELLOW,
    "#00FF00": WD_COLOR_INDEX.BRIGHT_GREEN,
    "#00FFFF": WD_COLOR_INDEX.TURQUOISE,
    "#FF00FF": WD_COLOR_INDEX.PINK,
    "#0000FF": WD_COLOR_INDEX.BLUE,
    "#FF0000": WD_COLOR_INDEX.RED,
    "#000080": WD_COLOR_INDEX.DARK_BLUE,
    "#008080": WD_COLOR_INDEX.TEAL,
    "#008000": WD_COLOR_INDEX.GREEN,
    "#800080": WD_COLOR_INDEX.VIOLET,
    "#800000": WD_COLOR_INDEX.DARK_RED,
    "#808000": WD_COLOR_INDEX.DARK_YELLOW,
    "#808080": WD_COLOR_INDEX.GRAY_50,
    "#C0C0C0": WD_COLOR_INDEX.GRAY_25,
    "#000000": WD_COLOR_INDEX.BLACK,
    "#FFFFFF": WD_COLOR_INDEX.WHITE,
}


class DocxWriter(FlowSpacingMixin, TableWriterMixin, PageWriterMixin):
    def __init__(
        self,
        document: Document,
        output: Path,
        *,
        base: Path,
        reference: Path | None = None,
    ):
        self.model = document
        self.output = output
        self.base = base
        self.reference = reference
        self.doc = load_docx(reference) if reference else load_docx()
        self.counts: dict[str, int] = {}
        self._pending_space = 0.0
        self._pending_block_space = 0.0
        self._pending_page_break = False
        self.source_text = ""
        self.source_records: dict[str, Any] = {}

    def write(self) -> Document:
        self._metadata()
        normal = self.model.styles.get("Normal")
        hyphenate = self.model.hyphenate if self.model.hyphenate is not None else bool(normal and normal.justify)
        hyphenation = OxmlElement("w:autoHyphenation")
        hyphenation.set(qn("w:val"), "1" if hyphenate else "0")
        self.doc.settings.element.append(hyphenation)
        self._styles()
        first = self.doc.sections[0]
        self._page(first, self.model.page)
        self._header_footer_pair(
            first,
            self.model.header,
            self.model.footer,
            header_linked=not self.model.header,
            footer_linked=not self.model.footer,
        )
        self._blocks(self.doc, self.model.blocks)
        from docx.text.paragraph import Paragraph as WordParagraph
        paragraphs = self.doc.element.body.xpath(".//w:p")
        if paragraphs:
            last = WordParagraph(paragraphs[-1], self.doc)
            if not self.doc.element.body.xpath('.//w:bookmarkStart[@w:name="_dxt_document_end"]'):
                self._bookmark(last, "_dxt_document_end", start=True)
                self._bookmark(last, "_dxt_document_end", start=False)
        from .source_preservation import attach_sources
        attach_sources(self)
        self.output.parent.mkdir(parents=True, exist_ok=True)
        self.doc.save(self.output)
        return self.model

    # -- metadata and styles --------------------------------------------

    def _metadata(self) -> None:
        properties = self.doc.core_properties
        metadata = self.model.metadata
        properties.title = metadata.title or ""
        properties.subject = metadata.subject or ""
        properties.author = metadata.author or ""
        properties.keywords = metadata.keywords or ""
        properties.comments = metadata.comments or ""
        properties.category = metadata.category or ""

    def _styles(self) -> None:
        for name, spec in self.model.styles.items():
            self._ensure_style(name, spec)
        for name, spec in self.model.styles.items():
            canonical = canonical_style_name(name) or name
            try:
                style = self.doc.styles[canonical]
            except KeyError:
                continue
            self._apply_style(style, spec)

    def _ensure_style(self, name: str, spec: StyleDef | None = None) -> Any:
        name = canonical_style_name(name) or name
        try:
            return self.doc.styles[name]
        except KeyError:
            style = self.doc.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
            try:
                style.base_style = self.doc.styles["Normal"]
            except KeyError:
                pass
            return style

    def _apply_style(self, style: Any, spec: StyleDef) -> None:
        try:
            if spec.based_on and spec.based_on != spec.name:
                style.base_style = self.doc.styles[spec.based_on]
        except (KeyError, AttributeError):
            pass
        try:
            if spec.next_style:
                style.next_paragraph_style = self.doc.styles[spec.next_style]
        except (KeyError, AttributeError):
            pass
        font = style.font
        if spec.word_font:
            font.name = spec.word_font
            font.bold = False
        elif spec.font:
            from .font_weights import weight_face
            font.name = weight_face(spec.font, spec.weight)
        if spec.word_font or spec.font:
            rpr = style.element.get_or_add_rPr()
            rfonts = rpr.get_or_add_rFonts()
            for attr_name in ("asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme"):
                rfonts.attrib.pop(qn("w:" + attr_name), None)
        if spec.east_asia_font:
            rpr = style.element.get_or_add_rPr()
            fonts = rpr.get_or_add_rFonts()
            fonts.set(qn("w:eastAsia"), spec.east_asia_font)
        if spec.size is not None:
            font.size = Pt(spec.size)
        if spec.bold is not None:
            font.bold = spec.bold
        if spec.weight:
            font.bold = spec.weight.lower() not in {"regular", "normal", "light"}
            if spec.weight.lower() not in {"bold", "regular", "normal"}:
                font.bold = False
        if spec.word_font:
            font.bold = None
        if spec.tracking is not None:
            self._character_spacing(style.element, spec.tracking)
        if spec.italic is not None:
            font.italic = spec.italic
        if spec.underline is not None:
            font.underline = spec.underline
        if spec.strike is not None:
            font.strike = spec.strike
        if spec.small_caps is not None:
            font.small_caps = spec.small_caps
        if spec.all_caps is not None:
            font.all_caps = spec.all_caps
        if spec.color:
            try:
                font.color.rgb = RGBColor.from_string(spec.color.lstrip("#"))
            except ValueError:
                pass
        paragraph = style.paragraph_format
        if spec.align:
            paragraph.alignment = ALIGN_MAP.get(spec.align)
        if spec.justify:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        if spec.space_before is not None or spec.word_before is not None:
            value = spec.word_before if spec.word_before is not None else spec.space_before
            paragraph.space_before = Pt(value)
        if spec.space_after is not None or spec.word_after is not None:
            value = spec.word_after if spec.word_after is not None else spec.space_after
            paragraph.space_after = Pt(value)
        if spec.line_spacing_pt is not None:
            paragraph.line_spacing = Pt(spec.line_spacing_pt)
        elif spec.line_spacing is not None:
            paragraph.line_spacing = spec.line_spacing
        if spec.left_indent is not None:
            paragraph.left_indent = Pt(spec.left_indent)
        if spec.right_indent is not None:
            paragraph.right_indent = Pt(spec.right_indent)
        if spec.first_line_indent is not None:
            paragraph.first_line_indent = Pt(spec.first_line_indent)
        if spec.keep_with_next is not None:
            paragraph.keep_with_next = spec.keep_with_next
        if spec.keep_lines is not None:
            paragraph.keep_together = spec.keep_lines
        if spec.page_break_before is not None:
            paragraph.page_break_before = spec.page_break_before
        if spec.outline_level is not None:
            ppr = style.element.get_or_add_pPr()
            for existing in ppr.findall(qn("w:outlineLvl")):
                ppr.remove(existing)
            outline = OxmlElement("w:outlineLvl")
            outline.set(qn("w:val"), str(spec.outline_level))
            anchor = ppr.find(qn("w:rPr"))
            if anchor is not None:
                anchor.addprevious(outline)
            else:
                ppr.append(outline)

    def _character_spacing(self, style_element: Any, points: float) -> None:
        rpr = style_element.get_or_add_rPr()
        spacing = rpr.find(qn("w:spacing"))
        if spacing is None:
            spacing = OxmlElement("w:spacing")
            try:
                rpr.insert_element_before(
                    spacing,
                    "w:w",
                    "w:kern",
                    "w:position",
                    "w:sz",
                    "w:szCs",
                    "w:highlight",
                    "w:u",
                    "w:effect",
                    "w:bdr",
                    "w:shd",
                    "w:fitText",
                    "w:vertAlign",
                    "w:rtl",
                    "w:cs",
                    "w:em",
                    "w:lang",
                    "w:eastAsianLayout",
                    "w:specVanish",
                    "w:oMath",
                )
            except (AttributeError, ValueError, KeyError):
                rpr.append(spacing)
        spacing.set(qn("w:val"), str(round(points * 20)))

    # -- blocks ----------------------------------------------------------

    def _blocks(self, container: Any, blocks: list[Block]) -> None:
        for block in blocks:
            if isinstance(block, Paragraph):
                self._paragraph(container, block)
            elif isinstance(block, CanvasBlock):
                from .positioned import write_canvas
                write_canvas(self, container, block)
            elif isinstance(block, TableBlock):
                self._flush_page_break(container)
                self._table(container, block)
            elif isinstance(block, ImageBlock):
                self._image_block(container, block)
            elif isinstance(block, CodeBlock):
                self._code(container, block)
            elif isinstance(block, PageBreak):
                self._pending_page_break = True
                self._pending_block_space = 0.0
                self.counts["page_breaks"] = self.counts.get("page_breaks", 0) + 1
            elif isinstance(block, Spacer):
                self._pending_space += max(0.0, block.height_pt)
            elif isinstance(block, HorizontalRule):
                self._flush_page_break(container)
                paragraph = container.add_paragraph()
                fmt = paragraph.paragraph_format
                fmt.space_before = Pt(self._pending_space)
                self._pending_space = 0.0
                fmt.space_after = Pt(0)
                fmt.line_spacing = Pt(1)
                self._bottom_border(paragraph, block.stroke)
            elif isinstance(block, DecorativeRule):
                from .decorations import write_rule
                write_rule(self, container, block)
            elif isinstance(block, SectionBreak):
                self._pending_page_break = False
                self._pending_block_space = 0.0
                self._section(block)
            elif isinstance(block, UnsupportedBlock):
                self._flush_page_break(container)
                paragraph = container.add_paragraph()
                run = paragraph.add_run(f"[Unsupported content: {block.kind}]")
                run.italic = True
                self.counts["unsupported"] = self.counts.get("unsupported", 0) + 1

    def _flush_page_break(self, container: Any) -> None:
        if not self._pending_page_break:
            return
        # A real break character remains visible and deletable with Word's
        # formatting marks. Keep imported pageBreakBefore formatting separate.
        paragraphs = container.paragraphs
        if paragraphs and (paragraphs[-1]._p.getnext() is None or paragraphs[-1]._p.getnext().tag == qn("w:sectPr")):
            paragraph = paragraphs[-1]
        else:
            paragraph = container.add_paragraph()
            paragraph.paragraph_format.space_before = Pt(0)
            paragraph.paragraph_format.space_after = Pt(0)
            paragraph.paragraph_format.line_spacing = Pt(1)
        paragraph.add_run().add_break(WD_BREAK.PAGE)
        self._pending_page_break = False

    def _paragraph(self, container: Any, block: Paragraph) -> None:
        self._flush_page_break(container)
        pending = self._pending_space
        weak = self._pending_block_space
        self._pending_space = 0.0
        self._pending_block_space = 0.0
        spec = self.model.styles.get(block.style or "Normal")
        if weak and spec is not None and spec.word_before is not None:
            weak = spec.word_before
        if weak and block.native_cap_height is not None and block.line_spacing_pt is not None:
            # A standalone native text block starts at the cap edge and ends
            # at its final baseline. Word includes the whole first line box;
            # account for that box in the automatic gap, without changing its
            # font size, line spacing, or explicit authored vertical space.
            weak = max(0.0, weak - (block.line_spacing_pt - block.native_cap_height))
        original_before = block.space_before
        if pending or weak:
            before = block.space_before
            if before is None and spec is not None:
                before = spec.word_before if spec.word_before is not None else spec.space_before
            block.space_before = max(weak, before or 0.0) + pending
        style = self._ensure_style(block.style or "Normal")
        paragraph = container.add_paragraph(style=style)
        self._paragraph_format(paragraph, block)
        from .numbering import apply_numbering
        apply_numbering(self, paragraph, block)
        block.space_before = original_before
        if block.bookmarks:
            for name in block.bookmarks:
                self._bookmark(paragraph, name, start=True)
        for run in block.runs:
            self._run(paragraph, run)
        if block.native_cap_height is not None:
            identity = next((name for name in block.bookmarks if name.startswith("dxt_text_")),
                            f"dxt_text_{len(self.source_records)}")
            if identity not in block.bookmarks:
                self._bookmark(paragraph, identity, start=True)
                self._bookmark(paragraph, identity, start=False)
            self.source_records["text:" + identity] = {"cap": block.native_cap_height}
        if block.align is None and (
            block.heading_level is not None
            or any(run.kind == "break" for run in block.runs)
        ):
            paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
        if block.bookmarks:
            for name in reversed(block.bookmarks):
                self._bookmark(paragraph, name, start=False)
        self.counts["paragraphs"] = self.counts.get("paragraphs", 0) + 1
        if block.heading_level is not None:
            self.counts["headings"] = self.counts.get("headings", 0) + 1
        if block.list_kind:
            self.counts["list_items"] = self.counts.get("list_items", 0) + 1

    def _paragraph_format(self, paragraph: Any, block: Paragraph) -> None:
        if block.align in {"left", "right", "center"} or block.heading_level is not None:
            suppress = OxmlElement("w:suppressAutoHyphens")
            suppress.set(qn("w:val"), "1")
            paragraph._p.get_or_add_pPr().append(suppress)
        if block.align:
            paragraph.alignment = ALIGN_MAP.get(block.align)
        fmt = paragraph.paragraph_format
        if block.space_before is not None:
            fmt.space_before = Pt(block.space_before)
        if block.space_after is not None:
            fmt.space_after = Pt(block.space_after)
        if block.line_spacing_pt is not None:
            fmt.line_spacing = Pt(block.line_spacing_pt)
        elif block.line_spacing is not None:
            fmt.line_spacing = block.line_spacing
        if block.indent_left is not None:
            fmt.left_indent = Pt(block.indent_left)
        if block.indent_right is not None:
            fmt.right_indent = Pt(block.indent_right)
        if block.first_line_indent is not None:
            fmt.first_line_indent = Pt(block.first_line_indent)
        if block.keep_with_next is not None:
            fmt.keep_with_next = block.keep_with_next
        if block.keep_lines is not None:
            fmt.keep_together = block.keep_lines
        if block.page_break_before:
            fmt.page_break_before = True

    # -- runs ------------------------------------------------------------

    def _run(self, paragraph: Any, run: Run) -> None:
        if run.kind == "text" and not run.text.strip():
            if paragraph._p.xpath("./w:r[last()]/w:br"):
                return
        if run.kind == "break":
            paragraph.add_run().add_break(WD_BREAK.LINE)
            return
        if run.kind == "page-break":
            paragraph.add_run().add_break(WD_BREAK.PAGE)
            return
        if run.kind == "column-break":
            paragraph.add_run().add_break(WD_BREAK.COLUMN)
            return
        if run.kind == "tab":
            paragraph.add_run().add_tab()
            return
        if run.kind == "field":
            self._field(paragraph, run)
            return
        if run.kind == "image":
            if run.image is not None:
                self._image(paragraph, run.image)
            return
        if run.kind == "footnote":
            marker = paragraph.add_run(f"[{run.text or 'note'}]")
            marker.font.superscript = True
            self.model.warn(
                "footnote_export",
                "approximated",
                "output",
                source_type="footnote",
                fallback="superscript note text",
                reason="python-docx has no footnote writer",
            )
            return
        if run.kind == "math":
            text_run = paragraph.add_run(run.text)
            text_run.italic = True
            return
        if run.link or run.anchor:
            self._hyperlink(paragraph, run)
            return
        text_run = paragraph.add_run(run.text)
        self._format_run(text_run, run)
        self.counts["runs"] = self.counts.get("runs", 0) + 1

    def _format_run(self, run: Any, model: Run) -> None:
        font = run.font
        if model.code:
            font.name = "Consolas"
        elif model.font:
            font.name = model.font
        if model.size is not None:
            font.size = Pt(model.size)
        if model.bold is not None:
            font.bold = model.bold
        if model.weight and model.weight not in {"bold", "regular", "normal"}:
            from .font_weights import weight_face
            family = model.font or run._parent.style.font.name or self.model.default_font
            font.name = weight_face(family, model.weight)
            font.bold = False
        if model.font and any("\u1780" <= c <= "\u17ff" for c in model.text):
            fonts = run._r.get_or_add_rPr().get_or_add_rFonts()
            fonts.set(qn("w:cs"), model.font)
        if model.italic is not None:
            font.italic = model.italic
        if model.underline is not None:
            font.underline = model.underline
        if model.strike is not None:
            font.strike = model.strike
        if model.small_caps is not None:
            font.small_caps = model.small_caps
        if model.all_caps is not None:
            font.all_caps = model.all_caps
        if model.superscript:
            font.superscript = True
        if model.subscript:
            font.subscript = True
        if model.color:
            try:
                font.color.rgb = RGBColor.from_string(model.color.lstrip("#"))
            except ValueError:
                pass
        if model.highlight:
            self._highlight(run, model.highlight)
        if model.east_asia_font:
            rpr = run._r.get_or_add_rPr()
            fonts = rpr.get_or_add_rFonts()
            fonts.set(qn("w:eastAsia"), model.east_asia_font)

    def _highlight(self, run: Any, color: str) -> None:
        key = color.upper()
        if key in HIGHLIGHT_MAP:
            run.font.highlight_color = HIGHLIGHT_MAP[key]
            return
        rpr = run._r.get_or_add_rPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), color.lstrip("#").upper())
        rpr.append(shd)

    def _hyperlink(self, paragraph: Any, model: Run) -> None:
        run = paragraph.add_run(model.text)
        self._format_run(run, model)
        if model.link and model.color is None:
            run.font.color.rgb = RGBColor(0x05, 0x63, 0xC1)
        if model.link and model.underline is None:
            run.font.underline = True
        hyperlink = OxmlElement("w:hyperlink")
        if model.link:
            try:
                rel_id = paragraph.part.relate_to(
                    model.link, RT.HYPERLINK, is_external=True
                )
                hyperlink.set(qn("r:id"), rel_id)
            except Exception:
                paragraph._p.remove(run._r)
                paragraph.add_run(model.text)
                return
        elif model.anchor:
            hyperlink.set(qn("w:anchor"), model.anchor)
        paragraph._p.append(hyperlink)
        hyperlink.append(run._r)
        if model.style_name == "Hyperlink":
            run.style = "Hyperlink"

    def _field(self, paragraph: Any, model: Run) -> None:
        field = OxmlElement("w:fldSimple")
        name = model.field or "PAGE"
        padded = name.endswith("_PADDED")
        if padded:
            name = name[: -len("_PADDED")]
        target = "PAGEREF _dxt_document_end" if name == "NUMPAGES" else name
        instruction = f' {target} \\# "00" ' if padded else f" {target} "
        field.set(qn("w:instr"), instruction)
        run = paragraph.add_run("01" if padded and name == "PAGE" else "1")
        paragraph._p.append(field)
        field.append(run._r)

    def _bookmark(self, paragraph: Any, name: str, start: bool) -> None:
        element = OxmlElement("w:bookmarkStart" if start else "w:bookmarkEnd")
        element.set(qn("w:id"), str(abs(hash(name)) % 100000))
        if start:
            element.set(qn("w:name"), name)
            p = paragraph._p
            ppr = p.pPr
            if ppr is not None:
                ppr.addnext(element)
            else:
                p.insert(0, element)
        else:
            paragraph._p.append(element)

    # -- media -----------------------------------------------------------

    def _image(self, paragraph: Any, image: Any) -> None:
        path = (self.base / image.path).resolve()
        if not image.path or not path.is_file():
            run = paragraph.add_run(f"[missing image: {image.path}]")
            run.italic = True
            self.model.warn(
                "missing_image",
                "unsupported",
                "output",
                source_type="image",
                fallback="text placeholder",
                reason=f"asset not found: {path}",
            )
            return
        svg_path = path if path.suffix.lower() == ".svg" else None
        path, warning = rasterize_for_docx(path)
        if warning:
            self.model.warn(
                "svg_rasterization",
                "approximated",
                "output",
                source_type="SVG",
                fallback="SVG omitted",
                reason=warning,
            )
            return
        kwargs = {}
        if image.width_pt:
            kwargs["width"] = Pt(image.width_pt)
        elif image.width_percent:
            content_width = self._content_width().pt
            kwargs["width"] = Pt(content_width * image.width_percent)
        if image.height_pt:
            kwargs["height"] = Pt(image.height_pt)
        elif svg_path is not None and "width" in kwargs:
            viewbox = etree.fromstring(svg_path.read_bytes()).get("viewBox", "").replace(",", " ").split()
            if len(viewbox) == 4 and float(viewbox[2]) > 0:
                kwargs["height"] = Pt(kwargs["width"].pt * float(viewbox[3]) / float(viewbox[2]))
        try:
            picture = paragraph.add_run().add_picture(str(path), **kwargs)
            if svg_path is not None:
                from .positioned import embed_svg
                embed_svg(paragraph.part, picture._inline, svg_path)
            if image.typst_source:
                from .source_preservation import record_static
                record_static(self, picture._inline, image)
            self.counts["images"] = self.counts.get("images", 0) + 1
        except Exception as exc:
            self.model.warn(
                "image_embed",
                "unsupported",
                "output",
                source_type="image",
                fallback="text placeholder",
                reason=str(exc),
            )

    def _image_block(self, container: Any, block: ImageBlock) -> None:
        self._flush_page_break(container)
        paragraph = container.add_paragraph()
        paragraph.paragraph_format.line_spacing = 1.0
        paragraph.paragraph_format.space_after = Pt(0)
        paragraph.paragraph_format.space_before = Pt(self._block_gap(paragraph._p, container, block.space_before))
        if block.align:
            paragraph.alignment = ALIGN_MAP.get(block.align, WD_ALIGN_PARAGRAPH.CENTER)
        self._image(paragraph, block.image)
        # Image-only paragraphs must not add an invisible body-font descent
        # below the artwork. The image dimensions themselves remain unchanged.
        mark = OxmlElement("w:rPr")
        size = OxmlElement("w:sz")
        size.set(qn("w:val"), "2")
        mark.append(size)
        paragraph._p.get_or_add_pPr().append(mark)
        for run in paragraph.runs:
            run.font.size = Pt(1)
        self._pending_block_space = block.space_after or 0.0
        if block.caption:
            caption = container.add_paragraph(style=self._ensure_style("Caption"))
            for item in block.caption:
                if isinstance(item, Paragraph):
                    for run in item.runs:
                        self._run(caption, run)

    def _code(self, container: Any, block: CodeBlock) -> None:
        self._flush_page_break(container)
        paragraph = container.add_paragraph()
        if self._pending_space:
            paragraph.paragraph_format.space_before = Pt(self._pending_space)
            self._pending_space = 0.0
        run = paragraph.add_run(block.text)
        run.font.name = "Consolas"
        run.font.size = Pt(9)
        self.counts["code_blocks"] = self.counts.get("code_blocks", 0) + 1
