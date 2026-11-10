"""Read a Word document into the shared semantic model."""

from __future__ import annotations

from collections import Counter
from datetime import date, datetime
import re
from typing import Any, Iterable
from pathlib import Path

from docx import Document as load_docx
from lxml import etree

from .model import (
    Asset,
    Block,
    Diagnostic,
    Document,
    Paragraph,
    SectionBreak,
    StyleDef,
)
from .ooxml import (
    W,
    NumberingResolver,
    StyleResolver,
    attr,
    first,
    half_points,
    local_name,
    page_setup,
    parse_theme,
    section_break_before,
)


class BaseDocxReader:
    """Own the package, style cascade, section flow, and report state."""

    def __init__(self, source: Path, output: Path, *, extract_media: bool = True, keep_empty: bool = True):
        self.source = source.resolve()
        self.output = output.resolve()
        self.extract_media = extract_media
        self.keep_empty = keep_empty
        self.docx = load_docx(self.source)
        self.model = Document(source=self.source.name)
        self.theme = parse_theme(self._theme_root())
        self.styles = StyleResolver(self.docx.styles.element, self.theme)
        self.numbering = NumberingResolver(self._numbering_root())
        self.media_dir = self.output / "assets" / "media"
        self.assets_by_hash: dict[str, Asset] = {}
        self.image_counter = 0
        self.used_style_ids: set[str] = set()
        self.counts: Counter[str] = Counter()
        self._warned: set[str] = set()
        self._footnotes = self._load_notes("footnotes")
        self._endnotes = self._load_notes("endnotes")
        self._current_part: Any = self.docx.part
        self._default_font = "Calibri"
        self._default_size = 11.0

    # -- package parts ---------------------------------------------------

    def _theme_root(self) -> Any:
        for rel in self.docx.part.rels.values():
            if rel.reltype.endswith("/theme") and not rel.is_external:
                try:
                    return etree.fromstring(rel.target_part.blob)
                except Exception:
                    return None
        return None

    def _numbering_root(self) -> Any:
        try:
            return self.docx.part.numbering_part.element
        except Exception:
            return None

    def _load_notes(self, name: str) -> dict[str, list[Block]]:
        notes: dict[str, list[Block]] = {}
        target = f"/word/{name}.xml"
        for part in self.docx.part.package.iter_parts():
            if str(part.partname) != target:
                continue
            try:
                root = part.element
            except Exception:
                continue
            for note in root:
                if local_name(note) not in {"footnote", "endnote"}:
                    continue
                note_id = note.get(f"{{{W}}}id")
                if note_id is None:
                    continue
                notes[note_id] = self._read_container(note, part, section_break=False)
            break
        return notes

    # -- warnings --------------------------------------------------------

    def warn(self, issue: str, outcome: str, location: str, **kwargs: str) -> None:
        key = f"{issue}:{location}"
        if key in self._warned:
            return
        self._warned.add(key)
        self.model.diagnostics.append(
            Diagnostic(issue=issue, outcome=outcome, location=location, **kwargs)
        )

    # -- document --------------------------------------------------------

    def read(self) -> Document:
        self._read_defaults()
        self.model.metadata = self._metadata()
        sections = list(self.docx.sections)
        self.model.section_count = max(1, len(sections))
        first_section = sections[0] if sections else None
        if first_section is not None:
            self.model.page = page_setup(first_section)
            self.model.header, self.model.footer = self._read_header_footer(first_section)
        self.model.blocks = self._read_body(sections)
        self.model.styles = self._build_styles()
        self.model.assets = list(self.assets_by_hash.values())
        self.model.default_font = self._default_font
        self.model.default_size = self._default_size
        self._finalize_counts()
        return self.model

    def _read_defaults(self) -> None:
        rpr = self.styles.default_rpr
        if rpr is not None:
            fonts = first(rpr, "./w:rFonts")
            if fonts is not None:
                font = (
                    fonts.get(f"{{{W}}}ascii")
                    or fonts.get(f"{{{W}}}hAnsi")
                    or self.theme.font(fonts.get(f"{{{W}}}asciiTheme"))
                )
                if font:
                    self._default_font = font
            size = half_points(attr(first(rpr, "./w:sz"), "w:val"))
            if size:
                self._default_size = size

    def _metadata(self) -> Any:
        props = self.docx.core_properties
        model = self.model.metadata
        model.title = props.title or None
        model.subject = props.subject or None
        model.author = props.author or None
        model.last_modified_by = props.last_modified_by or None
        model.keywords = props.keywords or None
        model.comments = props.comments or None
        model.category = props.category or None
        model.created = _iso(props.created)
        model.modified = _iso(props.modified)
        return model

    # -- body ------------------------------------------------------------

    def _read_body(self, sections: list[Any]) -> list[Block]:
        body = self.docx.element.body
        blocks: list[Block] = []
        section_index = 0
        for node in body:
            name = local_name(node)
            if name == "p":
                blocks.extend(self._read_paragraph(node, self.docx.part))
                if section_break_before(node) and section_index + 1 < len(sections):
                    section_index += 1
                    blocks.append(self._section_break(sections[section_index]))
            elif name == "tbl":
                blocks.append(self._read_table(node, self.docx.part))
            elif name in {"sdt", "customXml", "smartTag"}:
                content = first(node, "./w:sdtContent")
                if content is None:
                    content = node
                blocks.extend(self._read_container(content, self.docx.part))
                self.warn(
                    "content_control",
                    "approximated",
                    f"body:{self.counts['blocks']}",
                    source_type="structured document tag",
                    fallback="control content flattened into ordinary paragraphs",
                    reason="Typst has no content-control equivalent",
                    expected_loss="Content is preserved; the control itself is not",
                )
            elif name == "altChunk":
                self.warn(
                    "alt_chunk",
                    "unsupported",
                    "body",
                    source_type="altChunk",
                    fallback="omitted",
                    reason="embedded alternative-format content is not parsed",
                    expected_loss="Embedded content is missing from the Typst project",
                )
            elif name == "sectPr":
                continue
            self.counts["blocks"] += 1
        return blocks

    def _read_container(self, container: Any, part: Any, *, section_break: bool = True) -> list[Block]:
        previous = self._current_part
        self._current_part = part
        blocks: list[Block] = []
        try:
            for node in container:
                name = local_name(node)
                if name == "p":
                    blocks.extend(self._read_paragraph(node, part))
                elif name == "tbl":
                    blocks.append(self._read_table(node, part))
                elif name in {"sdt", "customXml", "smartTag"}:
                    content = first(node, "./w:sdtContent")
                    blocks.extend(
                        self._read_container(
                            content if content is not None else node,
                            part,
                            section_break=False,
                        )
                    )
                elif name == "altChunk":
                    self.warn(
                        "alt_chunk",
                        "unsupported",
                        "container",
                        source_type="altChunk",
                        fallback="omitted",
                    )
        finally:
            self._current_part = previous
        return blocks

    def _read_header_footer(self, section: Any) -> tuple[list[Block], list[Block]]:
        header_blocks: list[Block] = []
        footer_blocks: list[Block] = []
        try:
            if not section.header.is_linked_to_previous:
                header_blocks = self._read_container(
                    section.header._element, section.header.part, section_break=False
                )
        except Exception:
            pass
        try:
            if not section.footer.is_linked_to_previous:
                footer_blocks = self._read_container(
                    section.footer._element, section.footer.part, section_break=False
                )
        except Exception:
            pass
        if section.different_first_page_header_footer:
            self.warn(
                "first_page_header",
                "approximated",
                "section:1",
                source_type="section header/footer",
                fallback="default header and footer used",
                reason="Typst's document header is not section-aware",
                expected_loss="A distinct first-page header or footer may not be reproduced",
            )
        return header_blocks, footer_blocks

    def _section_break(self, section: Any) -> SectionBreak:
        page = page_setup(section)
        header: list[Block] = []
        footer: list[Block] = []
        header_linked = True
        footer_linked = True
        try:
            header_linked = bool(section.header.is_linked_to_previous)
            if not header_linked:
                header = self._read_container(
                    section.header._element, section.header.part, section_break=False
                )
        except Exception:
            pass
        try:
            footer_linked = bool(section.footer.is_linked_to_previous)
            if not footer_linked:
                footer = self._read_container(
                    section.footer._element, section.footer.part, section_break=False
                )
        except Exception:
            pass
        start_type = "nextPage"
        try:
            start_type = str(section.start_type)
        except Exception:
            pass
        if start_type not in {"NEW_PAGE", "nextPage"}:
            self.warn(
                "section_start_type",
                "approximated",
                f"section:{self.model.section_count}",
                source_type=start_type,
                fallback="page break",
                reason="Typst does not expose Word section start modes",
                expected_loss="Continuous or even/odd section starts may add a page break",
            )
        return SectionBreak(
            page=page,
            header=header,
            footer=footer,
            start_type=start_type,
            header_linked=header_linked,
            footer_linked=footer_linked,
        )

    # -- styles ----------------------------------------------------------

    def _build_styles(self) -> dict[str, StyleDef]:
        styles: dict[str, StyleDef] = {}
        style_ids = set(self.used_style_ids)
        style_ids.add(self.styles.default_paragraph_id)
        for style_id in sorted(filter(None, style_ids)):
            name = self.styles.style_name(style_id) or style_id
            ppr = self.styles.style_paragraph_properties(style_id)
            rpr = self.styles.style_run_properties(style_id)
            spec = StyleDef(
                name=name,
                kind="paragraph",
                based_on=self.styles.style_name(self.styles.based_on(style_id)),
                next_style=self.styles.style_name(self.styles.next_style(style_id)),
            )
            for key in (
                "font",
                "east_asia_font",
                "size",
                "bold",
                "italic",
                "underline",
                "strike",
                "small_caps",
                "all_caps",
                "color",
                "highlight",
            ):
                if key in rpr:
                    setattr(spec, key, rpr[key])
            for key in (
                "align",
                "space_before",
                "space_after",
                "line_spacing",
                "line_spacing_pt",
                "left_indent",
                "right_indent",
                "first_line_indent",
                "keep_with_next",
                "keep_lines",
                "page_break_before",
                "outline_level",
            ):
                if key in ppr:
                    setattr(spec, key, ppr[key])
            if spec.align == "justify":
                spec.justify = True
            styles[name] = spec
        if "Normal" not in styles:
            styles["Normal"] = StyleDef(
                name="Normal",
                font=self._default_font,
                size=self._default_size,
            )
        return styles

    # -- counts ----------------------------------------------------------

    def _finalize_counts(self) -> None:
        text = self.model.plain_text(" ")
        self.counts["words"] = len(re.findall(r"\S+", text))
        self.counts["characters"] = len(text)
        self.counts["styles"] = len(self.model.styles)
        self.counts["sections"] = self.model.section_count
        self.counts["unsupported"] = sum(
            1 for item in self.model.diagnostics if item.outcome == "unsupported"
        )
        self.counts["approximated"] = sum(
            1 for item in self.model.diagnostics if item.outcome == "approximated"
        )


# ---------------------------------------------------------------------------
# Helpers shared by the reader mixins
# ---------------------------------------------------------------------------


def _iso(value: Any) -> str | None:
    if isinstance(value, (datetime, date)):
        return value.isoformat()
    return None


def _diff(effective: dict[str, Any], base: dict[str, Any]) -> dict[str, Any]:
    return {key: value for key, value in effective.items() if base.get(key) != value}


def _heading_level(style_name: str, style_id: str | None, outline_level: Any) -> int | None:
    if style_name in {"Title", "Subtitle", "TOC Heading"}:
        return None
    match = re.match(r"^(?:Heading|heading)\s*(\d+)$", style_name)
    if match:
        return min(6, max(1, int(match.group(1))))
    if style_id and re.match(r"^Heading\d+$", style_id):
        return min(6, max(1, int(re.sub(r"\D", "", style_id) or 1)))
    if isinstance(outline_level, int) and 0 <= outline_level < 6:
        return outline_level + 1
    return None
