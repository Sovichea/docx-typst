"""Shared document model used by both conversion directions.

The model deliberately describes document semantics (styles, paragraphs, runs,
tables, sections) rather than WordprocessingML or Typst syntax.  The forward
converter fills it from DOCX, the Typst renderer serializes it, the Typst
parser fills it from source, and the DOCX renderer writes it back.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Union


# ---------------------------------------------------------------------------
# Page and style primitives
# ---------------------------------------------------------------------------


@dataclass
class PageSetup:
    width_pt: float = 595.28
    height_pt: float = 841.89
    margin_top: float = 72.0
    margin_right: float = 72.0
    margin_bottom: float = 72.0
    margin_left: float = 72.0
    header_distance: float = 36.0
    footer_distance: float = 36.0
    orientation: str = "portrait"
    columns: int = 1
    title_page: bool = False
    page_number_start: int | None = None

    def as_dict(self) -> dict[str, float | int | str | bool | None]:
        return {
            "width_pt": self.width_pt,
            "height_pt": self.height_pt,
            "margin_top": self.margin_top,
            "margin_right": self.margin_right,
            "margin_bottom": self.margin_bottom,
            "margin_left": self.margin_left,
            "header_distance": self.header_distance,
            "footer_distance": self.footer_distance,
            "orientation": self.orientation,
            "columns": self.columns,
            "title_page": self.title_page,
            "page_number_start": self.page_number_start,
        }


@dataclass
class StyleDef:
    """Resolved paragraph style properties.

    ``None`` means the property is inherited or unspecified.  Values are
    already resolved through the DOCX style inheritance chain when the model is
    produced by the forward converter.
    """

    name: str
    kind: str = "paragraph"
    based_on: str | None = None
    next_style: str | None = None
    font: str | None = None
    east_asia_font: str | None = None
    word_font: str | None = None
    size: float | None = None
    bold: bool | None = None
    italic: bool | None = None
    underline: bool | None = None
    strike: bool | None = None
    small_caps: bool | None = None
    all_caps: bool | None = None
    color: str | None = None
    highlight: str | None = None
    align: str | None = None
    justify: bool | None = None
    weight: str | None = None
    tracking: float | None = None
    leading: float | None = None
    space_before: float | None = None
    space_after: float | None = None
    word_before: float | None = None
    word_after: float | None = None
    line_spacing: float | None = None
    line_spacing_pt: float | None = None
    left_indent: float | None = None
    right_indent: float | None = None
    first_line_indent: float | None = None
    keep_with_next: bool | None = None
    keep_lines: bool | None = None
    page_break_before: bool | None = None
    outline_level: int | None = None

    def spec(self) -> dict[str, object]:
        out: dict[str, object] = {}
        for key in (
            "based_on",
            "next_style",
            "font",
            "east_asia_font",
            "word_font",
            "size",
            "bold",
            "italic",
            "underline",
            "strike",
            "small_caps",
            "all_caps",
            "color",
            "highlight",
            "align",
            "justify",
            "weight",
            "tracking",
            "leading",
            "space_before",
            "space_after",
            "word_before",
            "word_after",
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
            value = getattr(self, key)
            if value is not None:
                out[key] = value
        return out


# ---------------------------------------------------------------------------
# Inline content
# ---------------------------------------------------------------------------


@dataclass
class ImageRef:
    path: str
    width_pt: float | None = None
    height_pt: float | None = None
    width_percent: float | None = None
    alt: str | None = None
    title: str | None = None
    crop: tuple[float, float, float, float] | None = None
    rotation: float = 0.0
    floating: bool = False
    align: str | None = None
    source_part: str | None = None
    typst_source: str | None = None


@dataclass
class Footnote:
    number: int
    blocks: list["Block"] = field(default_factory=list)


@dataclass
class Run:
    text: str = ""
    kind: str = "text"  # text, break, page-break, column-break, tab, image, field, footnote, math
    bold: bool | None = None
    weight: str | None = None
    italic: bool | None = None
    underline: bool | None = None
    strike: bool | None = None
    small_caps: bool | None = None
    all_caps: bool | None = None
    superscript: bool | None = None
    subscript: bool | None = None
    color: str | None = None
    highlight: str | None = None
    font: str | None = None
    east_asia_font: str | None = None
    size: float | None = None
    code: bool = False
    link: str | None = None
    anchor: str | None = None
    image: ImageRef | None = None
    footnote: Footnote | None = None
    field: str | None = None
    style_name: str | None = None

    def has_formatting(self) -> bool:
        return any(
            value is not None
            for value in (
                self.bold,
                self.italic,
                self.underline,
                self.strike,
                self.small_caps,
                self.all_caps,
                self.superscript,
                self.subscript,
                self.color,
                self.highlight,
                self.font,
                self.size,
            )
        ) or self.code or self.link is not None or self.anchor is not None or self.style_name is not None


# ---------------------------------------------------------------------------
# Block content
# ---------------------------------------------------------------------------


@dataclass
class Paragraph:
    runs: list[Run] = field(default_factory=list)
    style: str | None = "Normal"
    align: str | None = None
    list_kind: str | None = None  # bullet or number
    list_level: int = 0
    list_style: str | None = None
    marker: str | None = None
    number_id: int | None = None
    list_restart: bool = False
    indent_left: float | None = None
    indent_right: float | None = None
    first_line_indent: float | None = None
    space_before: float | None = None
    space_after: float | None = None
    line_spacing: float | None = None
    line_spacing_pt: float | None = None
    keep_with_next: bool | None = None
    keep_lines: bool | None = None
    page_break_before: bool | None = None
    heading_level: int | None = None
    outline_level: int | None = None
    bookmarks: list[str] = field(default_factory=list)
    quoted: bool = False

    def plain_text(self) -> str:
        return "".join(run.text for run in self.runs if run.kind == "text")


@dataclass
class TableCell:
    blocks: list["Block"] = field(default_factory=list)
    colspan: int = 1
    rowspan: int = 1
    fill: str | None = None
    valign: str | None = None
    width_pt: float | None = None
    row: int = 0
    col: int = 0


@dataclass
class TableBlock:
    rows: list[list[TableCell]] = field(default_factory=list)
    columns: list[float | None] = field(default_factory=list)
    header_rows: int = 0
    style: str | None = None
    align: str | None = None
    stroke: str | None = None
    autofit: bool = True
    keep_together: bool = False
    cell_margins_pt: float | None = None
    space_before: float | None = None
    space_after: float | None = None
    native_text_bounds: bool = False
    borders: dict[str, str] = field(default_factory=dict)
    column_weights: list[float] = field(default_factory=list)


@dataclass
class ImageBlock:
    image: ImageRef
    caption: list["Block"] = field(default_factory=list)
    align: str = "center"
    space_before: float | None = None
    space_after: float | None = None

    @property
    def path(self) -> str:
        return self.image.path


@dataclass
class CodeBlock:
    text: str
    language: str | None = None
    style: str | None = None


@dataclass
class PageBreak:
    pass


@dataclass
class Spacer:
    height_pt: float = 0.0


@dataclass
class HorizontalRule:
    stroke: str | None = None


@dataclass
class DecorativeRule:
    width: float = 38.0
    height: float = 3.0
    fill: str = "#F7EE16"
    after: float = 10.0


@dataclass
class SectionBreak:
    page: PageSetup
    header: list["Block"] = field(default_factory=list)
    footer: list["Block"] = field(default_factory=list)
    start_type: str = "nextPage"
    header_linked: bool = True
    footer_linked: bool = True


@dataclass
class UnsupportedBlock:
    kind: str
    detail: str = ""
    text: str = ""


@dataclass
class PositionedBlock:
    blocks: list["Block"] = field(default_factory=list)
    x: float = 0.0
    y: float = 0.0
    width: float = 100.0
    height: float = 100.0
    style: str | None = None
    align: str = "left"
    source: str = ""


@dataclass
class CanvasBlock:
    blocks: list["Block"] = field(default_factory=list)
    width: float = 595.2756
    height: float = 841.8898
    identifier: str = "canvas"
    source: str = ""


@dataclass
class NativeBlock:
    source: str = ""


Block = Union[
    DecorativeRule,
    CanvasBlock,
    PositionedBlock,
    NativeBlock,
    Paragraph,
    TableBlock,
    ImageBlock,
    CodeBlock,
    PageBreak,
    Spacer,
    HorizontalRule,
    SectionBreak,
    UnsupportedBlock,
]


# ---------------------------------------------------------------------------
# Document
# ---------------------------------------------------------------------------


@dataclass
class Metadata:
    title: str | None = None
    subject: str | None = None
    author: str | None = None
    last_modified_by: str | None = None
    keywords: str | None = None
    comments: str | None = None
    category: str | None = None
    created: str | None = None
    modified: str | None = None


@dataclass
class Asset:
    original_name: str
    output_path: str
    sha256: str
    byte_count: int
    content_type: str = ""
    source_part: str = ""
    derived_path: str | None = None
    warning: str | None = None
    emitted: bool = True


@dataclass
class Diagnostic:
    issue: str
    outcome: str  # supported, approximated, unsupported, skipped
    location: str
    source_type: str = ""
    fallback: str = ""
    reason: str = ""
    expected_loss: str = ""

    def as_dict(self) -> dict[str, str]:
        return {
            "issue": self.issue,
            "outcome": self.outcome,
            "location": self.location,
            "source_type": self.source_type,
            "fallback": self.fallback,
            "reason": self.reason,
            "expected_loss": self.expected_loss,
        }


@dataclass
class Document:
    metadata: Metadata = field(default_factory=Metadata)
    hyphenate: bool | None = None
    page: PageSetup = field(default_factory=PageSetup)
    blocks: list[Block] = field(default_factory=list)
    styles: dict[str, StyleDef] = field(default_factory=dict)
    header: list[Block] = field(default_factory=list)
    footer: list[Block] = field(default_factory=list)
    assets: list[Asset] = field(default_factory=list)
    diagnostics: list[Diagnostic] = field(default_factory=list)
    source: str | None = None
    section_count: int = 1
    default_font: str = "Calibri"
    default_size: float = 11.0

    def warn(
        self,
        issue: str,
        outcome: str,
        location: str,
        *,
        source_type: str = "",
        fallback: str = "",
        reason: str = "",
        expected_loss: str = "",
    ) -> None:
        self.diagnostics.append(
            Diagnostic(
                issue=issue,
                outcome=outcome,
                location=location,
                source_type=source_type,
                fallback=fallback,
                reason=reason,
                expected_loss=expected_loss,
            )
        )

    def plain_text(self, separator: str = "\n") -> str:
        parts: list[str] = []

        def walk(blocks: list[Block]) -> None:
            for block in blocks:
                if isinstance(block, Paragraph):
                    parts.append(block.plain_text())
                elif isinstance(block, TableBlock):
                    for row in block.rows:
                        parts.append("\t".join(cell_text(cell) for cell in row))
                elif isinstance(block, CodeBlock):
                    parts.append(block.text)
                elif isinstance(block, SectionBreak):
                    walk(block.header)
                    walk(block.footer)

        def cell_text(cell: TableCell) -> str:
            before = len(parts)
            walk(cell.blocks)
            return " ".join(parts[before:])

        walk(self.blocks)
        return separator.join(parts)
