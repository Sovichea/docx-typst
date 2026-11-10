"""Parse the documented static Typst subset into the shared document model."""

from __future__ import annotations

from pathlib import Path
import re
from typing import Any

from .model import (
    Block,
    CodeBlock,
    Document,
    PageBreak,
    Paragraph,
    Run,
    SectionBreak,
    StyleDef,
)
from .typst_blocks import BlockBuilderMixin
from .typst_inline import InlineParser
from .typst_syntax import (
    Content,
    balanced,
    length_to_pt,
    parse_call_arguments,
    strip_comments,
)
from .typst_values import (
    page_setup_from_dict,
    string_value as _string,
    style_from_spec,
)


class TypstParser(BlockBuilderMixin):
    def __init__(self, source: Path, text: str):
        self.source = source
        self.raw_text = text
        self.text = strip_comments(text)
        self.document = Document(source=source.name)
        self.inline = InlineParser(self.document, "body")
        self._current_page = self.document.page
        self._pending_page_number_start: int | None = None
        self._table_settings: dict[str, Any] = {}

    def parse(self) -> Document:
        style_value, style_span = self._style_table()
        if style_value is not None:
            self._apply_styles(style_value)
        show_args, content_start = self._document_show()
        if show_args is not None:
            self._apply_show_args(show_args)
            content = self.text[content_start:]
        else:
            direct = self._direct_document_body()
            if direct is not None:
                args, body, content_start = direct
                self._apply_show_args(args)
                content = body
            else:
                content = self.text
                if style_span is not None:
                    content = content[: style_span[0]] + "\n" * (
                        content[: style_span[1]].count("\n")
                    ) + content[style_span[1] :]
        self._current_page = self.document.page
        self.document.blocks = self._parse_blocks(content, "body")
        if "Normal" not in self.document.styles:
            self.document.styles["Normal"] = StyleDef(
                name="Normal",
                font=self.document.default_font,
                size=self.document.default_size,
            )
        return self.document

    # -- document preamble ----------------------------------------------

    def _style_table(self) -> tuple[Any, tuple[int, int] | None]:
        match = re.search(r"#let\s+docx-styles\s*=\s*\(", self.text)
        if not match:
            return None, None
        try:
            inner, end = balanced(self.text, match.end() - 1, "(", ")")
        except ValueError:
            return None, None
        from .typst_syntax import parse_value

        return parse_value("(" + inner + ")"), (match.start(), end)

    def _document_show(self) -> tuple[dict[str, Any] | None, int]:
        match = re.search(r"#show\s*:\s*docx-document\.with\s*\(", self.text)
        if not match:
            return None, 0
        try:
            inner, end = balanced(self.text, match.end() - 1, "(", ")")
        except ValueError:
            return None, 0
        _positional, named = parse_call_arguments(inner)
        return named, end

    def _direct_document_body(self) -> tuple[dict[str, Any], str, int] | None:
        match = re.search(r"#docx-document\s*\(", self.text)
        if not match:
            return None
        from .typst_syntax import parse_call_at

        call = parse_call_at(self.text, match.start() + 1)
        if call is None or call.body is None:
            return None
        return call.named, call.body.text, match.start()

    def _apply_show_args(self, args: dict[str, Any]) -> None:
        metadata = self.document.metadata
        metadata.title = _string(args.get("title"))
        metadata.author = _string(args.get("author"))
        metadata.subject = _string(args.get("subject"))
        metadata.keywords = _string(args.get("keywords"))
        page = args.get("page")
        if isinstance(page, dict):
            self.document.page = page_setup_from_dict(page)
        for key, attr_name in (
            ("header-distance", "header_distance"),
            ("footer-distance", "footer_distance"),
        ):
            value = args.get(key)
            if value is not None:
                setattr(
                    self.document.page,
                    attr_name,
                    length_to_pt(value, getattr(self.document.page, attr_name)),
                )
        header = args.get("header")
        footer = args.get("footer")
        if isinstance(header, Content):
            self.document.header = self._parse_blocks(header.text, "header")
        if isinstance(footer, Content):
            self.document.footer = self._parse_blocks(footer.text, "footer")

    def _apply_styles(self, value: Any) -> None:
        if not isinstance(value, dict):
            return
        for name, spec in value.items():
            if not isinstance(name, str) or not isinstance(spec, dict):
                continue
            self.document.styles[name] = style_from_spec(name, spec)
        normal = self.document.styles.get("Normal")
        if normal is not None:
            if normal.font:
                self.document.default_font = normal.font
            if normal.size:
                self.document.default_size = normal.size

    # -- block parsing ---------------------------------------------------

    def _parse_blocks(self, text: str, location: str) -> list[Block]:
        outer_table_settings = self._table_settings.copy()
        blocks: list[Block] = []
        i = 0
        length = len(text)
        while i < length:
            while i < length and text[i] in " \t\r\n":
                i += 1
            if i >= length:
                break
            line_end = text.find("\n", i)
            if line_end < 0:
                line_end = length
            line = text[i:line_end]
            stripped = line.lstrip()
            indent = len(line) - len(stripped)
            if stripped.startswith("```"):
                fence = stripped[: len(stripped) - len(stripped.lstrip("`"))]
                language = stripped[len(fence) :].strip() or None
                body_start = line_end + 1
                close = text.find(fence, body_start)
                if close < 0:
                    close = length
                blocks.append(
                    CodeBlock(
                        text=text[body_start:close].rstrip("\n"), language=language
                    )
                )
                i = close + len(fence)
                continue
            heading = re.match(r"^(=+)\s+(.*)$", stripped)
            if heading:
                level = min(6, len(heading.group(1)))
                blocks.append(
                    Paragraph(
                        style=f"Heading {level}",
                        heading_level=level,
                        runs=self.inline.parse(heading.group(2)),
                        outline_level=level - 1,
                    )
                )
                i = line_end + 1
                continue
            item = re.match(r"^([-+])\s+(.*)$", stripped)
            if item:
                kind = "bullet" if item.group(1) == "-" else "number"
                level = min(8, indent // 2)
                blocks.append(
                    Paragraph(
                        style="List Bullet" if kind == "bullet" else "List Number",
                        list_kind=kind,
                        list_level=level,
                        list_style=(
                            "List Bullet" if kind == "bullet" else "List Number"
                        ),
                        runs=self.inline.parse(item.group(2)),
                    )
                )
                i = line_end + 1
                continue
            if stripped.startswith("#set "):
                end = self._set_rule(text, i + indent, blocks, location)
                if end is not None:
                    i = end
                    continue
            if stripped.startswith("#show"):
                i = self._skip_show_rule(text, i + indent)
                continue
            if stripped.startswith("#import") or stripped.startswith("#let "):
                i = line_end + 1
                continue
            if stripped.startswith("#"):
                from .typst_syntax import parse_call_at

                call = parse_call_at(text, i + indent + 1)
                if call is not None:
                    block = self._call_block(call, location)
                    if block is not None:
                        if isinstance(block, list):
                            blocks.extend(block)
                        else:
                            blocks.append(block)
                    i = i + indent + 1 + len(call.raw)
                    continue
            start = i
            end = line_end
            cursor = line_end + 1
            while cursor < length:
                next_end = text.find("\n", cursor)
                if next_end < 0:
                    next_end = length
                next_line = text[cursor:next_end]
                next_stripped = next_line.strip()
                if not next_stripped:
                    break
                if (
                    next_stripped.startswith("#")
                    or next_stripped.startswith("```")
                    or re.match(r"^=+\s+", next_stripped)
                    or re.match(r"^[-+]\s+", next_stripped)
                ):
                    break
                end = next_end
                cursor = next_end + 1
            body = text[start:end].strip()
            if body:
                blocks.append(
                    Paragraph(style="Normal", runs=self.inline.parse(body))
                )
            i = end + 1
        self._table_settings = outer_table_settings
        return blocks

    def _skip_show_rule(self, text: str, start: int) -> int:
        line_end = text.find("\n", start)
        if line_end < 0:
            line_end = len(text)
        brace = text.find("{", start, line_end + 1)
        if brace < 0:
            return line_end + 1
        try:
            _inner, end = balanced(text, brace, "{", "}")
            return end
        except ValueError:
            return line_end + 1

    def _set_rule(
        self, text: str, start: int, blocks: list[Block], location: str
    ) -> int | None:
        match = re.match(r"#set\s+([A-Za-z-]+)\s*\(", text[start:])
        if not match:
            return None
        name = match.group(1)
        paren = start + match.end() - 1
        try:
            inner, end = balanced(text, paren, "(", ")")
        except ValueError:
            return None
        _positional, named = parse_call_arguments(inner)
        if name == "page":
            page = page_setup_from_dict(named, self._current_page)
            if self._pending_page_number_start is not None:
                page.page_number_start = self._pending_page_number_start
                self._pending_page_number_start = None
            self._current_page = page
            header = named.get("header")
            footer = named.get("footer")
            header_blocks = (
                self._parse_blocks(header.text, "header")
                if isinstance(header, Content)
                else []
            )
            footer_blocks = (
                self._parse_blocks(footer.text, "footer")
                if isinstance(footer, Content)
                else []
            )
            if not blocks:
                self.document.page = page
                if "header" in named:
                    self.document.header = header_blocks
                if "footer" in named:
                    self.document.footer = footer_blocks
            else:
                if blocks and isinstance(blocks[-1], PageBreak):
                    blocks.pop()
                blocks.append(
                    SectionBreak(
                        page=page,
                        header=header_blocks,
                        footer=footer_blocks,
                        header_linked="header" not in named
                        or named.get("header") is None,
                        footer_linked="footer" not in named
                        or named.get("footer") is None,
                    )
                )
        elif name == "table":
            self._table_settings.update(named)
        elif name == "text":
            font = named.get("font")
            if isinstance(font, str):
                self.document.default_font = font
            size = named.get("size")
            if size is not None:
                self.document.default_size = length_to_pt(
                    size, self.document.default_size
                )
        elif name == "document":
            metadata = self.document.metadata
            metadata.title = _string(named.get("title")) or metadata.title
            metadata.author = _string(named.get("author")) or metadata.author
            metadata.keywords = _string(named.get("keywords")) or metadata.keywords
        return end


def parse_typst(source: str | Path) -> Document:
    path = Path(source)
    text = path.read_text(encoding="utf-8")
    return parse_typst_text(path, text)


def parse_typst_text(source: str | Path, text: str) -> Document:
    path = Path(source)
    return TypstParser(path, text).parse()
