"""Paragraph, run, and field reading for the forward converter."""

from __future__ import annotations

import re
from typing import Any, Iterable

from .docx_reader import _diff, _heading_level
from .model import Block, ImageBlock, PageBreak, Paragraph, Run
from .ooxml import attr, first, local_name, relationship_target


class ParagraphMixin:
    def _read_paragraph(self, paragraph: Any, part: Any) -> list[Block]:
        from .source_preservation import read_canvas
        native = read_canvas(self, paragraph)
        if native is not None:
            self.counts["canvases"] += 1
            return [native]
        style_id = self.styles.paragraph_style_id(paragraph)
        style_name = self.styles.style_name(style_id) or "Normal"
        self.used_style_ids.add(style_id)
        num_ilvl, num_id = self.styles.numbering_properties(paragraph)
        numbering_level = self.numbering.level(num_id, num_ilvl or 0) if num_id else None
        ppr = self.styles.paragraph_properties(
            paragraph, numbering_level.ppr if numbering_level else None
        )
        style_ppr = self.styles.style_paragraph_properties(style_id)
        direct_ppr = _diff(ppr, style_ppr)

        heading_level = _heading_level(
            style_name, style_id, style_ppr.get("outline_level")
        )
        marker = None
        list_kind = None
        number_id = None
        if num_id and numbering_level is not None and numbering_level.fmt not in {"none"}:
            marker_text, marker_level = self.numbering.next(num_id, num_ilvl or 0)
            if marker_level is not None:
                marker = marker_text
                list_kind = "bullet" if marker_level.fmt == "bullet" else "number"
                number_id = num_id
        elif num_id:
            self.warn(
                "numbering_none",
                "skipped",
                f"paragraph:{self.counts['paragraphs']}",
                source_type="numbering",
                fallback="ordinary paragraph",
                reason="numbering format is 'none'",
            )

        runs, bookmarks = self._read_inline(paragraph, part, style_id)
        blocks: list[Block] = []
        current: list[Run] = []
        for run in runs:
            if run.kind in {"page-break", "column-break"}:
                if current or not self.keep_empty:
                    blocks.append(
                        self._paragraph_block(
                            current,
                            style_name=style_name,
                            ppr=direct_ppr,
                            heading_level=heading_level,
                            list_kind=list_kind,
                            marker=marker,
                            number_id=number_id,
                            list_level=num_ilvl or 0,
                            style_id=style_id,
                            bookmarks=bookmarks,
                        )
                    )
                blocks.append(PageBreak())
                current = []
                if run.kind == "column-break":
                    self.warn(
                        "column_break",
                        "approximated",
                        f"paragraph:{self.counts['paragraphs']}",
                        source_type="column break",
                        fallback="page break",
                        reason="Typst columns do not expose Word's column-break control",
                    )
            else:
                current.append(run)
        if current or not self.keep_empty:
            blocks.append(
                self._paragraph_block(
                    current,
                    style_name=style_name,
                    ppr=direct_ppr,
                    heading_level=heading_level,
                    list_kind=list_kind,
                    marker=marker,
                    number_id=number_id,
                    list_level=num_ilvl or 0,
                    style_id=style_id,
                    bookmarks=bookmarks,
                )
            )
        self.counts["paragraphs"] += 1
        if heading_level is not None:
            self.counts["headings"] += 1
        if list_kind:
            self.counts["list_items"] += 1
        return blocks

    def _paragraph_block(
        self,
        runs: list[Run],
        *,
        style_name: str,
        ppr: dict[str, Any],
        heading_level: int | None,
        list_kind: str | None,
        marker: str | None,
        number_id: int | None,
        list_level: int,
        style_id: str | None,
        bookmarks: list[str],
    ) -> Block:
        image_runs = [run for run in runs if run.kind == "image"]
        has_text = any(run.text for run in runs if run.kind in {"text", "field", "footnote"})
        if len(image_runs) == 1 and len(runs) == 1 and not has_text:
            image = image_runs[0].image
            if image is not None and self.extract_media and image.path:
                self.counts["images"] += 1
                return ImageBlock(image=image, align=image.align or "center")
        return Paragraph(
            runs=runs,
            style=style_name,
            align=ppr.get("align"),
            list_kind=list_kind,
            list_level=list_level,
            list_style=style_name
            if style_name.lower().startswith("list")
            else (
                "List Bullet"
                if list_kind == "bullet"
                else "List Number"
                if list_kind == "number"
                else None
            ),
            marker=marker,
            number_id=number_id,
            indent_left=ppr.get("left_indent"),
            indent_right=ppr.get("right_indent"),
            first_line_indent=ppr.get("first_line_indent"),
            space_before=ppr.get("space_before"),
            space_after=ppr.get("space_after"),
            line_spacing=ppr.get("line_spacing"),
            line_spacing_pt=ppr.get("line_spacing_pt"),
            keep_with_next=ppr.get("keep_with_next"),
            keep_lines=ppr.get("keep_lines"),
            page_break_before=ppr.get("page_break_before"),
            heading_level=heading_level,
            outline_level=ppr.get("outline_level"),
            bookmarks=bookmarks,
        )

    # -- inline content --------------------------------------------------

    def _read_inline(
        self, paragraph: Any, part: Any, style_id: str
    ) -> tuple[list[Run], list[str]]:
        bookmarks: list[str] = []
        runs: list[Run] = []
        field: dict[str, Any] | None = None

        def finish_field(state: dict[str, Any]) -> list[Run]:
            instruction = state.get("instruction", "")
            result: list[Run] = state.get("result", [])
            command = (
                instruction.strip().split(" ", 1)[0] if instruction.strip() else ""
            ).upper()
            if command in {"PAGE", "NUMPAGES", "SECTIONPAGES"}:
                self.counts["fields"] += 1
                return [Run(kind="field", field=command, text="")]
            if command == "HYPERLINK":
                match = re.search(r'"([^"]+)"', instruction)
                target = match.group(1) if match else ""
                for run in result:
                    if target:
                        run.link = target
                if not result and target:
                    result = [Run(text=target, link=target)]
                self.counts["links"] += 1
                return result
            if command in {"TOC", "REF", "PAGEREF", "SEQ", "DOCPROPERTY", "STYLEREF"}:
                self.warn(
                    "field_" + command.lower(),
                    "approximated",
                    f"paragraph:{self.counts['paragraphs']}",
                    source_type=f"{command} field",
                    fallback="cached field result emitted as text",
                    reason="the reverse converter does not recompute Word fields",
                    expected_loss="The field becomes static text",
                )
                return result or [Run(text=instruction)]
            if command:
                self.warn(
                    "field",
                    "approximated",
                    f"paragraph:{self.counts['paragraphs']}",
                    source_type=f"{command} field",
                    fallback="cached result emitted as text",
                    reason="Typst has no Word field equivalent",
                )
            return result or ([Run(text=instruction)] if instruction else [])

        def walk(nodes: Iterable[Any], link: str | None = None, anchor: str | None = None) -> None:
            nonlocal field
            for node in nodes:
                name = local_name(node)
                if name == "r":
                    fld = first(node, "./w:fldChar")
                    if fld is not None:
                        kind = attr(fld, "w:fldCharType")
                        if kind == "begin":
                            field = {"instruction": "", "result": [], "phase": "instruction"}
                        elif kind == "separate" and field is not None:
                            field["phase"] = "result"
                        elif kind == "end" and field is not None:
                            finished = finish_field(field)
                            field = None
                            runs.extend(finished)
                        continue
                    instr = first(node, "./w:instrText")
                    if (
                        instr is not None
                        and field is not None
                        and field.get("phase") == "instruction"
                    ):
                        field["instruction"] = field.get("instruction", "") + (
                            instr.text or ""
                        )
                        continue
                    run_list = self._read_run(
                        node, part, style_id, link=link, anchor=anchor
                    )
                    if field is not None and field.get("phase") == "result":
                        field.setdefault("result", []).extend(run_list)
                    else:
                        runs.extend(run_list)
                elif name == "hyperlink":
                    rel_id = node.get(
                        "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
                    )
                    target, _external = relationship_target(part, rel_id)
                    node_anchor = attr(node, "w:anchor")
                    walk(node, link=target or link, anchor=node_anchor or anchor)
                    self.counts["links"] += 1
                elif name in {"ins", "moveTo"}:
                    self.warn(
                        "tracked_change",
                        "approximated",
                        f"paragraph:{self.counts['paragraphs']}",
                        source_type="tracked insertion",
                        fallback="inserted content included; change metadata dropped",
                        reason="Typst source records accepted content, not revision history",
                        expected_loss="Tracked-change metadata is lost",
                    )
                    walk(node, link, anchor)
                elif name in {"del", "moveFrom"}:
                    self.warn(
                        "tracked_deletion",
                        "skipped",
                        f"paragraph:{self.counts['paragraphs']}",
                        source_type="tracked deletion",
                        fallback="deleted content omitted",
                        reason="the Typst project represents the accepted document",
                        expected_loss="Deleted content and revision history are lost",
                    )
                elif name in {"smartTag", "customXml", "sdt"}:
                    content = first(node, "./w:sdtContent")
                    walk(content if content is not None else node, link, anchor)
                elif name == "fldSimple":
                    instruction = attr(node, "w:instr", "") or ""
                    before = len(runs)
                    walk(node, link, anchor)
                    temp = runs[before:]
                    if temp:
                        del runs[before:]
                    state = {
                        "instruction": instruction,
                        "result": temp,
                        "phase": "result",
                    }
                    runs.extend(finish_field(state))
                elif name == "bookmarkStart":
                    bookmark = attr(node, "w:name")
                    if bookmark and not bookmark.startswith("_"):
                        bookmarks.append(bookmark)
                elif name in {"commentRangeStart", "commentRangeEnd", "commentReference"}:
                    self.warn(
                        "comment",
                        "skipped",
                        f"paragraph:{self.counts['paragraphs']}",
                        source_type="comment",
                        fallback="comment omitted",
                        reason="the supported Typst subset has no comment model",
                        expected_loss="Review comments and replies are lost",
                    )
                elif name in {
                    "proofErr",
                    "bookmarkEnd",
                    "permStart",
                    "permEnd",
                    "moveFromRangeStart",
                    "moveFromRangeEnd",
                    "moveToRangeStart",
                    "moveToRangeEnd",
                }:
                    continue
                elif name in {"oMath", "oMathPara"}:
                    text = "".join(node.itertext())
                    self.warn(
                        "math",
                        "approximated",
                        f"paragraph:{self.counts['paragraphs']}",
                        source_type="Office Math",
                        fallback="linear text emitted",
                        reason="Office Math is not translated to Typst math",
                        expected_loss="Equation layout and symbols may be lost",
                    )
                    if text:
                        runs.append(Run(text=text))
                elif name == "altChunk":
                    continue
                else:
                    walk(node, link, anchor)

        walk(paragraph)
        if field is not None:
            runs.extend(finish_field(field))
        return _merge_runs(runs), bookmarks

    def _read_run(
        self,
        run: Any,
        part: Any,
        paragraph_style_id: str,
        *,
        link: str | None = None,
        anchor: str | None = None,
    ) -> list[Run]:
        props = self.styles.run_properties(run, paragraph_style_id)
        base = self.styles.style_run_properties(paragraph_style_id)
        direct = _diff(props, base)
        results: list[Run] = []

        def add(text: str = "", **kwargs: Any) -> None:
            data = dict(direct)
            data.update(kwargs)
            data["text"] = text
            if link:
                data["link"] = link
            if anchor:
                data["anchor"] = anchor
            results.append(Run(**data))

        for node in run:
            name = local_name(node)
            if name == "t":
                add(node.text or "")
            elif name == "tab":
                add("", kind="tab")
            elif name == "br":
                kind = attr(node, "w:type")
                if kind == "page":
                    add("", kind="page-break")
                elif kind == "column":
                    add("", kind="column-break")
                else:
                    add("", kind="break")
            elif name == "cr":
                add("", kind="break")
            elif name == "noBreakHyphen":
                add("\u2011")
            elif name == "softHyphen":
                add("\u00ad")
            elif name == "sym":
                char = attr(node, "w:char")
                font = attr(node, "w:font")
                if char:
                    try:
                        value = int(char, 16)
                        if value >= 0xF000:
                            value -= 0xF000
                        add(chr(value))
                    except ValueError:
                        add("")
                if font:
                    self.warn(
                        "symbol_font",
                        "approximated",
                        f"paragraph:{self.counts['paragraphs']}",
                        source_type="symbol",
                        fallback="Unicode codepoint emitted",
                        reason=f"symbol font {font!r} is not available in Typst",
                    )
            elif name == "drawing":
                results.extend(self._read_drawing(node, part))
            elif name == "pict":
                results.extend(self._read_vml(node, part))
            elif name == "object":
                self.warn(
                    "ole_object",
                    "unsupported",
                    f"paragraph:{self.counts['paragraphs']}",
                    source_type="OLE object",
                    fallback="omitted",
                    reason="embedded OLE objects are not converted",
                    expected_loss="The embedded object and its preview are missing",
                )
            elif name == "footnoteReference":
                note_id = attr(node, "w:id")
                note = self._footnotes.get(note_id or "")
                number = self.counts["footnotes"] + 1
                self.counts["footnotes"] += 1
                footnote = None
                if note:
                    from .model import Footnote

                    footnote = Footnote(number=number, blocks=note)
                add("", kind="footnote", footnote=footnote)
                self.warn(
                    "footnote",
                    "approximated",
                    f"paragraph:{self.counts['paragraphs']}",
                    source_type="footnote",
                    fallback="Typst footnote; reverse export uses a superscript marker",
                    reason="python-docx does not expose a footnote writer",
                    expected_loss="Footnote text is preserved but may not be a native Word footnote",
                )
            elif name == "endnoteReference":
                self.warn(
                    "endnote",
                    "unsupported",
                    f"paragraph:{self.counts['paragraphs']}",
                    source_type="endnote",
                    fallback="marker omitted",
                    reason="endnotes are outside the supported subset",
                    expected_loss="Endnote content is missing",
                )
            elif name == "commentReference":
                self.warn(
                    "comment",
                    "skipped",
                    f"paragraph:{self.counts['paragraphs']}",
                    source_type="comment",
                    fallback="comment omitted",
                    reason="the supported Typst subset has no comment model",
                    expected_loss="Review comments are lost",
                )
            elif name in {"lastRenderedPageBreak", "rPr", "fldChar", "instrText"}:
                continue
            elif name == "tbl":
                continue
            else:
                text = node.text
                if text:
                    add(text)
        return results


def _merge_runs(runs: list[Run]) -> list[Run]:
    merged: list[Run] = []
    for run in runs:
        if run.kind != "text" or not run.text:
            merged.append(run)
            continue
        if merged and merged[-1].kind == "text" and _same_format(merged[-1], run):
            merged[-1].text += run.text
        else:
            merged.append(run)
    return merged


def _same_format(left: Run, right: Run) -> bool:
    fields = (
        "bold",
        "italic",
        "underline",
        "strike",
        "small_caps",
        "all_caps",
        "superscript",
        "subscript",
        "color",
        "highlight",
        "font",
        "east_asia_font",
        "size",
        "code",
        "link",
        "anchor",
        "style_name",
    )
    return all(getattr(left, field) == getattr(right, field) for field in fields)
