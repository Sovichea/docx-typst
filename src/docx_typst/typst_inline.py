"""Parse the supported inline Typst markup subset into model runs."""

from __future__ import annotations

from typing import Any

from .model import Document, ImageRef, Run
from .typst_syntax import (
    Identifier,
    TypstCall,
    color_to_rgb,
    length_to_pt,
    parse_call_at,
    value_text,
)
from .typst_values import image_from_call


class InlineParser:
    def __init__(self, document: Document, location: str = "body"):
        self.document = document
        self.location = location

    def parse(self, text: str) -> list[Run]:
        runs: list[Run] = []
        buffer: list[str] = []

        def flush() -> None:
            if buffer:
                runs.append(Run(text="".join(buffer)))
                buffer.clear()

        index = 0
        while index < len(text):
            char = text[index]
            if char == "\\" and index + 1 < len(text):
                next_char = text[index + 1]
                if next_char in "\r\n":
                    flush()
                    runs.append(Run(kind="break"))
                    if next_char == "\r" and index + 2 < len(text) and text[index + 2] == "\n":
                        index += 3
                    else:
                        index += 2
                    continue
                buffer.append(next_char)
                index += 2
                continue
            if char == "*":
                close = _find_delimiter(text, index + 1, "*")
                if close >= 0:
                    flush()
                    inner = self.parse(text[index + 1 : close])
                    _apply(inner, bold=True)
                    runs.extend(inner)
                    index = close + 1
                    continue
            if char == "_":
                close = _find_delimiter(text, index + 1, "_")
                if close >= 0:
                    flush()
                    inner = self.parse(text[index + 1 : close])
                    _apply(inner, italic=True)
                    runs.extend(inner)
                    index = close + 1
                    continue
            if char == "`":
                ticks = 1
                while index + ticks < len(text) and text[index + ticks] == "`":
                    ticks += 1
                close = _find_code_end(text, index + ticks, ticks)
                if close >= 0:
                    flush()
                    runs.append(Run(text=text[index + ticks : close], code=True))
                    index = close + ticks
                    continue
            if char == "$":
                close = _find_delimiter(text, index + 1, "$")
                if close >= 0:
                    flush()
                    runs.append(Run(text=text[index + 1 : close], kind="math"))
                    self.document.warn(
                        "typst_math",
                        "approximated",
                        self.location,
                        source_type="Typst math",
                        fallback="equation text emitted as plain text",
                        reason="the reverse converter does not evaluate Typst math",
                    )
                    index = close + 1
                    continue
            if char == "#":
                call = parse_call_at(text, index + 1)
                if call is not None:
                    flush()
                    consumed = len(call.raw)
                    handled = self._call(call, runs)
                    if not handled:
                        if call.body is not None:
                            runs.extend(self.parse(call.body.text))
                        else:
                            buffer.append(call.raw)
                    index += 1 + consumed
                    continue
            if char == "\n" or char == " " or char == "\t" or char == "\r":
                if buffer:
                    if not buffer[-1].isspace():
                        buffer.append(" ")
                elif runs:
                    buffer.append(" ")
                index += 1
                continue
            buffer.append(char)
            index += 1
        flush()
        _trim(runs)
        return _merge(runs)

    # -- calls -----------------------------------------------------------

    def _call(self, call: TypstCall, runs: list[Run]) -> bool:
        name = call.name
        if name in {"linebreak", "parbreak"}:
            runs.append(Run(kind="break"))
            return True
        if name == "pagebreak":
            runs.append(Run(kind="page-break"))
            return True
        if name == "colbreak":
            runs.append(Run(kind="column-break"))
            return True
        if name == "h":
            amount = length_to_pt(call.positional[0]) if call.positional else 0.0
            runs.append(Run(kind="tab" if amount > 6.0 else "text", text="" if amount > 6.0 else " "))
            return True
        if name in {"line", "hline"}:
            runs.append(Run(text="________"))
            return True
        if name in {"strong", "emph", "underline", "strike", "smallcaps", "sub", "super", "upper", "docx-allcaps"}:
            if call.body is None:
                return False
            inner = self.parse(call.body.text)
            if name == "strong":
                _apply(inner, bold=True)
            elif name == "emph":
                _apply(inner, italic=True)
            elif name == "underline":
                _apply(inner, underline=True)
            elif name == "strike":
                _apply(inner, strike=True)
            elif name == "smallcaps":
                _apply(inner, small_caps=True)
            elif name == "sub":
                _apply(inner, subscript=True)
            elif name == "super":
                _apply(inner, superscript=True)
            else:
                _apply(inner, all_caps=True)
            runs.extend(inner)
            return True
        if name == "text":
            if call.body is None:
                return False
            inner = self.parse(call.body.text)
            font = call.arg("font")
            size = call.arg("size")
            fill = call.arg("fill")
            weight = call.arg("weight")
            style = call.arg("style")
            kwargs: dict[str, Any] = {}
            if isinstance(font, str):
                kwargs["font"] = font
            if size is not None:
                kwargs["size"] = length_to_pt(size)
            color = color_to_rgb(fill)
            if color is not None:
                kwargs["color"] = "#" + "".join(f"{value:02X}" for value in color)
            if isinstance(weight, str):
                kwargs["bold"] = weight.lower() == "bold"
            if isinstance(style, str):
                kwargs["italic"] = style.lower() == "italic"
            _apply(inner, **kwargs)
            runs.extend(inner)
            return True
        if name == "highlight":
            if call.body is None:
                return False
            inner = self.parse(call.body.text)
            fill = color_to_rgb(call.arg("fill"))
            if fill is not None:
                _apply(inner, highlight="#" + "".join(f"{value:02X}" for value in fill))
            runs.extend(inner)
            return True
        if name == "link":
            if call.body is None:
                return False
            target: Any = call.positional[0] if call.positional else call.arg("dest")
            inner = self.parse(call.body.text)
            if isinstance(target, Identifier):
                _apply(inner, anchor=target.name)
            elif isinstance(target, str):
                _apply(inner, link=target)
            runs.extend(inner)
            return True
        if name == "docx-anchor-link":
            if call.body is None:
                return False
            anchor = value_text(call.positional[0] if call.positional else "")
            inner = self.parse(call.body.text)
            _apply(inner, anchor=anchor)
            runs.extend(inner)
            return True
        if name == "raw":
            if call.positional:
                value = call.positional[0]
                if isinstance(value, str):
                    runs.append(Run(text=value, code=True))
                    return True
            if call.body is not None:
                runs.append(Run(text=call.body.text, code=True))
                return True
            return False
        if name == "docx-inline-image":
            image = image_from_call(call)
            if image is not None:
                runs.append(Run(kind="image", image=image))
            return True
        if name == "docx-static":
            path = call.arg("path")
            if isinstance(path, str):
                width = call.arg("width")
                height = call.arg("height")
                runs.append(
                    Run(
                        kind="image",
                        image=ImageRef(
                            path=path,
                            width_pt=length_to_pt(width) if width is not None else None,
                            height_pt=length_to_pt(height) if height is not None else None,
                        ),
                    )
                )
                return True
            if call.body is not None:
                runs.extend(self.parse(call.body.text))
                return True
        if name == "footnote":
            if call.body is None:
                return False
            runs.append(Run(kind="footnote", text=call.body.text))
            return True
        if name == "docx-math":
            value = call.positional[0] if call.positional else ""
            runs.append(Run(text=value_text(value), kind="math"))
            return True
        if name in {"docx-page-number", "docx-page-count"}:
            runs.append(
                Run(kind="field", field="PAGE" if name == "docx-page-number" else "NUMPAGES")
            )
            return True
        if name in {"docx-page-number-padded", "docx-page-count-padded"}:
            runs.append(
                Run(
                    kind="field",
                    field="PAGE_PADDED"
                    if name == "docx-page-number-padded"
                    else "NUMPAGES_PADDED",
                )
            )
            return True
        if name == "docx-field":
            runs.append(Run(kind="field", field=value_text(call.positional[0] if call.positional else "")))
            return True
        if name in {"docx-allcaps", "docx-superscript", "docx-subscript"}:
            if call.body is None:
                return False
            inner = self.parse(call.body.text)
            if name == "docx-superscript":
                _apply(inner, superscript=True)
            elif name == "docx-subscript":
                _apply(inner, subscript=True)
            else:
                _apply(inner, all_caps=True)
            runs.extend(inner)
            return True
        if name in {"box", "block", "pad", "place", "docx-para", "docx-heading"}:
            if call.body is None:
                return False
            runs.extend(self.parse(call.body.text))
            return True
        if name == "docx-unsupported-image":
            return True
        return False


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _apply(runs: list[Run], **kwargs: Any) -> None:
    for run in runs:
        for key, value in kwargs.items():
            if value is None:
                continue
            if key in {"bold", "italic", "underline", "strike", "small_caps", "all_caps", "superscript", "subscript"}:
                current = getattr(run, key)
                if current is None:
                    setattr(run, key, value)
                elif value:
                    setattr(run, key, True)
            elif getattr(run, key) is None:
                setattr(run, key, value)


def _trim(runs: list[Run]) -> None:
    if runs and runs[0].kind == "text" and runs[0].text:
        runs[0].text = runs[0].text.lstrip()
    if runs and runs[-1].kind == "text" and runs[-1].text:
        runs[-1].text = runs[-1].text.rstrip()
    runs[:] = [run for run in runs if run.kind != "text" or run.text]


def _merge(runs: list[Run]) -> list[Run]:
    merged: list[Run] = []
    for run in runs:
        if (
            merged
            and run.kind == "text"
            and merged[-1].kind == "text"
            and run.text
            and _same_format(merged[-1], run)
        ):
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
    )
    return all(getattr(left, field) == getattr(right, field) for field in fields)


def _find_delimiter(text: str, start: int, delimiter: str) -> int:
    index = start
    while index < len(text):
        char = text[index]
        if char == "\\":
            index += 2
            continue
        if char == "#":
            call = parse_call_at(text, index + 1)
            if call is not None:
                index += 1 + len(call.raw)
                continue
        if char == delimiter:
            return index
        index += 1
    return -1


def _find_code_end(text: str, start: int, ticks: int) -> int:
    index = start
    while index < len(text):
        if text[index] == "`":
            count = 1
            while index + count < len(text) and text[index + count] == "`":
                count += 1
            if count >= ticks:
                return index
            index += count
            continue
        index += 1
    return -1
