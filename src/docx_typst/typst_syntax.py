"""Small, dependency-free helpers for the static Typst subset.

This is not a Typst interpreter.  It reads literal values, function calls and
content blocks so the reverse converter can understand files written in the
documented subset without evaluating arbitrary Typst code.
"""

from __future__ import annotations

from dataclasses import dataclass, field
import json
import math
import re
from typing import Any


# ---------------------------------------------------------------------------
# Value types
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class UnitValue:
    value: float
    unit: str = "pt"

    def to_pt(self) -> float:
        factors = {
            "pt": 1.0,
            "mm": 72.0 / 25.4,
            "cm": 72.0 / 2.54,
            "in": 72.0,
            "em": 11.0,
            "rem": 11.0,
            "pc": 12.0,
            "px": 0.75,
        }
        return self.value * factors.get(self.unit, 1.0)

    def __str__(self) -> str:
        return f"{format_number(self.value)}{self.unit}"


@dataclass(frozen=True)
class RawExpr:
    text: str

    def __str__(self) -> str:
        return self.text


@dataclass(frozen=True)
class Identifier:
    name: str

    def __str__(self) -> str:
        return self.name


@dataclass(frozen=True)
class Content:
    text: str


@dataclass
class TypstCall:
    name: str
    positional: list[Any] = field(default_factory=list)
    named: dict[str, Any] = field(default_factory=dict)
    body: Content | None = None
    raw: str = ""

    def arg(self, name: str, default: Any = None) -> Any:
        return self.named.get(name, default)


# ---------------------------------------------------------------------------
# Basic text helpers
# ---------------------------------------------------------------------------


def format_number(value: float) -> str:
    if abs(value - round(value)) < 1e-9:
        return str(int(round(value)))
    return f"{value:.4f}".rstrip("0").rstrip(".")


def format_length(value: float, unit: str = "pt") -> str:
    return f"{format_number(value)}{unit}"


def typst_string(value: str) -> str:
    """Quote a Python string for Typst source."""
    return json.dumps(value, ensure_ascii=False)


def typst_escape(text: str) -> str:
    """Escape literal text for Typst markup.

    Characters are escaped everywhere; line-leading block markers are handled
    separately by the renderer because paragraphs are emitted on their own
    indented lines inside content blocks.
    """
    out: list[str] = []
    for index, char in enumerate(text):
        if char in "\\#$*_`[]<>@~":
            out.append("\\" + char)
        elif char == "\u00a0":
            out.append("~")
        else:
            out.append(char)
    result = "".join(out)
    # A literal hyphen, plus, slash or equals at the start of a line would be
    # parsed as a Typst list, term or heading marker inside a content block.
    lines = result.split("\n")
    for i, line in enumerate(lines):
        stripped = line.lstrip(" \t")
        if stripped[:1] in {"-", "+", "=", "/"}:
            prefix_len = len(line) - len(stripped)
            lines[i] = line[:prefix_len] + "\\" + stripped
    return "\n".join(lines)


def balanced(text: str, start: int, left: str = "(", right: str = ")") -> tuple[str, int]:
    """Return the contents of a balanced pair and the index after it."""
    if start >= len(text) or text[start] != left:
        raise ValueError(f"expected {left!r} at {start}")
    depth = 1
    quote: str | None = None
    escaped = False
    comment_line = False
    comment_block = 0
    raw_ticks = 0
    i = start + 1
    while i < len(text):
        char = text[i]
        if comment_line:
            if char == "\n":
                comment_line = False
            i += 1
            continue
        if comment_block:
            if char == "/" and i + 1 < len(text) and text[i + 1] == "*":
                comment_block += 1
                i += 2
                continue
            if char == "*" and i + 1 < len(text) and text[i + 1] == "/":
                comment_block -= 1
                i += 2
                continue
            i += 1
            continue
        if raw_ticks:
            if char == "`":
                count = 1
                while i + count < len(text) and text[i + count] == "`":
                    count += 1
                if count >= raw_ticks:
                    i += raw_ticks
                    raw_ticks = 0
                    continue
            i += 1
            continue
        if quote:
            if escaped:
                escaped = False
            elif char == "\\":
                escaped = True
            elif char == quote:
                quote = None
            i += 1
            continue
        if char in {'"', "'"}:
            quote = char
        elif char == "/" and i + 1 < len(text) and text[i + 1] == "/":
            comment_line = True
            i += 2
            continue
        elif char == "/" and i + 1 < len(text) and text[i + 1] == "*":
            comment_block = 1
            i += 2
            continue
        elif char == "`":
            count = 1
            while i + count < len(text) and text[i + count] == "`":
                count += 1
            raw_ticks = count
            i += count
            continue
        elif char == left:
            depth += 1
        elif char == right:
            depth -= 1
            if depth == 0:
                return text[start + 1 : i], i + 1
        i += 1
    raise ValueError(f"unbalanced {left!r} at {start}")


def split_top_level(text: str, separator: str = ",") -> list[str]:
    parts: list[str] = []
    start = 0
    i = 0
    stack: list[str] = []
    opens = {"(": ")", "[": "]", "{": "}"}
    quote: str | None = None
    escaped = False
    comment_line = False
    comment_block = 0
    raw_ticks = 0
    while i < len(text):
        char = text[i]
        if comment_line:
            if char == "\n":
                comment_line = False
            i += 1
            continue
        if comment_block:
            if char == "/" and i + 1 < len(text) and text[i + 1] == "*":
                comment_block += 1
                i += 2
                continue
            if char == "*" and i + 1 < len(text) and text[i + 1] == "/":
                comment_block -= 1
                i += 2
                continue
            i += 1
            continue
        if raw_ticks:
            if char == "`":
                count = 1
                while i + count < len(text) and text[i + count] == "`":
                    count += 1
                if count >= raw_ticks:
                    i += raw_ticks
                    raw_ticks = 0
                    continue
            i += 1
            continue
        if quote:
            if escaped:
                escaped = False
            elif char == "\\":
                escaped = True
            elif char == quote:
                quote = None
            i += 1
            continue
        if char in {'"', "'"}:
            quote = char
        elif char == "/" and i + 1 < len(text) and text[i + 1] == "/":
            comment_line = True
            i += 2
            continue
        elif char == "/" and i + 1 < len(text) and text[i + 1] == "*":
            comment_block = 1
            i += 2
            continue
        elif char == "`":
            count = 1
            while i + count < len(text) and text[i + count] == "`":
                count += 1
            raw_ticks = count
            i += count
            continue
        elif char in opens:
            stack.append(opens[char])
        elif stack and char == stack[-1]:
            stack.pop()
        elif char == separator and not stack:
            parts.append(text[start:i].strip())
            start = i + 1
        i += 1
    tail = text[start:].strip()
    if tail:
        parts.append(tail)
    return parts


def strip_comments(text: str) -> str:
    """Remove Typst line and nested block comments, preserving strings/raw."""
    out: list[str] = []
    i = 0
    quote: str | None = None
    escaped = False
    comment_line = False
    comment_block = 0
    raw_ticks = 0
    while i < len(text):
        char = text[i]
        if comment_line:
            if char == "\n":
                comment_line = False
                out.append(char)
            i += 1
            continue
        if comment_block:
            if char == "/" and i + 1 < len(text) and text[i + 1] == "*":
                comment_block += 1
                i += 2
                continue
            if char == "*" and i + 1 < len(text) and text[i + 1] == "/":
                comment_block -= 1
                i += 2
                continue
            i += 1
            continue
        if raw_ticks:
            out.append(char)
            if char == "`":
                count = 1
                while i + count < len(text) and text[i + count] == "`":
                    out.append("`")
                    count += 1
                if count >= raw_ticks:
                    raw_ticks = 0
                i += count
                continue
            i += 1
            continue
        if quote:
            out.append(char)
            if escaped:
                escaped = False
            elif char == "\\":
                escaped = True
            elif char == quote:
                quote = None
            i += 1
            continue
        if char in {'"', "'"}:
            quote = char
            out.append(char)
        elif char == "/" and i + 1 < len(text) and text[i + 1] == "/":
            comment_line = True
            i += 2
            continue
        elif char == "/" and i + 1 < len(text) and text[i + 1] == "*":
            comment_block = 1
            i += 2
            continue
        elif char == "`":
            count = 1
            while i + count < len(text) and text[i + count] == "`":
                count += 1
            raw_ticks = count
            out.append("`" * count)
            i += count
            continue
        else:
            out.append(char)
        i += 1
    return "".join(out)


# ---------------------------------------------------------------------------
# Literal value parser
# ---------------------------------------------------------------------------


NAMED_COLORS = {
    "black": "#000000",
    "white": "#FFFFFF",
    "red": "#FF0000",
    "green": "#00FF00",
    "blue": "#0000FF",
    "yellow": "#FFFF00",
    "orange": "#FFA500",
    "purple": "#800080",
    "gray": "#808080",
    "grey": "#808080",
    "silver": "#C0C0C0",
    "maroon": "#800000",
    "navy": "#000080",
    "teal": "#008080",
    "olive": "#808000",
    "aqua": "#00FFFF",
    "fuchsia": "#FF00FF",
    "lime": "#00FF00",
}


class ValueParser:
    def __init__(self, text: str, pos: int = 0):
        self.text = text
        self.pos = pos

    def skip(self) -> None:
        while self.pos < len(self.text):
            if self.text[self.pos].isspace():
                self.pos += 1
                continue
            if self.text.startswith("//", self.pos):
                end = self.text.find("\n", self.pos)
                self.pos = len(self.text) if end < 0 else end + 1
                continue
            if self.text.startswith("/*", self.pos):
                depth = 1
                self.pos += 2
                while self.pos < len(self.text) and depth:
                    if self.text.startswith("/*", self.pos):
                        depth += 1
                        self.pos += 2
                    elif self.text.startswith("*/", self.pos):
                        depth -= 1
                        self.pos += 2
                    else:
                        self.pos += 1
                continue
            break

    def peek(self) -> str:
        self.skip()
        return self.text[self.pos : self.pos + 1]

    def parse(self) -> Any:
        value = self.parse_value()
        self.skip()
        return value

    def parse_value(self) -> Any:
        self.skip()
        if self.pos >= len(self.text):
            return RawExpr("")
        char = self.text[self.pos]
        if char == '"':
            return self.parse_string()
        if char == "[":
            inner, self.pos = balanced(self.text, self.pos, "[", "]")
            return Content(inner)
        if char == "(":
            return self.parse_group()
        if char.isdigit() or (char in "+-" and self.pos + 1 < len(self.text) and self.text[self.pos + 1].isdigit()):
            return self.parse_number()
        if char == "#":
            self.pos += 1
            return self.parse_value()
        if char == "<":
            end = self.text.find(">", self.pos)
            if end >= 0:
                name = self.text[self.pos + 1 : end]
                self.pos = end + 1
                return Identifier(name)
        return self.parse_ident_or_call()

    def parse_string(self) -> str:
        start = self.pos
        self.pos += 1
        out: list[str] = []
        while self.pos < len(self.text):
            char = self.text[self.pos]
            if char == '"':
                self.pos += 1
                return "".join(out)
            if char == "\\":
                self.pos += 1
                if self.pos >= len(self.text):
                    break
                esc = self.text[self.pos]
                mapping = {
                    "n": "\n",
                    "t": "\t",
                    "r": "\r",
                    '"': '"',
                    "\\": "\\",
                    "[": "[",
                    "]": "]",
                }
                if esc == "u" and self.pos + 1 < len(self.text) and self.text[self.pos + 1] == "{":
                    end = self.text.find("}", self.pos + 2)
                    if end >= 0:
                        try:
                            out.append(chr(int(self.text[self.pos + 2 : end], 16)))
                        except ValueError:
                            pass
                        self.pos = end + 1
                        continue
                out.append(mapping.get(esc, esc))
                self.pos += 1
                continue
            out.append(char)
            self.pos += 1
        self.pos = start + 1
        return RawExpr(self.text[start:])

    def parse_number(self) -> Any:
        match = re.match(r"[-+]?\d+(?:\.\d+)?", self.text[self.pos :])
        if not match:
            return RawExpr(self.text[self.pos : self.pos + 1])
        raw = match.group()
        self.pos += len(raw)
        number: int | float = float(raw) if "." in raw else int(raw)
        unit_match = re.match(r"[a-zA-Z%]+", self.text[self.pos :])
        if unit_match:
            unit = unit_match.group()
            self.pos += len(unit)
            if unit == "%":
                return number / 100.0
            return UnitValue(float(number), unit)
        return number

    def parse_group(self) -> Any:
        inner, self.pos = balanced(self.text, self.pos, "(", ")")
        stripped = inner.strip()
        if not stripped:
            return []
        if stripped == ":":
            return {}
        parts = split_top_level(stripped)
        named: dict[str, Any] = {}
        positional: list[Any] = []
        is_dict = False
        for part in parts:
            quoted = re.match(r'^"((?:[^"\\]|\\.)*)"\s*:\s*(.*)$', part, re.S)
            if quoted:
                is_dict = True
                key = ValueParser('"' + quoted.group(1) + '"').parse_string()
                if not isinstance(key, str):
                    key = quoted.group(1)
                named[key] = ValueParser(quoted.group(2)).parse()
                continue
            match = re.match(r"^([A-Za-z_][\w-]*)\s*:\s*(.*)$", part, re.S)
            if match and not part.lstrip().startswith('"'):
                is_dict = True
                named[match.group(1)] = ValueParser(match.group(2)).parse()
            else:
                positional.append(ValueParser(part).parse())
        if is_dict:
            return named
        if len(positional) == 1 and not stripped.endswith(","):
            return positional[0]
        return positional

    def parse_ident_or_call(self) -> Any:
        match = re.match(r"[A-Za-z_][\w-]*", self.text[self.pos :])
        if not match:
            start = self.pos
            self.pos += 1
            return RawExpr(self.text[start : self.pos])
        name = match.group()
        self.pos += len(name)
        self.skip()
        if self.pos < len(self.text) and self.text[self.pos] == "(":
            inner, self.pos = balanced(self.text, self.pos, "(", ")")
            args, kwargs = parse_call_arguments(inner)
            raw = f"{name}({inner})"
            lower = name.lower()
            if lower in {"rgb", "color.rgb"}:
                return color_from_args(args, kwargs)
            if lower in {"luma", "color.luma"}:
                return color_from_luma(args, kwargs)
            if lower in {"pt", "mm", "cm", "in", "em", "rem", "pc", "px"}:
                if args:
                    return UnitValue(float(args[0]), lower)
            if lower in {"black", "white"} and not args:
                return NAMED_COLORS[lower]
            return TypstCall(name=name, positional=args, named=kwargs, raw=raw)
        lower = name.lower()
        if lower == "true":
            return True
        if lower == "false":
            return False
        if lower == "none":
            return None
        if lower == "auto":
            return Identifier("auto")
        if lower in NAMED_COLORS:
            return NAMED_COLORS[lower]
        return Identifier(name)


def parse_value(text: str) -> Any:
    parser = ValueParser(text)
    value = parser.parse()
    parser.skip()
    if parser.pos < len(text):
        # Preserve an unsupported trailing expression rather than silently
        # discarding it.
        return RawExpr(text.strip())
    return value


def parse_call_arguments(inner: str) -> tuple[list[Any], dict[str, Any]]:
    positional: list[Any] = []
    named: dict[str, Any] = {}
    for part in split_top_level(inner):
        if not part:
            continue
        match = re.match(r"^([A-Za-z_][\w-]*)\s*:\s*(.*)$", part, re.S)
        if match and not part.lstrip().startswith('"'):
            named[match.group(1)] = parse_value(match.group(2))
        else:
            positional.append(parse_value(part))
    return positional, named


def color_from_args(args: list[Any], kwargs: dict[str, Any]) -> str:
    if args:
        first = args[0]
        if isinstance(first, str) and first.startswith("#"):
            digits = first[1:]
            if len(digits) == 3:
                digits = "".join(c * 2 for c in digits)
            if len(digits) == 6:
                return "#" + digits.upper()
            if len(digits) == 8:
                return "#" + digits.upper()
        if len(args) >= 3 and all(isinstance(value, (int, float)) for value in args[:3]):
            values = [max(0, min(255, round(float(v) * (255 if float(v) <= 1 else 1)))) for v in args[:3]]
            alpha = "" if len(args) < 4 else f"{max(0, min(255, round(float(args[3]) * 255))):02X}"
            return "#" + "".join(f"{value:02X}" for value in values) + alpha
    return "#000000"


def color_from_luma(args: list[Any], kwargs: dict[str, Any]) -> str:
    if not args:
        return "#000000"
    value = args[0]
    if isinstance(value, UnitValue):
        value = value.value
    if not isinstance(value, (int, float)):
        return "#000000"
    if value <= 1:
        value *= 255
    gray = max(0, min(255, round(float(value))))
    return f"#{gray:02X}{gray:02X}{gray:02X}"


# ---------------------------------------------------------------------------
# Function call extraction
# ---------------------------------------------------------------------------


def parse_call_at(text: str, pos: int) -> TypstCall | None:
    match = re.match(r"([A-Za-z_][\w-]*)", text[pos:])
    if not match:
        return None
    name = match.group(1)
    index = pos + len(name)
    while index < len(text) and text[index] in " \t":
        index += 1
    positional: list[Any] = []
    named: dict[str, Any] = {}
    start = pos
    if index < len(text) and text[index] == "(":
        try:
            inner, index = balanced(text, index, "(", ")")
        except ValueError:
            return None
        positional, named = parse_call_arguments(inner)
    body: Content | None = None
    if index < len(text) and text[index] == "[":
        try:
            inner, index = balanced(text, index, "[", "]")
        except ValueError:
            return None
        body = Content(inner)
    return TypstCall(name=name, positional=positional, named=named, body=body, raw=text[start:index])


def iter_calls(text: str, name: str) -> list[TypstCall]:
    calls: list[TypstCall] = []
    pattern = re.compile(r"(?<![\w-])#?" + re.escape(name) + r"\s*\(")
    for match in pattern.finditer(text):
        call = parse_call_at(text, match.start() + (1 if match.group().startswith("#") else 0))
        if call is not None:
            calls.append(call)
    return calls


# ---------------------------------------------------------------------------
# Color / length conversion helpers
# ---------------------------------------------------------------------------


def color_to_rgb(value: Any) -> tuple[int, int, int] | None:
    if isinstance(value, str):
        digits = value.lstrip("#")
        if len(digits) in (6, 8):
            try:
                return tuple(int(digits[i : i + 2], 16) for i in (0, 2, 4))  # type: ignore[return-value]
            except ValueError:
                return None
    if isinstance(value, Identifier) and value.name in NAMED_COLORS:
        digits = NAMED_COLORS[value.name].lstrip("#")
        return tuple(int(digits[i : i + 2], 16) for i in (0, 2, 4))  # type: ignore[return-value]
    return None


def length_to_pt(value: Any, default: float = 0.0) -> float:
    if isinstance(value, UnitValue):
        return value.to_pt()
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        match = re.search(r"[-+]?\d+(?:\.\d+)?", value)
        if match:
            return float(match.group())
    return default


def value_text(value: Any, default: str = "") -> str:
    if value is None:
        return default
    if isinstance(value, str):
        return value
    if isinstance(value, (Identifier, RawExpr)):
        return str(value)
    return str(value)
