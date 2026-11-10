"""WordprocessingML helpers for raw properties python-docx does not expose."""

from __future__ import annotations

from dataclasses import dataclass, field
import re
from typing import Any, Iterable, Iterator

from lxml import etree

from .model import PageSetup


W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
V = "urn:schemas-microsoft-com:vml"
O = "urn:schemas-microsoft-com:office:office"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
ASVG = "http://schemas.microsoft.com/office/drawing/2016/SVG/main"
WPS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
WPG = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
C = "http://schemas.openxmlformats.org/drawingml/2006/chart"
DGM = "http://schemas.openxmlformats.org/drawingml/2006/diagram"

NS = {
    "w": W,
    "r": R,
    "a": A,
    "wp": WP,
    "pic": PIC,
    "m": M,
    "v": V,
    "o": O,
    "mc": MC,
    "asvg": ASVG,
    "wps": WPS,
    "wpg": WPG,
    "c": C,
    "dgm": DGM,
}


def qn(tag: str) -> str:
    prefix, local = tag.split(":", 1)
    return f"{{{NS[prefix]}}}{local}"


def local_name(node: Any) -> str:
    if node is None:
        return ""
    tag = node.tag if hasattr(node, "tag") else str(node)
    if not isinstance(tag, str):
        return ""
    return tag.rsplit("}", 1)[-1]


def xpath(node: Any, path: str) -> list[Any]:
    if node is None:
        return []
    try:
        return node.xpath(path, namespaces=NS)
    except (AttributeError, etree.XPathEvalError):
        try:
            return node.xpath(path)
        except (AttributeError, etree.XPathEvalError):
            return []


def first(node: Any, path: str) -> Any | None:
    values = xpath(node, path)
    return values[0] if values else None


def attr(node: Any, name: str, default: str | None = None) -> str | None:
    if node is None:
        return default
    value = node.get(qn(name))
    if value is None:
        value = node.get(name)
    return default if value is None else value


def child(node: Any, name: str) -> Any | None:
    if node is None:
        return None
    return node.find(qn(name))


def on_off(node: Any, name: str) -> bool | None:
    element = child(node, name)
    if element is None:
        return None
    value = attr(element, "w:val")
    if value is None:
        return True
    return value.lower() not in {"0", "false", "off", "no", "none"}


def twips(value: str | None, default: float | None = None) -> float | None:
    if value is None:
        return default
    try:
        return float(value) / 20.0
    except ValueError:
        return default


def emu_to_pt(value: str | int | None, default: float | None = None) -> float | None:
    if value is None:
        return default
    try:
        return float(value) / 12700.0
    except (ValueError, TypeError):
        return default


def eighth_points(value: str | None, default: float | None = None) -> float | None:
    if value is None:
        return default
    try:
        return float(value) / 2.0
    except ValueError:
        return default


def half_points(value: str | None, default: float | None = None) -> float | None:
    if value is None:
        return default
    try:
        return float(value) / 2.0
    except ValueError:
        return default


def bool_value(value: str | None) -> bool | None:
    if value is None:
        return None
    return value.lower() not in {"0", "false", "off", "no", "none"}


# ---------------------------------------------------------------------------
# Theme
# ---------------------------------------------------------------------------


THEME_COLOR_ALIASES = {
    "dark1": "dk1",
    "light1": "lt1",
    "dark2": "dk2",
    "light2": "lt2",
    "text1": "dk1",
    "background1": "lt1",
    "text2": "dk2",
    "background2": "lt2",
    "hyperlink": "hlink",
    "followedHyperlink": "folHlink",
}


def _theme_color_value(node: Any) -> str | None:
    if node is None:
        return None
    for item in node:
        name = local_name(item)
        if name in {"srgbClr", "scrgbClr"}:
            value = item.get("val")
            if value:
                return "#" + value.upper()
        if name == "sysClr":
            value = item.get("lastClr")
            if value:
                return "#" + value.upper()
    return None


def apply_tint_shade(hex_color: str, tint: str | None = None, shade: str | None = None) -> str:
    digits = hex_color.lstrip("#")[:6]
    try:
        rgb = [int(digits[i : i + 2], 16) for i in (0, 2, 4)]
    except ValueError:
        return hex_color
    if tint:
        try:
            amount = int(tint, 16) / 255.0
        except ValueError:
            amount = 0.0
        rgb = [round(value + (255 - value) * amount) for value in rgb]
    if shade:
        try:
            amount = int(shade, 16) / 255.0
        except ValueError:
            amount = 1.0
        rgb = [round(value * amount) for value in rgb]
    return "#" + "".join(f"{max(0, min(255, value)):02X}" for value in rgb)


@dataclass
class Theme:
    colors: dict[str, str] = field(default_factory=dict)
    fonts: dict[str, str] = field(default_factory=dict)

    def color(self, name: str | None, tint: str | None = None, shade: str | None = None) -> str | None:
        if not name:
            return None
        key = THEME_COLOR_ALIASES.get(name, name)
        value = self.colors.get(key)
        if value is None:
            return None
        return apply_tint_shade(value, tint, shade)

    def font(self, name: str | None) -> str | None:
        if not name:
            return None
        key = name
        if key.startswith("major"):
            base = "major"
        elif key.startswith("minor"):
            base = "minor"
        else:
            return None
        if "EastAsia" in key:
            return self.fonts.get(base + "EastAsia") or self.fonts.get(base)
        if "Bidi" in key:
            return self.fonts.get(base + "Bidi") or self.fonts.get(base)
        return self.fonts.get(base)


def parse_theme(root: Any) -> Theme:
    theme = Theme()
    if root is None:
        return theme
    schemes = xpath(root, "//a:clrScheme")
    if schemes:
        for item in schemes[0]:
            name = local_name(item)
            theme.colors[name] = _theme_color_value(item) or ""
    for key in ("major", "minor"):
        nodes = xpath(root, f"//a:{key}Font")
        if nodes:
            for child_node in nodes[0]:
                child_name = local_name(child_node)
                face = child_node.get("typeface")
                if not face:
                    continue
                if child_name == "latin":
                    theme.fonts[key] = face
                elif child_name == "ea":
                    theme.fonts[key + "EastAsia"] = face
                elif child_name == "cs":
                    theme.fonts[key + "Bidi"] = face
    theme.fonts.setdefault("major", "Calibri Light")
    theme.fonts.setdefault("minor", "Calibri")
    return theme


# ---------------------------------------------------------------------------
# Style cascade
# ---------------------------------------------------------------------------


class StyleResolver:
    def __init__(self, styles_root: Any, theme: Theme):
        self.root = styles_root
        self.theme = theme
        self.by_id: dict[str, Any] = {}
        if styles_root is not None:
            for node in xpath(styles_root, "./w:style"):
                style_id = attr(node, "w:styleId")
                if style_id:
                    self.by_id[style_id] = node
        self.default_paragraph_id = ""
        for node in self.by_id.values():
            if attr(node, "w:type") == "paragraph" and attr(node, "w:default") in {"1", "true"}:
                self.default_paragraph_id = attr(node, "w:styleId") or ""
                break
        defaults = first(styles_root, "./w:docDefaults")
        self.default_rpr = first(defaults, "./w:rPrDefault/w:rPr")
        self.default_ppr = first(defaults, "./w:pPrDefault/w:pPr")

    def style_name(self, style_id: str | None) -> str | None:
        if not style_id:
            return None
        node = self.by_id.get(style_id)
        if node is None:
            return style_id
        name = first(node, "./w:name")
        return attr(name, "w:val") or style_id

    def style_id(self, name: str | None) -> str | None:
        if not name:
            return None
        for style_id, node in self.by_id.items():
            if self.style_name(style_id) == name:
                return style_id
        return None

    def style_type(self, style_id: str | None) -> str | None:
        if not style_id:
            return None
        node = self.by_id.get(style_id)
        return attr(node, "w:type") if node is not None else None

    def based_on(self, style_id: str | None) -> str | None:
        if not style_id:
            return None
        node = self.by_id.get(style_id)
        if node is None:
            return None
        return attr(first(node, "./w:basedOn"), "w:val")

    def next_style(self, style_id: str | None) -> str | None:
        if not style_id:
            return None
        node = self.by_id.get(style_id)
        if node is None:
            return None
        return attr(first(node, "./w:next"), "w:val")

    def chain(self, style_id: str | None) -> Iterator[Any]:
        seen: set[str] = set()
        current = style_id
        while current and current not in seen:
            seen.add(current)
            node = self.by_id.get(current)
            if node is None:
                break
            yield node
            current = self.based_on(current)

    def paragraph_style_id(self, paragraph: Any) -> str:
        ppr = first(paragraph, "./w:pPr")
        style_id = attr(first(ppr, "./w:pStyle"), "w:val")
        return style_id or self.default_paragraph_id

    def paragraph_style_name(self, paragraph: Any) -> str:
        return self.style_name(self.paragraph_style_id(paragraph)) or "Normal"

    def _ppr_sources(self, paragraph: Any, numbering_ppr: Any | None = None) -> list[Any]:
        sources: list[Any] = []
        ppr = first(paragraph, "./w:pPr")
        if ppr is not None:
            sources.append(ppr)
        for node in self.chain(self.paragraph_style_id(paragraph)):
            item = first(node, "./w:pPr")
            if item is not None:
                sources.append(item)
        if numbering_ppr is not None:
            sources.append(numbering_ppr)
        if self.default_ppr is not None:
            sources.append(self.default_ppr)
        return sources

    def paragraph_properties(self, paragraph: Any, numbering_ppr: Any | None = None) -> dict[str, Any]:
        return self._paragraph_properties_from(self._ppr_sources(paragraph, numbering_ppr))

    def style_paragraph_properties(self, style_id: str | None) -> dict[str, Any]:
        sources: list[Any] = []
        for node in self.chain(style_id):
            item = first(node, "./w:pPr")
            if item is not None:
                sources.append(item)
        if self.default_ppr is not None:
            sources.append(self.default_ppr)
        return self._paragraph_properties_from(sources)

    def _paragraph_properties_from(self, sources: list[Any]) -> dict[str, Any]:
        def pick(path: str) -> Any | None:
            for source in sources:
                found = first(source, path)
                if found is not None:
                    return found
            return None

        result: dict[str, Any] = {}
        jc = pick("./w:jc")
        if jc is not None:
            value = attr(jc, "w:val")
            mapping = {"both": "justify", "distribute": "justify", "start": "left", "end": "right"}
            result["align"] = mapping.get(value or "", value)
        ind = pick("./w:ind")
        if ind is not None:
            left = attr(ind, "w:left") or attr(ind, "w:start")
            right = attr(ind, "w:right") or attr(ind, "w:end")
            first_line = attr(ind, "w:firstLine")
            hanging = attr(ind, "w:hanging")
            if left is not None:
                result["left_indent"] = twips(left, 0.0)
            if right is not None:
                result["right_indent"] = twips(right, 0.0)
            if hanging is not None:
                result["first_line_indent"] = -(twips(hanging, 0.0) or 0.0)
            elif first_line is not None:
                result["first_line_indent"] = twips(first_line, 0.0)
        spacing = pick("./w:spacing")
        if spacing is not None:
            before = attr(spacing, "w:before")
            after = attr(spacing, "w:after")
            line = attr(spacing, "w:line")
            rule = attr(spacing, "w:lineRule")
            if before is not None:
                result["space_before"] = twips(before, 0.0)
            if after is not None:
                result["space_after"] = twips(after, 0.0)
            if line is not None:
                if rule in {"exact", "atLeast"}:
                    result["line_spacing_pt"] = twips(line, 0.0)
                else:
                    result["line_spacing"] = float(line) / 240.0
        for key, name in (
            ("keep_with_next", "keepNext"),
            ("keep_lines", "keepLines"),
            ("page_break_before", "pageBreakBefore"),
        ):
            for source in sources:
                flag = on_off(source, "w:" + name)
                if flag is not None:
                    result[key] = flag
                    break
        outline = pick("./w:outlineLvl")
        if outline is not None:
            value = attr(outline, "w:val")
            if value is not None:
                result["outline_level"] = int(value)
        return result

    def numbering_properties(self, paragraph: Any) -> tuple[int | None, int | None]:
        for source in self._ppr_sources(paragraph):
            num_pr = first(source, "./w:numPr")
            if num_pr is None:
                continue
            ilvl = attr(first(num_pr, "./w:ilvl"), "w:val")
            num_id = attr(first(num_pr, "./w:numId"), "w:val")
            if num_id is None:
                continue
            try:
                return int(ilvl or 0), int(num_id)
            except ValueError:
                return int(ilvl or 0), None
        return None, None

    def run_properties(
        self,
        run: Any,
        paragraph_style_id: str,
        numbering_rpr: Any | None = None,
    ) -> dict[str, Any]:
        rpr = first(run, "./w:rPr")
        sources: list[Any] = []
        if rpr is not None:
            sources.append(rpr)
        style_id = attr(first(rpr, "./w:rStyle"), "w:val") if rpr is not None else None
        for node in self.chain(style_id):
            item = first(node, "./w:rPr")
            if item is not None:
                sources.append(item)
        for node in self.chain(paragraph_style_id):
            item = first(node, "./w:rPr")
            if item is not None:
                sources.append(item)
        if numbering_rpr is not None:
            sources.append(numbering_rpr)
        if self.default_rpr is not None:
            sources.append(self.default_rpr)
        return self._run_properties_from(sources, style_id)

    def style_run_properties(self, style_id: str | None) -> dict[str, Any]:
        sources: list[Any] = []
        for node in self.chain(style_id):
            item = first(node, "./w:rPr")
            if item is not None:
                sources.append(item)
        if self.default_rpr is not None:
            sources.append(self.default_rpr)
        return self._run_properties_from(sources, None)

    def _run_properties_from(self, sources: list[Any], character_style_id: str | None) -> dict[str, Any]:
        def pick(path: str) -> Any | None:
            for source in sources:
                found = first(source, path)
                if found is not None:
                    return found
            return None

        result: dict[str, Any] = {}
        fonts = pick("./w:rFonts")
        if fonts is not None:
            font = (
                fonts.get(qn("w:ascii"))
                or fonts.get(qn("w:hAnsi"))
                or self.theme.font(fonts.get(qn("w:asciiTheme")))
                or self.theme.font(fonts.get(qn("w:hAnsiTheme")))
            )
            east_asia = fonts.get(qn("w:eastAsia")) or self.theme.font(fonts.get(qn("w:eastAsiaTheme")))
            if font:
                result["font"] = font
            if east_asia and east_asia != font:
                result["east_asia_font"] = east_asia
        for key, name in (
            ("bold", "b"),
            ("italic", "i"),
            ("small_caps", "smallCaps"),
            ("all_caps", "caps"),
            ("strike", "strike"),
        ):
            for source in sources:
                flag = on_off(source, "w:" + name)
                if flag is not None:
                    result[key] = flag
                    break
        underline = pick("./w:u")
        if underline is not None:
            value = attr(underline, "w:val")
            result["underline"] = bool(value and value not in {"none", "0", "false"})
        size = pick("./w:sz")
        if size is not None:
            result["size"] = half_points(attr(size, "w:val"))
        color = pick("./w:color")
        if color is not None:
            value = attr(color, "w:val")
            theme_name = attr(color, "w:themeColor")
            tint = attr(color, "w:themeTint")
            shade = attr(color, "w:themeShade")
            if theme_name:
                resolved = self.theme.color(theme_name, tint, shade)
                if resolved:
                    result["color"] = resolved
            elif value and value.lower() != "auto":
                result["color"] = "#" + value.upper()
        highlight = pick("./w:highlight")
        if highlight is not None:
            value = attr(highlight, "w:val")
            if value and value.lower() != "none":
                result["highlight"] = highlight_color(value)
        shading = pick("./w:shd")
        if shading is not None and "highlight" not in result:
            value = attr(shading, "w:fill")
            theme_name = attr(shading, "w:themeFill")
            if theme_name:
                resolved = self.theme.color(theme_name, attr(shading, "w:themeFillTint"), attr(shading, "w:themeFillShade"))
                if resolved:
                    result["highlight"] = resolved
            elif value and value.lower() not in {"auto", "none", "ffffff"}:
                result["highlight"] = "#" + value.upper()
        vert = pick("./w:vertAlign")
        if vert is not None:
            value = attr(vert, "w:val")
            if value == "superscript":
                result["superscript"] = True
            elif value == "subscript":
                result["subscript"] = True
        if style_id:
            result["style_name"] = self.style_name(style_id)
        return result


HIGHLIGHT_COLORS = {
    "black": "#000000",
    "blue": "#0000FF",
    "cyan": "#00FFFF",
    "green": "#00FF00",
    "magenta": "#FF00FF",
    "red": "#FF0000",
    "yellow": "#FFFF00",
    "white": "#FFFFFF",
    "darkBlue": "#000080",
    "darkCyan": "#008080",
    "darkGreen": "#008000",
    "darkMagenta": "#800080",
    "darkRed": "#800000",
    "darkYellow": "#808000",
    "darkGray": "#808080",
    "lightGray": "#C0C0C0",
}


def highlight_color(value: str) -> str:
    return HIGHLIGHT_COLORS.get(value, "#FFFF00")


# ---------------------------------------------------------------------------
# Numbering
# ---------------------------------------------------------------------------


def roman(number: int) -> str:
    if number <= 0:
        return str(number)
    values = (
        (1000, "M"),
        (900, "CM"),
        (500, "D"),
        (400, "CD"),
        (100, "C"),
        (90, "XC"),
        (50, "L"),
        (40, "XL"),
        (10, "X"),
        (9, "IX"),
        (5, "V"),
        (4, "IV"),
        (1, "I"),
    )
    out: list[str] = []
    remaining = number
    for value, glyph in values:
        while remaining >= value:
            out.append(glyph)
            remaining -= value
    return "".join(out)


def letters(number: int) -> str:
    if number <= 0:
        return str(number)
    out: list[str] = []
    remaining = number
    while remaining:
        remaining, remainder = divmod(remaining - 1, 26)
        out.append(chr(ord("A") + remainder))
    return "".join(reversed(out))


def format_number(value: int, fmt: str) -> str:
    if fmt in {"decimal", "decimalHalfWidth", "decimalFullWidth", "numberInDash"}:
        return str(value)
    if fmt == "decimalZero":
        return f"{value:02d}"
    if fmt == "lowerLetter":
        return letters(value).lower()
    if fmt == "upperLetter":
        return letters(value)
    if fmt == "lowerRoman":
        return roman(value).lower()
    if fmt == "upperRoman":
        return roman(value)
    if fmt in {"ordinal", "cardinalText", "ordinalText"}:
        return str(value)
    return str(value)


@dataclass
class NumberingLevel:
    ilvl: int
    fmt: str = "decimal"
    text: str = "%1."
    start: int = 1
    suffix: str = "tab"
    ppr: Any | None = None
    rpr: Any | None = None


class NumberingResolver:
    def __init__(self, root: Any):
        self.root = root
        self.nums: dict[int, Any] = {}
        self.abstracts: dict[int, Any] = {}
        self.counters: dict[int, dict[int, int]] = {}
        if root is None:
            return
        for node in xpath(root, "./w:num"):
            num_id = attr(node, "w:numId")
            if num_id is None:
                continue
            try:
                self.nums[int(num_id)] = node
            except ValueError:
                continue
        for node in xpath(root, "./w:abstractNum"):
            abstract_id = attr(node, "w:abstractNumId")
            if abstract_id is None:
                continue
            try:
                self.abstracts[int(abstract_id)] = node
            except ValueError:
                continue

    def level(self, num_id: int | None, ilvl: int) -> NumberingLevel | None:
        if num_id is None:
            return None
        num = self.nums.get(num_id)
        if num is None:
            return None
        abstract_id = attr(first(num, "./w:abstractNumId"), "w:val")
        try:
            abstract = self.abstracts.get(int(abstract_id or -1))
        except ValueError:
            abstract = None
        if abstract is None:
            return None
        start_override: int | None = None
        level_override: Any | None = None
        for override in xpath(num, "./w:lvlOverride"):
            try:
                if int(attr(override, "w:ilvl", "-1") or -1) == ilvl:
                    level_override = first(override, "./w:lvl")
                    start_value = attr(first(override, "./w:startOverride"), "w:val")
                    if start_value is not None:
                        start_override = int(start_value)
                    break
            except ValueError:
                continue
        node = level_override
        if node is None:
            for candidate in xpath(abstract, "./w:lvl"):
                if attr(candidate, "w:ilvl") == str(ilvl):
                    node = candidate
                    break
        if node is None:
            return None
        fmt = attr(first(node, "./w:numFmt"), "w:val", "decimal") or "decimal"
        text = attr(first(node, "./w:lvlText"), "w:val", "%1.") or ""
        start_value = attr(first(node, "./w:start"), "w:val")
        start = start_override if start_override is not None else int(start_value or 1)
        suffix = attr(first(node, "./w:suff"), "w:val", "tab") or "tab"
        return NumberingLevel(
            ilvl=ilvl,
            fmt=fmt,
            text=text,
            start=start,
            suffix=suffix,
            ppr=first(node, "./w:pPr"),
            rpr=first(node, "./w:rPr"),
        )

    def next(self, num_id: int | None, ilvl: int) -> tuple[str | None, NumberingLevel | None]:
        level = self.level(num_id, ilvl)
        if level is None or num_id is None:
            return None, None
        counters = self.counters.setdefault(num_id, {})
        current = counters.get(ilvl, level.start - 1) + 1
        counters[ilvl] = current
        for deeper in list(counters):
            if deeper > ilvl:
                del counters[deeper]
        if level.fmt in {"bullet", "none"}:
            return level.text, level
        marker = level.text
        for index in range(9):
            placeholder = f"%{index + 1}"
            if placeholder not in marker:
                continue
            if index <= ilvl:
                deeper_level = self.level(num_id, index)
                start = deeper_level.start if deeper_level is not None else 1
                value = counters.get(index, start - 1)
                if index == ilvl:
                    value = current
                marker = marker.replace(placeholder, format_number(value, deeper_level.fmt if deeper_level else level.fmt))
            else:
                marker = marker.replace(placeholder, "")
        return marker, level


# ---------------------------------------------------------------------------
# Page setup
# ---------------------------------------------------------------------------


def page_setup(section: Any) -> PageSetup:
    def points(value: Any, default: float) -> float:
        try:
            return float(value.pt) if value is not None else default
        except (AttributeError, TypeError, ValueError):
            return default

    width = points(getattr(section, "page_width", None), 595.28)
    height = points(getattr(section, "page_height", None), 841.89)
    setup = PageSetup(
        width_pt=width,
        height_pt=height,
        margin_top=points(getattr(section, "top_margin", None), 72.0),
        margin_right=points(getattr(section, "right_margin", None), 72.0),
        margin_bottom=points(getattr(section, "bottom_margin", None), 72.0),
        margin_left=points(getattr(section, "left_margin", None), 72.0),
        header_distance=points(getattr(section, "header_distance", None), 36.0),
        footer_distance=points(getattr(section, "footer_distance", None), 36.0),
        orientation="landscape" if width > height else "portrait",
    )
    try:
        sect_pr = section._sectPr
        cols = first(sect_pr, "./w:cols")
        if cols is not None:
            setup.columns = int(attr(cols, "w:num", "1") or 1)
        setup.title_page = bool(on_off(sect_pr, "w:titlePg") or False)
    except Exception:
        pass
    return setup


def section_break_before(paragraph: Any) -> bool:
    ppr = first(paragraph, "./w:pPr")
    return first(ppr, "./w:sectPr") is not None


# ---------------------------------------------------------------------------
# Misc
# ---------------------------------------------------------------------------


def relationship_target(part: Any, rel_id: str | None) -> tuple[str | None, bool]:
    if not part or not rel_id:
        return None, False
    try:
        rel = part.rels[rel_id]
    except KeyError:
        return None, False
    return rel.target_ref, bool(rel.is_external)


def image_extension(content_type: str, partname: str) -> str:
    mapping = {
        "image/png": ".png",
        "image/jpeg": ".jpg",
        "image/jpg": ".jpg",
        "image/gif": ".gif",
        "image/bmp": ".bmp",
        "image/tiff": ".tiff",
        "image/x-emf": ".emf",
        "image/emf": ".emf",
        "image/x-wmf": ".wmf",
        "image/wmf": ".wmf",
        "image/svg+xml": ".svg",
        "image/webp": ".webp",
    }
    if content_type in mapping:
        return mapping[content_type]
    suffix = ""
    if partname:
        match = re.search(r"(\.[A-Za-z0-9]+)$", str(partname))
        if match:
            suffix = match.group(1).lower()
    return suffix or ".bin"


def iter_block_children(element: Any) -> Iterator[Any]:
    """Yield paragraphs and tables, unwrapping common containers."""
    if element is None:
        return
    for node in element:
        name = local_name(node)
        if name in {"p", "tbl"}:
            yield node
        elif name in {"sdt", "customXml", "smartTag"}:
            content = first(node, "./w:sdtContent")
            if content is None:
                content = first(node, "./w:customXml")
            if content is None:
                content = node
            yield from iter_block_children(content)
        elif name == "altChunk":
            yield node
