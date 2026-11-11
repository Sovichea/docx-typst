"""Materialize explicit Typst computations into editable export-only markup."""
import json
from pathlib import Path
import re

import typst

from .typst_syntax import parse_call_at, strip_comments, typst_escape


def _length(value):
    value = re.sub(r"^0%\s*\+\s*", "", str(value))
    if "%" in value:
        raise ValueError(f"computed percentage track needs an explicit export width: {value}")
    return value


def _plain(node):
    if node.get("kind") == "text":
        return node.get("text", "")
    if node.get("kind") == "space":
        return " "
    if "body" in node:
        return _plain(node["body"])
    return "".join(_plain(c) for c in node.get("children", []))


def _markup(node):
    kind = node["kind"]
    if kind == "text":
        return typst_escape(node.get("text", ""))
    if kind == "space":
        return " "
    if kind in {"linebreak", "parbreak"}:
        return "#linebreak()" if kind == "linebreak" else "\n\n"
    if kind == "v":
        return f'#v({_length(node["amount"])})'
    if kind in {"metadata", "label"}:
        return ""
    if kind == "sequence":
        return "".join(_markup(c) for c in node.get("children", []))
    if kind in {"strong", "emph"}:
        delimiter = "*" if kind == "strong" else "_"
        return delimiter + _markup(node["body"]) + delimiter
    if kind in {"block", "box"}:
        return _markup(node["body"])
    if kind == "align":
        alignment = "right" if "right" in node["align"] else "center" if "center" in node["align"] else "left"
        return f'#align({alignment})[{_markup(node["body"])}]'
    if kind == "link":
        return f'#link({json.dumps(node["dest"])})[{_markup(node["body"])}]'
    if kind == "grid":
        cells = node["cells"]
        if len(node["columns"]) == 2 and len(cells) == 2 and _plain(cells[0]["body"]) == "$":
            return f'#docx-money({json.dumps(_plain(cells[1]["body"]))})'
        args = ['columns: (' + ', '.join(_length(v) for v in node["columns"]) + ',)',
                f'column-gutter: {_length(node["column-gutter"])}',
                f'row-gutter: {_length(node["row-gutter"])}']
        return '#grid(' + ', '.join(args) + ', ' + ', '.join('[' + _markup(c["body"]) + ']' for c in cells) + ')'
    if kind == "table":
        return _table(node)
    raise ValueError(f'unsupported computed content: {kind}; use explicit supported formatting')


def _table(node):
    widths = [_length(v) for v in node["columns"]]
    # A zero-width trailing bookkeeping column has no visible Word cell.
    while widths and widths[-1] in {"0pt", "0fr"}:
        widths.pop()
    ncols = len(widths)
    cells = [c for c in node["cells"] if c["col"] < ncols]
    rows = max(c["row"] + c["rowspan"] for c in cells)
    lines = [f'#docx-table(columns: {ncols}, rows: {rows}, widths: (' + ', '.join(widths) + ',),',
             f'header-rows: {node["header-rows"]}, stroke: none, inset: 0pt,']
    merges, fills, insets, borders, aligns, valigns, bodies = [], [], [], [], [], [], []
    for cell in cells:
        r, c = cell["row"], cell["col"]
        span = min(cell["colspan"], ncols - c)
        if cell["rowspan"] > 1 or span > 1:
            merges.append(f'({r}, {c}, {cell["rowspan"]}, {span})')
        if cell["fill"] != "none":
            fills.append(f'({r}, {c}, {cell["fill"]})')
        insets.append(f'({r}, {c}, (' + ', '.join(f'{k}: {_length(v)}' for k, v in cell["insets"].items()) + '))')
        paint = {k: cell["borders"].get(k, cell["borders"].get("rest", "none")) for k in ("top", "bottom", "left", "right")}
        borders.append(f'({r}, {c}, (' + ', '.join(f'{k}: {v}' for k, v in paint.items()) + '))')
        alignment = "right" if "right" in cell["align"] else "center" if "center" in cell["align"] else "left"
        aligns.append(f'({r}, {c}, {json.dumps(alignment)})')
        valigns.append(f'({r}, {c}, "center")' if "horizon" in cell["align"] else f'({r}, {c}, "top")')
        bodies.append('[' + _markup(cell["body"]) + ']')
    for key, values in (("merges", merges), ("fills", fills), ("cell-insets", insets),
                        ("cell-borders", borders), ("cell-aligns", aligns), ("valigns", valigns)):
        if values:
            lines.append(f'{key}: (' + ', '.join(values) + ',),')
    lines.append('cells: (' + ', '.join(bodies) + ',))')
    return "\n".join(lines)


def materialize_computed(source: Path, text: str):
    if not re.search(r"#docx-eval\s*\(", strip_comments(text)):
        return text, {}
    values = json.loads(typst.query(str(source), "<docx-typst-eval>", field="value"))
    results = {}
    for item in values:
        identity = item["id"]
        if identity in results:
            raise ValueError(f"duplicate docx-eval id: {identity}")
        results[identity] = item["value"]
    working = strip_comments(text)
    edits = []
    end = -1
    for match in re.finditer(r"#docx-eval\s*\(", working):
        if match.start() < end:
            continue
        call = parse_call_at(working, match.start() + 1)
        if call is None or not call.positional or not isinstance(call.positional[0], str):
            raise ValueError("docx-eval requires a literal, unique string id")
        identity = call.positional[0]
        if identity not in results:
            raise ValueError(f"docx-eval was not laid out by Typst: {identity}")
        end = match.start() + 1 + len(call.raw)
        edits.append((match.start(), end, _markup(results[identity])))
    for start, end, replacement in reversed(edits):
        working = working[:start] + replacement + working[end:]
    return working, results
