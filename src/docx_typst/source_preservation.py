"""Package native Typst artwork and dependencies with exported Word drawings."""
from __future__ import annotations

import base64
import json
from pathlib import Path
import re
from typing import Any

from docx.opc.packuri import PackURI
from docx.opc.part import Part
from docx.oxml.ns import qn

from .model import NativeBlock
from .positioned import SOURCE_NS, _box_text
from .typst_syntax import parse_call_at, strip_comments, typst_escape

REL = SOURCE_NS + "/package"


def calls(text: str, name: str):
    for match in re.finditer(r"#" + re.escape(name) + r"\s*\(", text):
        call = parse_call_at(text, match.start() + 1)
        if call is not None:
            yield match.start(), call


def record_canvas(writer: Any, canvas: Any, positions: list[dict[str, Any]], bookmark: str) -> None:
    source = canvas.source
    for _, call in calls(writer.source_text, "docx-canvas"):
        if call.arg("id", "canvas") == canvas.identifier:
            source = "#" + call.raw
            break
    for record in positions:
        for _, call in calls(source, "docx-position"):
            if strip_comments("#" + call.raw) == record["source"]:
                record["source"] = "#" + call.raw
                break
    writer.source_records["canvas:" + canvas.identifier] = {
        "source": source, "positions": positions, "bookmark": bookmark,
    }


def record_static(writer: Any, inline: Any, image: Any) -> None:
    call = parse_call_at(image.typst_source, 1)
    if call is None or call.body is None:
        return
    source = image.typst_source
    for _, original in calls(writer.source_text, "docx-static"):
        if original.body and strip_comments(original.body.text) == call.body.text:
            source = "#" + original.raw
            break
    identity = "static:" + str(len(writer.source_records))
    inline.find(qn("wp:docPr")).set("descr", identity)
    writer.source_records[identity] = {"source": source}


def _dependencies(base: Path, text: str) -> dict[str, str]:
    files: dict[str, str] = {}
    base = base.resolve()

    def collect(parent: Path, content: str) -> None:
        for match in re.finditer(r'"([^"\r\n]+)"', content):
            candidate = (parent / match.group(1)).resolve()
            if not candidate.is_relative_to(base) or not candidate.is_file():
                continue
            relative = candidate.relative_to(base).as_posix()
            if relative in files:
                continue
            data = candidate.read_bytes()
            files[relative] = base64.b64encode(data).decode("ascii")
            if candidate.suffix == ".typ":
                collect(candidate.parent, data.decode("utf-8-sig"))

    collect(base, text)
    # A helper's image(path) resolves relative to its defining module, not
    # the call site. Retain existing assets at those module-relative paths too.
    modules = [(base, text)] + [((base / name).parent, base64.b64decode(data).decode("utf-8-sig"))
                               for name, data in list(files.items()) if name.endswith(".typ")]
    references = {match.group(1) for _, content in modules
                  for match in re.finditer(r'"([^"\r\n]+)"', content)}
    for folder, _ in modules:
        for reference in references:
            collect(folder, json.dumps(reference))
    return files


def _imports(text: str) -> str:
    lines = text.splitlines()
    result = []
    index = 0
    while index < len(lines):
        if lines[index].startswith("#import "):
            line = lines[index]
            balance = line.count("(") - line.count(")")
            while balance > 0 and index + 1 < len(lines):
                index += 1
                line += "\n" + lines[index]
                balance += lines[index].count("(") - lines[index].count(")")
            # The forward converter supplies the current reusable primitives.
            if not re.search(r'#import "(?:docx-typst/)?template.typ"', line):
                result.append(line)
        index += 1
    return "\n".join(result)


def attach_sources(writer: Any) -> None:
    if not writer.source_text and not writer.source_records:
        return
    data = {"version": 2, "records": writer.source_records,
            "imports": _imports(writer.source_text),
            "files": _dependencies(writer.base, writer.source_text)}
    from .roundtrip_source import snapshot_entry, document_state
    entry = snapshot_entry(writer)
    if entry is not None:
        data["entry"] = entry
    package = writer.doc.part.package
    part = Part(PackURI("/docxTypst/source.json"), "application/json",
                json.dumps(data, ensure_ascii=False).encode("utf-8"), package)
    writer.doc.part.relate_to(part, REL)
    from .source_assets import pack_files
    data["files"] = pack_files(part, data["files"])
    part._blob = json.dumps(data, ensure_ascii=False).encode("utf-8")
    if entry is not None:
        entry["baseline"] = document_state(writer.doc)
        part._blob = json.dumps(data, ensure_ascii=False).encode("utf-8")


def load_sources(reader: Any) -> dict[str, Any]:
    for rel in reader.docx.part.rels.values():
        if rel.reltype == REL and not rel.is_external:
            data = json.loads(rel.target_part.blob)
            if data.get("version") not in {1, 2}:
                return {}
            root = reader.output.resolve()
            from .source_assets import unpack_files
            files, issues = unpack_files(rel.target_part, data.get("files", {}))
            if issues:
                data["_asset_issues"] = issues
                reader.warn("native_source_asset_changed", "approximated", "document",
                            source_type="native source dependencies", fallback="reconstruct current Word content",
                            reason="original source asset bytes are missing or changed: " + ", ".join(issues))
            for relative, blob in files.items():
                target = (root / relative).resolve()
                if not target.is_relative_to(root):
                    raise ValueError("native source dependency escapes output directory")
                target.parent.mkdir(parents=True, exist_ok=True)
                target.write_bytes(blob)
            return data
    return {}


def read_canvas(reader: Any, paragraph: Any) -> NativeBlock | None:
    identity = paragraph.get(f"{{{SOURCE_NS}}}canvas")
    records = getattr(reader, "native_sources", {}).get("records", {})
    if identity is None:
        # Word strips unknown attributes when saving. Its drawing names remain
        # stable, so recognize the same canvas through its editable text boxes.
        names = {node.get("name") for node in paragraph.findall(".//" + qn("wp:docPr"))}
        bookmarks = {node.get(qn("w:name")) for node in paragraph.findall(qn("w:bookmarkStart"))}
        for key, candidate in records.items():
            if key.startswith("canvas:") and (candidate.get("bookmark") in bookmarks
                    or any(item["id"] in names for item in candidate["positions"])):
                identity = key.removeprefix("canvas:")
                break
    record = records.get("canvas:" + identity) if identity is not None else None
    if record is None:
        return None
    source = record["source"]
    for node in paragraph.findall(".//" + qn("w:pStyle")):
        reader.used_style_ids.add(node.get(qn("w:val")))
    for item in record["positions"]:
        shapes = paragraph.findall(".//" + qn("wp:anchor"))
        shape = next((node for node in shapes
                      if node.find(qn("wp:docPr")).get("name") == item["id"]), None)
        if shape is None:
            # A deleted text box is an intentional Word edit.
            source = source.replace(item["source"], "", 1)
            continue
        content = shape.find(".//" + qn("w:txbxContent"))
        current = _box_text(content)
        if current == item["text"]:
            continue
        call = parse_call_at(item["source"], 1)
        if call is None or call.body is None:
            continue
        # Preserve the position arguments; text edits are rendered through the
        # same exported paragraph style instead of discarding the drawing code.
        style = json.dumps(item.get("style") or "Normal")
        body = f'#docx-para(styles: docx-styles, style: {style}, before: 0pt, after: 0pt)['
        body += "#linebreak()".join(typst_escape(line) for line in current.split("\n")) + "]"
        start = item["source"].find("[", item["source"].find(")"))
        replacement = item["source"][:start + 1] + body + "]"
        source = source.replace(item["source"], replacement, 1)
    return NativeBlock(source=source)
