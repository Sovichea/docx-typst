"""Restore exact native source only when Word content and layout still match."""
from dataclasses import asdict, replace
import base64
import hashlib
import json
from pathlib import Path
from tempfile import TemporaryDirectory

from docx.oxml.ns import qn
from lxml import etree

BOOLS = {"bold", "italic", "underline", "strike", "small_caps", "all_caps",
         "keep_with_next", "keep_lines", "page_break_before"}
ASVG = "{http://schemas.microsoft.com/office/drawing/2016/SVG/main}svgBlip"


def _model_state(model, output: Path):
    data = asdict(model)
    styles = data["styles"]

    def effective(name, seen=()):
        own = styles.get(name, {})
        parent = own.get("based_on")
        base = effective(parent, seen + (name,)) if parent and parent not in seen else {}
        result = {**base, **{k: v for k, v in own.items() if v is not None}}
        if own.get("line_spacing") is not None and own.get("line_spacing_pt") is None:
            result.pop("line_spacing_pt", None)
        return result

    resolved = {name: {key: effective(name).get(key) for key in own}
                for name, own in styles.items()}

    def normalize(value, key="", style=None):
        if isinstance(value, list):
            return [normalize(v, key, style) for v in value]
        if isinstance(value, dict):
            current = style or resolved.get("Normal", {})
            if "style" in value:
                current = resolved.get(value.get("style") or "Normal", current)
            elif value.get("name") in resolved:
                current = resolved[value["name"]]
            result = {}
            for k, v in value.items():
                if k in {"source_part", "original_name", "derived_path", "original_path", "path"}:
                    continue
                if k in BOOLS and v is None:
                    v = current.get(k, False)
                result[k] = normalize(v, k, current)
            return result
        if key in BOOLS:
            return bool(value)
        return value

    for name, spec in resolved.items():
        styles[name] = spec
    for key in ("assets", "diagnostics", "source"):
        data.pop(key, None)
    for key in ("created", "modified", "last_modified_by"):
        data["metadata"].pop(key, None)
    return normalize(data)


def _package_state(document):
    drawings, indents, extra = [], [], []
    for part in document.part.package.iter_parts():
        name = str(part.partname)
        if any(name.startswith(prefix) for prefix in
               ("/word/comments", "/word/charts/", "/word/embeddings/")):
            extra.append(hashlib.sha256(part.blob).hexdigest())
        if part is not document.part and part.content_type not in {
            "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"}:
            continue
        root = getattr(part, "element", None)
        if root is None:
            root = etree.fromstring(part.blob)
        table_indents = []
        for table in root.iter(qn("w:tbl")):
            node = table.find("./" + qn("w:tblPr") + "/" + qn("w:tblInd"))
            table_indents.append(int(node.get(qn("w:w"), "0")) if node is not None else 0)
        if table_indents:
            indents.append((name, table_indents))
        for node in list(root.iter(qn("wp:inline"))) + list(root.iter(qn("wp:anchor"))):
            geometry = []
            for tag in ("wp:extent", "wp:positionH", "wp:positionV", "wp:wrapNone", "wp:wrapSquare", "wp:wrapTight", "wp:wrapThrough", "wp:wrapTopAndBottom", "a:srcRect"):
                child = node.find(".//" + qn(tag)) if tag == "a:srcRect" else node.find(qn(tag))
                if child is not None:
                    geometry.append((tag, sorted(child.attrib.items()),
                                     [(c.tag, c.text) for c in child]))
            for fill in node.iter(qn("a:solidFill")):
                geometry.append(("fill", [(c.tag, sorted(c.attrib.items())) for c in fill]))
            for transform in node.iter(qn("a:xfrm")):
                changed = [(k, v) for k, v in transform.attrib.items() if v not in {"0", "false"}]
                if changed:
                    geometry.append(("transform", sorted(changed)))
            images = []
            svg = list(node.iter(ASVG))
            blips = svg or list(node.iter(qn("a:blip")))
            for blip in blips:
                rel_id = blip.get(qn("r:embed"))
                if rel_id and rel_id in part.rels:
                    images.append(hashlib.sha256(part.rels[rel_id].target_part.blob).hexdigest())
            drawings.append((geometry, images))
        for tag in ("w:ins", "w:del", "w:moveFrom", "w:moveTo"):
            extra.extend(etree.tostring(n).hex() for n in root.iter(qn(tag)))
    return {"drawings": sorted(drawings, key=lambda v: json.dumps(v, sort_keys=True)),
            "indents": sorted(indents), "extra": sorted(extra)}


def semantic_state(reader, model):
    counts = reader.counts.copy()
    state = _model_state(model, reader.output)
    # Native canvas restoration retains the original source. Read the actual
    # text-box paragraphs too, so a Word-only font/style edit cannot hide in it.
    boxes = []
    for part in reader.docx.part.package.iter_parts():
        if part is not reader.docx.part and not any(part.content_type.endswith(suffix)
                for suffix in ("header+xml", "footer+xml")):
            continue
        root = getattr(part, "element", None)
        if root is None:
            root = etree.fromstring(part.blob)
        for box in root.iter(qn("w:txbxContent")):
            if any(n.tag == "{http://schemas.openxmlformats.org/markup-compatibility/2006}Fallback" for n in box.iterancestors()):
                continue
            for paragraph in box.findall(qn("w:p")):
                boxes.extend(reader._read_paragraph(paragraph, part))
    for block in boxes:
        if hasattr(block, "space_before") and block.space_before is None:
            spec = model.styles.get(block.style or "Normal")
            block.space_before = spec.space_before if spec and spec.space_before is not None else 0
    if boxes:
        state["textboxes"] = _model_state(replace(model, blocks=boxes, header=[], footer=[]), reader.output)
    reader.counts.clear()
    reader.counts.update(counts)
    return state


def document_state(document):
    from .docx_to_typst import DocxReader
    with TemporaryDirectory(prefix="docx-typst-snapshot-") as temporary:
        root = Path(temporary)
        source = root / "document.docx"
        output = root / "project"
        output.mkdir()
        document.save(source)
        reader = DocxReader(source, output)
        state = semantic_state(reader, reader.read())
    return {"model": state, "package": _package_state(document)}


def snapshot_entry(writer):
    path = writer.base / (writer.model.source or "")
    if not path.is_file():
        return None
    content = path.read_bytes()
    return {"filename": path.name, "source": base64.b64encode(content).decode("ascii"),
            "sha256": hashlib.sha256(content).hexdigest()}


def equivalent(a, b, path=()):
    if isinstance(a, dict) and isinstance(b, dict):
        return a.keys() == b.keys() and all(equivalent(a[k], b[k], path + (k,)) for k in a)
    if isinstance(a, (list, tuple)) and isinstance(b, (list, tuple)):
        return len(a) == len(b) and all(equivalent(x, y, path + (str(i),)) for i, (x, y) in enumerate(zip(a, b)))
    # Word rounds table tracks to twips during a save. Other edits remain exact.
    if ("columns" in path or path[-1:] == ("width_pt",)) and isinstance(a, (int, float)) and isinstance(b, (int, float)):
        return abs(a - b) <= 0.050001
    return a == b


def restored_entry(reader, model):
    record = reader.native_sources.get("entry")
    if not record or "baseline" not in record:
        return None
    state = {"model": semantic_state(reader, model), "package": _package_state(reader.docx)}
    if not equivalent(record["baseline"], state):
        reader.warn("native_source_changed", "approximated", "document",
                    source_type="native source snapshot", fallback="reconstruct current Word content",
                    reason="Word content or layout differs from the exported source; stale source was not restored")
        return None
    content = base64.b64decode(record["source"], validate=True)
    if hashlib.sha256(content).hexdigest() != record["sha256"]:
        raise ValueError("native source snapshot checksum mismatch")
    return content
