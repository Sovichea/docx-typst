"""Share original source assets with existing OPC media by content hash."""
import base64
import hashlib

from docx.opc.packuri import PackURI
from docx.opc.part import Part

ASSET_REL = "https://docx-typst.local/source/asset"


def _assets(package):
    return {hashlib.sha256(part.blob).hexdigest(): part
            for part in package.iter_parts()
            if part.content_type.startswith("image/") or part.content_type == "application/octet-stream"
            or str(part.partname).startswith("/docxTypst/assets/")}


def pack_files(source_part, files):
    """Share visible media by hash, without links that make Word clone it."""
    blobs, issues = unpack_files(source_part, files)
    if issues:
        raise ValueError("cannot compact changed or missing original source assets: " + ", ".join(issues))
    parts = list(source_part.package.iter_parts())
    shared, private = {}, {}
    for owner in parts:
        if str(owner.partname).startswith("/docxTypst/"):
            if str(owner.partname).startswith("/docxTypst/assets/"):
                private[hashlib.sha256(owner.blob).hexdigest()] = owner
            continue
        for relation in owner.rels.values():
            if not relation.is_external and relation.target_part.content_type.startswith("image/"):
                target = relation.target_part
                shared[hashlib.sha256(target.blob).hexdigest()] = target
    # Word clones media when an opaque source part also links to it. A hash
    # reference needs no extra relationship; the visible document owns it.
    for rid, relation in list(source_part.rels.items()):
        if relation.reltype == ASSET_REL:
            del source_part.rels[rid]
    result = {}
    for relative, blob in blobs.items():
        if relative.lower().endswith(".typ"):
            result[relative] = base64.b64encode(blob).decode("ascii")
            continue
        digest = hashlib.sha256(blob).hexdigest()
        reference = {"sha256": digest}
        if digest not in shared:
            part = private.get(digest)
            if part is None:
                part = Part(PackURI(f"/docxTypst/assets/{digest}.bin"), "application/octet-stream",
                            blob, source_part.package)
                private[digest] = part
            reference["relationship"] = source_part.relate_to(part, ASSET_REL)
        result[relative] = reference
    return result


def unpack_files(source_part, files):
    """Resolve by hash across Word renames; detect changed or missing assets."""
    assets = _assets(source_part.package)
    result, issues = {}, []
    for relative, encoded in files.items():
        if isinstance(encoded, str):  # Original version-1 packages remain readable.
            result[relative] = base64.b64decode(encoded, validate=True)
            continue
        if not isinstance(encoded, dict) or not isinstance(encoded.get("sha256"), str):
            raise ValueError("invalid native source asset reference")
        part = assets.get(encoded["sha256"])
        if part is None:
            relationship = source_part.rels.get(encoded.get("relationship"))
            part = relationship.target_part if relationship is not None and not relationship.is_external else None
            if part is None or hashlib.sha256(part.blob).hexdigest() != encoded["sha256"]:
                issues.append(relative)
        if part is not None:
            result[relative] = part.blob
    return result, issues
