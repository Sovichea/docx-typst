"""Drawing and media reading for the forward converter."""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any

from .media_compat import (
    METAFILE,
    RASTER_CONVERT,
    RASTER_DIRECT,
    image_size,
    prepare_media,
    sha256_bytes,
)
from .model import Asset, ImageRef, Run
from .ooxml import R, emu_to_pt, first, image_extension, xpath


class MediaMixin:
    def _read_drawing(self, drawing: Any, part: Any) -> list[Run]:
        results: list[Run] = []
        pic = first(drawing, ".//pic:pic")
        if pic is not None:
            image = self._image_from_pic(
                pic,
                drawing,
                part,
                floating=first(drawing, ".//wp:anchor") is not None,
            )
            if image is not None:
                return [Run(kind="image", image=image)]
        if first(drawing, ".//c:chart") is not None:
            self.warn(
                "chart",
                "unsupported",
                f"paragraph:{self.counts['paragraphs']}",
                source_type="chart",
                fallback="omitted",
                reason="native chart data is not converted",
                expected_loss="The chart is missing from the Typst project",
            )
            return results
        if first(drawing, ".//dgm:relIds") is not None:
            self.warn(
                "smartart",
                "unsupported",
                f"paragraph:{self.counts['paragraphs']}",
                source_type="SmartArt",
                fallback="omitted",
                reason="SmartArt data is not converted",
                expected_loss="The SmartArt graphic is missing",
            )
            return results
        if first(drawing, ".//wps:wsp") is not None or first(drawing, ".//wpg:wgp") is not None:
            text = "".join(xpath(drawing, ".//w:txbxContent//w:t/text()"))
            self.warn(
                "text_box",
                "approximated",
                f"paragraph:{self.counts['paragraphs']}",
                source_type="text box or shape",
                fallback="text extracted; shape and position omitted",
                reason="floating Word shapes are not in the supported subset",
                expected_loss="Shape geometry, fill and frame are lost",
            )
            if text:
                results.append(Run(text=text))
            return results
        if first(drawing, ".//m:oMath") is not None:
            text = "".join(drawing.itertext())
            if text:
                self.warn(
                    "math",
                    "approximated",
                    f"paragraph:{self.counts['paragraphs']}",
                    source_type="equation",
                    fallback="linear text",
                )
                results.append(Run(text=text))
            return results
        return results

    def _image_from_pic(
        self, pic: Any, drawing: Any, part: Any, *, floating: bool
    ) -> ImageRef | None:
        blip = first(pic, ".//a:blip")
        rel_id = blip.get(f"{{{R}}}embed") if blip is not None else None
        svg_blip = first(pic, ".//asvg:svgBlip")
        svg_rel = svg_blip.get(f"{{{R}}}embed") if svg_blip is not None else None
        chosen_rel = svg_rel or rel_id
        if not chosen_rel:
            self.warn(
                "image_relationship",
                "unsupported",
                f"paragraph:{self.counts['paragraphs']}",
                source_type="image",
                fallback="omitted",
                reason="picture has no embedded relationship",
            )
            return None
        try:
            image_part = part.related_parts[chosen_rel]
        except (KeyError, AttributeError):
            self.warn(
                "missing_image_part",
                "unsupported",
                f"paragraph:{self.counts['paragraphs']}",
                source_type="image",
                fallback="omitted",
                reason="image relationship target is missing",
            )
            return None
        if not self.extract_media:
            self.counts["images"] += 1
            return ImageRef(path="", alt=None)
        data = image_part.blob
        content_type = getattr(image_part, "content_type", "") or ""
        partname = str(getattr(image_part, "partname", "image"))
        suffix = image_extension(content_type, partname)
        asset, emitted = self._save_asset(data, suffix, content_type, partname)
        extent = first(drawing, ".//wp:extent")
        width = emu_to_pt(extent.get("cx") if extent is not None else None)
        height = emu_to_pt(extent.get("cy") if extent is not None else None)
        if not width or not height:
            measured_w, measured_h = image_size(data)
            width = width or measured_w
            height = height or measured_h
        src_rect = first(pic, ".//a:srcRect")
        crop = None
        if src_rect is not None:
            values = []
            for key in ("l", "t", "r", "b"):
                raw = src_rect.get(key)
                values.append(float(raw or 0) / 100000.0)
            crop = tuple(values)  # type: ignore[assignment]
        rotation = 0.0
        xfrm = first(pic, ".//a:xfrm")
        if xfrm is not None and xfrm.get("rot"):
            try:
                rotation = float(xfrm.get("rot")) / 60000.0
            except ValueError:
                pass
        doc_pr = first(drawing, ".//wp:docPr")
        alt = doc_pr.get("descr") if doc_pr is not None else None
        title = doc_pr.get("title") if doc_pr is not None else None
        if floating:
            self.warn(
                "floating_image",
                "approximated",
                f"paragraph:{self.counts['paragraphs']}",
                source_type="floating image",
                fallback="placed as a centered block image",
                reason="Typst does not reproduce Word's text-wrap geometry",
                expected_loss="Exact position and text wrapping may differ",
            )
        self.counts["images"] += 1
        return ImageRef(
            path=str(emitted.relative_to(self.output).as_posix()) if emitted is not None else "",
            width_pt=width,
            height_pt=height,
            alt=alt,
            title=title,
            crop=crop,
            rotation=rotation,
            floating=floating,
            align="center" if floating else None,
            source_part=partname,
        )

    def _read_vml(self, pict: Any, part: Any) -> list[Run]:
        image = first(pict, ".//v:imagedata")
        if image is None:
            return []
        rel_id = image.get(f"{{{R}}}id")
        try:
            image_part = part.related_parts[rel_id]
        except (KeyError, TypeError):
            return []
        suffix = image_extension(
            getattr(image_part, "content_type", ""), str(image_part.partname)
        )
        asset, emitted = self._save_asset(
            image_part.blob,
            suffix,
            getattr(image_part, "content_type", ""),
            str(image_part.partname),
        )
        self.warn(
            "vml_image",
            "approximated",
            f"paragraph:{self.counts['paragraphs']}",
            source_type="VML image",
            fallback="image extracted; VML geometry omitted",
            reason="VML shapes are legacy and not positionally converted",
        )
        self.counts["images"] += 1
        width = None
        height = None
        style = image.get("style") or ""
        match = re.search(r"width:\s*([\d.]+)pt", style)
        if match:
            width = float(match.group(1))
        match = re.search(r"height:\s*([\d.]+)pt", style)
        if match:
            height = float(match.group(1))
        return [
            Run(
                kind="image",
                image=ImageRef(
                    path=str(emitted.relative_to(self.output).as_posix())
                    if emitted is not None
                    else "",
                    width_pt=width,
                    height_pt=height,
                ),
            )
        ]

    def _save_asset(
        self, data: bytes, suffix: str, content_type: str, partname: str
    ) -> tuple[Asset | None, Path | None]:
        digest = sha256_bytes(data)
        if digest in self.assets_by_hash:
            asset = self.assets_by_hash[digest]
            path = self.output / asset.output_path if asset.emitted else None
            return asset, path
        self.image_counter += 1
        base = f"image{self.image_counter}_{digest[:10]}"
        self.media_dir.mkdir(parents=True, exist_ok=True)
        suffix = suffix or ".png"
        original = self.media_dir / f"{base}{suffix}"
        original.write_bytes(data)
        emitted: Path | None = None
        warning: str | None = None
        derived: str | None = None
        if suffix in RASTER_DIRECT or suffix == ".svg":
            emitted = original
        elif suffix in RASTER_CONVERT or suffix in METAFILE:
            prepared, warning = prepare_media(
                data, suffix, content_type, self.media_dir / f"{base}_render"
            )
            if warning is None and prepared is not None:
                emitted = prepared
                derived = str(prepared.relative_to(self.output).as_posix())
                if emitted.suffix.lower() == suffix:
                    derived = None
            elif warning:
                self.warn(
                    "media_conversion",
                    "unsupported" if suffix in METAFILE else "approximated",
                    f"asset:{original.name}",
                    source_type=suffix.lstrip("."),
                    fallback="original asset preserved; not emitted in main.typ",
                    reason=warning,
                    expected_loss="The image may be missing from the rendered Typst document",
                )
        else:
            emitted = original
        if emitted is None:
            self.warn(
                "unsupported_media_format",
                "unsupported",
                f"asset:{original.name}",
                source_type=suffix.lstrip("."),
                fallback="original asset preserved",
                reason="no supported Typst renderer for this image format",
                expected_loss="The image is missing from the Typst project",
            )
        asset = Asset(
            original_name=original.name,
            output_path=str((emitted or original).relative_to(self.output).as_posix()),
            sha256=digest,
            byte_count=len(data),
            content_type=content_type,
            source_part=partname,
            derived_path=derived,
            warning=warning,
            emitted=emitted is not None,
        )
        self.assets_by_hash[digest] = asset
        self.counts["assets"] += 1
        return asset, emitted
