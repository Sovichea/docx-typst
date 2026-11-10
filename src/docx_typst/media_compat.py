"""Media normalization shared by both converters."""

from __future__ import annotations

import hashlib
import io
import os
from pathlib import Path
import shutil
import subprocess
import tempfile
from typing import BinaryIO

from PIL import Image


RASTER_DIRECT = {".png", ".jpg", ".jpeg"}
SVG = {".svg"}
METAFILE = {".emf", ".wmf"}
RASTER_CONVERT = {".gif", ".bmp", ".tif", ".tiff", ".webp", ".ico"}


def sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def extension_for(name: str, content_type: str = "") -> str:
    suffix = Path(name).suffix.lower()
    if suffix:
        return suffix
    mapping = {
        "image/png": ".png",
        "image/jpeg": ".jpg",
        "image/gif": ".gif",
        "image/bmp": ".bmp",
        "image/tiff": ".tiff",
        "image/svg+xml": ".svg",
        "image/x-emf": ".emf",
        "image/emf": ".emf",
        "image/x-wmf": ".wmf",
        "image/wmf": ".wmf",
        "image/webp": ".webp",
    }
    return mapping.get(content_type, ".bin")


def image_size(data: bytes) -> tuple[float | None, float | None]:
    try:
        with Image.open(io.BytesIO(data)) as image:
            width, height = image.size
            dpi = image.info.get("dpi", (96, 96))
            if isinstance(dpi, tuple) and dpi and dpi[0]:
                return width * 72.0 / float(dpi[0]), height * 72.0 / float(dpi[1] or dpi[0])
            return float(width), float(height)
    except Exception:
        return None, None


def _convert_with_pillow(data: bytes, target: Path, fmt: str = "PNG") -> str | None:
    try:
        with Image.open(io.BytesIO(data)) as image:
            if getattr(image, "is_animated", False):
                image.seek(0)
            if image.mode not in {"RGB", "RGBA", "L", "LA"}:
                image = image.convert("RGBA")
            image.save(target, format=fmt)
        return None
    except Exception as exc:
        return str(exc)


def _convert_metafile(data: bytes, source_suffix: str, target: Path) -> str | None:
    """Try Inkscape, then the Windows GDI+ helper, to render EMF/WMF."""
    with tempfile.TemporaryDirectory(prefix="docx-typst-") as temporary:
        source = Path(temporary) / ("source" + source_suffix)
        source.write_bytes(data)
        target_tmp = Path(temporary) / "render.png"

        inkscape = shutil.which("inkscape")
        if inkscape:
            try:
                subprocess.run(
                    [inkscape, str(source), "--export-type=png", f"--export-filename={target_tmp}"],
                    check=True,
                    capture_output=True,
                    timeout=60,
                )
                if target_tmp.is_file() and target_tmp.stat().st_size:
                    target.write_bytes(target_tmp.read_bytes())
                    return None
            except (OSError, subprocess.SubprocessError):
                pass

        if os.name == "nt":
            script = (
                "Add-Type -AssemblyName System.Drawing;"
                f"$i=[System.Drawing.Image]::FromFile('{source.as_posix()}');"
                f"$i.Save('{target_tmp.as_posix()}',[System.Drawing.Imaging.ImageFormat]::Png);"
                "$i.Dispose()"
            )
            powershell = shutil.which("powershell") or shutil.which("pwsh")
            if powershell:
                try:
                    subprocess.run(
                        [powershell, "-NoProfile", "-NonInteractive", "-Command", script],
                        check=True,
                        capture_output=True,
                        timeout=60,
                    )
                    if target_tmp.is_file() and target_tmp.stat().st_size:
                        target.write_bytes(target_tmp.read_bytes())
                        return None
                except (OSError, subprocess.SubprocessError):
                    pass
    return "no EMF/WMF renderer available (install Inkscape or use Windows PowerShell)"


def prepare_media(
    data: bytes,
    suffix: str,
    content_type: str,
    destination: Path,
) -> tuple[Path | None, str | None]:
    """Write a renderer-friendly asset and return ``(path, warning)``.

    The original bytes are preserved when a conversion is necessary by writing
    the converted derivative to ``destination``; the caller records the
    original asset separately in the manifest.
    """
    suffix = suffix.lower() or extension_for("asset", content_type)
    destination.parent.mkdir(parents=True, exist_ok=True)
    if suffix in RASTER_DIRECT or suffix in SVG:
        destination = destination.with_suffix(suffix)
        destination.write_bytes(data)
        return destination, None
    if suffix in RASTER_CONVERT:
        target = destination.with_suffix(".png")
        warning = _convert_with_pillow(data, target, "PNG")
        if warning is None:
            return target, None
        return None, warning
    if suffix in METAFILE:
        target = destination.with_suffix(".png")
        warning = _convert_metafile(data, suffix, target)
        if warning is None:
            return target, None
        return None, warning
    # Unknown image type: preserve it; the Typst renderer may still support it.
    target = destination.with_suffix(suffix if suffix != ".bin" else ".png")
    target.write_bytes(data)
    if suffix == ".bin":
        return target, "unknown image format; original bytes preserved"
    return target, None


def svg_to_png(data: bytes, destination: Path) -> str | None:
    try:
        import cairosvg  # type: ignore
    except (ImportError, OSError) as exc:
        return f"cairosvg unavailable: {exc}"
    try:
        destination.parent.mkdir(parents=True, exist_ok=True)
        cairosvg.svg2png(bytestring=data, write_to=str(destination))
        return None
    except Exception as exc:
        return str(exc)


def rasterize_for_docx(path: Path) -> tuple[Path, str | None]:
    """Return a path python-docx can embed, rasterizing SVG when needed."""
    if path.suffix.lower() != ".svg":
        return path, None
    target = path.with_suffix(".png")
    if target.is_file():
        return target, None
    warning = svg_to_png(path.read_bytes(), target)
    if warning is None:
        return target, None
    return path, warning
