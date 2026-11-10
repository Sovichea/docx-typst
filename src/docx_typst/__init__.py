"""Bidirectional Word and Typst conversion for a documented static subset."""

from __future__ import annotations

from typing import Any

__version__ = "0.1.0"


def docx_to_typst(*args: Any, **kwargs: Any) -> dict[str, Any]:
    """Convert a DOCX file to a Typst project."""
    from .docx_to_typst import convert

    return convert(*args, **kwargs)


def typst_to_docx(*args: Any, **kwargs: Any) -> dict[str, Any]:
    """Convert a supported Typst source file to DOCX."""
    from .typst_to_docx import convert

    return convert(*args, **kwargs)


__all__ = ["__version__", "docx_to_typst", "typst_to_docx"]
