"""Bidirectional Word and Typst conversion for a documented static subset."""

from __future__ import annotations

__version__ = "0.1.0"

from .docx_to_typst import convert as docx_to_typst
from .typst_to_docx import convert as typst_to_docx

__all__ = ["__version__", "docx_to_typst", "typst_to_docx"]
