# Supported subset

`docx-typst` is a semantic converter, not a complete clone of Word or Typst. This page describes the content that roundtrips predictably and the content that is reported as approximated or unsupported.

## Document structure

| Word content | Typst representation | Reverse behaviour |
|---|---|---|
| Body paragraphs | `#docx-para(style: "...")[...]` | Paragraph with the named style and direct formatting |
| Headings | `#docx-heading(level: ..., style: "...")[...]` | Heading level and matching Word style |
| Bullet and numbered lists | `#docx-item(kind: ..., level: ..., marker: ...)[...]` | `List Bullet` or `List Number` paragraph |
| Page breaks | `#pagebreak()` | Page break |
| Sections | `#pagebreak()` followed by `#set page(...)` | New Word section with page setup and headers/footers |
| Empty paragraphs | Empty `#docx-para(...)[]` | Empty paragraph |

The generated `docx-styles` table records effective paragraph style properties. It controls Typst rendering and is read back to recreate Word styles. Edit the table to change document-wide defaults.

## Character formatting

The following direct run properties are supported in both directions:

- bold and italic
- underline and strikethrough
- superscript and subscript
- small caps and all caps
- font name and size
- font color and highlight color
- external hyperlinks
- inline code
- line breaks and tabs

Character styles are resolved into their effective formatting. The style name itself is not guaranteed to survive a roundtrip unless it is also a paragraph style.

## Tables

Supported table properties:

- rectangular grids and explicit column widths
- header rows that repeat across pages
- horizontal and vertical cell merges
- cell background fills
- common single-line borders
- nested block content inside cells

Complex border combinations, diagonal borders, table indentation, and conditional table-style formatting are approximated with a default table grid.

## Images

Supported image handling:

- PNG and JPEG are preserved directly.
- SVG is preserved for Typst and rasterized for DOCX when Cairo support is available.
- GIF, BMP, TIFF, and WebP are converted to PNG for Typst.
- EMF and WMF are converted through Inkscape or the Windows GDI+ path when available.
- Duplicate images are stored once and referenced by content hash.
- Original bytes are retained when a derived rendering asset is created.

Alt text is carried into Typst. Crop rectangles and rotation are recorded but not reproduced; the conversion report flags them. Floating images are emitted as centered block images. Text boxes are flattened into their text content.

## Headers, footers, and page setup

The first section carries the document header and footer. Later sections can override or inherit them. Page numbers and page counts use Word fields in DOCX and `counter(page)` in Typst.

Page setup maps paper size, margins, column count, orientation, and section breaks. Continuous and even/odd section starts are approximated with a page break. First-page and even-page headers are reported when they differ from the default header.

## Word features reported as unsupported or approximated

The converter emits diagnostics rather than silently dropping difficult content:

- remaining comments and replies
- tracked-change metadata (accepted insertions are kept, deletions are omitted)
- native chart data
- SmartArt
- OLE objects and embedded files
- complex equations
- content controls (their content is flattened)
- footnotes (rendered as Typst footnotes, exported as superscript text)
- endnotes
- floating shape geometry and text wrapping
- complex image masks and crops

## Static Typst accepted by the reverse converter

`typst-to-docx` does not evaluate arbitrary Typst. It understands:

- `#docx-para`, `#docx-heading`, `#docx-item`, `#docx-image`, `#docx-table`
- `#docx-code`, `#docx-page-number`, `#docx-page-count`, `#docx-field`
- native headings (`= Heading`), bullet lists (`- item`), and numbered lists (`+ item`)
- `*strong*`, `_emphasis_`, inline code, and escaped markup
- `#link("url")[...]`, `#underline[...]`, `#strike[...]`, `#super[...]`, `#sub[...]`, `#smallcaps[...]`
- `#text(font: ..., size: ..., fill: ..., weight: ..., style: ...)[...]`
- `#highlight(fill: ...)[...]`, `#raw("...")`
- `#image(...)`, `#figure(image(...), caption: [...])`, and `#table(...)`
- `#set page(...)`, `#set document(...)`, and `#set text(...)`
- `#show: docx-document.with(...)`

Unknown function bodies are flattened into their content when possible. Unknown calls without content are preserved as text and reported.

## Conversion reports

Every conversion writes `conversion_report.json`. Each diagnostic includes an issue name, outcome (`approximated`, `unsupported`, or `skipped`), location, fallback behaviour, reason, and expected loss.

Run `docx-typst-validate source.docx project` to compare the source and generated project before editing or delivery.
