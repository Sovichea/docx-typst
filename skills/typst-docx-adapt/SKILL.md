---
name: typst-docx-adapt
description: Adapt an existing Typst document for editable DOCX export with docx-typst while preserving its reference design and native artwork. Use when a general .typ project must be made converter-compatible; not for authoring a new document from scratch.
---

# Adapt Typst for docx-typst

Produce a compatible Typst entrypoint and an editable DOCX. Preserve the original document and its appearance; successful compilation alone does not establish conversion fidelity.

## Establish the reference and tool version

- Read the source entrypoint, imported templates, style files, and representative content before editing. Compile the original to a reference PDF.
- Locate the available `docx-typst` package or checkout and inspect its current helper module and CLI. Do not assume an installed release has every construct mentioned here. In a checkout, read `docs/agentic-document-workflow.md` and the relevant supported-subset sections.
- Create a reconstruction entrypoint alongside the original when practical, preserving relative imports and assets. Keep the original intact unless the user asks to replace it. Copy the current helper module into a new project path with `docx-typst module --output docx-typst.typ`; do not overwrite a customized module without inspecting it.
- Preserve fonts, sizes, weights, colors, margins, line and paragraph spacing, page furniture, and drawing code. Conversion is permission to adapt representation, not redesign. Resolve any necessary design deviation with the user before implementing it.

## Adapt the representation

Read [Conversion patterns](references/conversion-patterns.md) when mapping content, artwork, tables, or pagination.

1. Classify each construct as supported editable content, static native artwork, or an unsupported layout/programming dependency. Do not assume imported macros or show rules execute in the exporter.
2. Build `docx-styles` from the reference's actual styles. Keep the existing visual template; use the matching paragraph styles for editable text. A simple text helper such as `contact-heading` can use the existing `Contact Heading` style.
3. Use supported semantic paragraphs, headings, lists, tables, and links. Keep text local to its source call site. Preserve list restarts and table alignment rules, including explicit right alignment in headers and footers.
4. Keep native artwork code intact inside `docx-static`, with explicit dimensions and SVG export when required. Separate editable cover text and logos from static decoration; use supported canvas/position constructs only when needed.
5. Replace complex heading decoration show rules with an explicit `docx-rule` after the heading. Preserve the reference's typography and rule geometry.
6. Represent a contact panel with one table, merged name rows, fills, and per-cell insets. Avoid a shaded outer table containing an inner contact grid. Retain the original text size when table text rules would change it.
7. Use measured Word-only spacing properties to reconcile line boxes. Keep the native values intact; do not shrink fonts or compress multiline leading to force pagination. Avoid guessed spacers and manual OOXML shapes as heading decorations.

## Verify the adapted document

```console
typst compile main-docx-typst.typ adapted.pdf
typst-to-docx main-docx-typst.typ --output adapted.docx
```

- Compare the adapted native render with the original before diagnosing Word differences. Review every conversion diagnostic; flattening may lose fills, insets, positioning, or styles even when text remains.
- Render the DOCX with Microsoft Word when that is the target application. Repaginate, update all story fields, repaginate again, and export PDF. Preserve the user's open documents. Inspect the saved DOCX as well if Word saves it during verification.
- Inspect every final page. Measure corresponding visible edges for header-to-heading, rule-to-body, text-block, and body-to-footer gaps. Compare corresponding sections when page counts differ. Review table fills, borders, frame position, padding, row heights, merged widths, image bounds, fonts, and script substitution at readable zoom.
- Check page, column, and section markers with formatting marks enabled: each marker belongs on a full separate paragraph line. Preserve required orientation sections. Use DOCX-only breaks only for intentional Word pagination.
- Re-render after corrections and run relevant converter tests when code changes. Keep commits coherent and below 500 changed lines each.

Deliver the adapted source, DOCX, and verification result. State remaining limitations, including Word tables' square corners or unsupported editability. Do not claim page-by-page or click testing that was not performed.
