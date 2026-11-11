---
name: typst-docx-author
description: Author new Typst documents designed for editable DOCX export with docx-typst. Use when creating a proposal, report, or other new .typ document that must support Word conversion; not for adapting an existing general Typst project.
---

# Author docx-typst-compatible Typst

Create the editable Typst source and its Word export together, using supported constructs from the beginning. Read [Authoring patterns](references/authoring-patterns.md) when selecting the project structure, text styles, tables, or artwork.

## Start from the right template

- Read any supplied reference and imported style files before authoring. Preserve their fonts, weights, colors, margins, line/paragraph spacing, and layout conventions. A request for content is not permission to redesign a template.
- Locate the installed converter or checkout and inspect its helper module and CLI. Check construct availability for that version. In a checkout, consult `docs/agentic-document-workflow.md` for the current review procedure.
- With no supplied template, `docx-typst init project` provides a compatible starter. Use a new directory; do not overwrite user files with `--force`. When integrating an approved template, copy the helper module into a new path using `docx-typst module --output project/docx-typst.typ`.
- Define the document's styles once in `docx-styles`. Keep user-facing content at its actual source call sites. Prefer supported semantic constructs over new document-specific rendering macros.

## Author for both renderers

- Use `docx-para`, `docx-heading`, and `docx-item` with the appropriate named style. Native supported markup is also available, but verify how its style is resolved. A simple imported text helper can use a matching existing paragraph style.
- Use `docx-table` for structured layouts requiring predictable cell geometry. Express widths, merges, fills, border strokes, and cell insets explicitly where the design requires them. Build contact panels as one table rather than a panel table containing a grid.
- Keep table text size explicit in the existing style system so an inherited native table rule does not shrink it. Keep labels non-justified where required, while preserving explicitly right-aligned header/footer paragraphs.
- Author decorative heading rules as separate `docx-rule` calls. Complex heading show rules that create decorations are not supported conversion constructs. Match the template's typography and spacing; do not replace rules with custom inline Word shapes.
- Keep illustrations and native drawing code inside `docx-static` only when the exported artwork can be static. Prefer SVG for vector artwork. Keep text outside static layers when it must remain editable; use supported cover canvas/position constructs when needed.
- Use standard image calls with deliberate sizes and alt text. Check bounds and proportions instead of forcing oversized artwork into a smaller frame.
- Keep links and page-number fields semantic. Use dynamic total fields. Ensure independent numbered lists restart correctly.
- Use explicit page breaks only when they serve the content. `docx-pagebreak()` is available for intentional Word-only pagination. Keep all page, column, and section markers on separate full paragraph lines in Word.

## Iterate against real Word output

```console
typst compile project/main.typ project/preview.pdf
typst-to-docx project/main.typ --output project/document.docx
```

Read the conversion report, then render the DOCX in Microsoft Word when Word is the target. Repaginate, update fields across all stories, repaginate again, and render PDF. Preserve the user's open documents. Verify the saved DOCX too if Word saves it during review.

Inspect every final page at readable zoom. Compare the native and Word renders for header-to-heading, rule-to-body, text-block, and body-to-footer gaps; right alignment; table fill/borders/insets/heights; image clipping; fonts; and page totals. Inspect corresponding sections when their page numbers differ.

Use measured Word-only style spacing or `docx-space` for actual line-box differences, keeping the native values intact. Do not copy another document's calibration numbers. Preserve font sizes and multiline baseline spacing; do not compress typography to make a page count match.

Verify script fonts in the rendered output, including Khmer. Word may substitute a font despite the source family name. Font embedding is limited by available font files and embedding rights; verify the font actually used after opening and saving in Word.

Re-render after final corrections. For converter changes, run relevant tests and commit coherently with at most 500 additions and deletions per commit. Deliver the source, DOCX, review result, and material limitations. Rounded native panels do not make Word table corners round, and successful compilation does not prove visual equivalence.
