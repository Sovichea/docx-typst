# Agentic document workflow

`docx-typst` provides a file-based document representation for agents and scripts. Edit supported Typst constructs, compile the native document, and export an editable DOCX. Use the target Word renderer to verify the final layout.

## Preserve the source design

Before editing, read the document, its imported template and style files, and representative content. Identify the authoritative reference: an existing Word document, a Typst document, or a supplied template.

Preserve its fonts, weights, sizes, colors, margins, headings, paragraph spacing, table geometry, headers, footers, and native drawing code. A content update does not authorize a redesign. Obtain explicit approval for a design change; converter corrections should restore the reference appearance.

Keep text and native artwork at their source locations. Reuse the existing template helpers. Commit incremental, coherent changes with no more than 500 additions and deletions combined per commit.

## Recommended loop

```console
docx-to-typst source.docx --output working
typst compile working/main.typ working/preview.pdf
# Read working/conversion_report.json and inspect the reference.
# Edit supported content and assets while preserving the template.
typst compile working/main.typ working/preview.pdf
typst-to-docx working/main.typ --output revised.docx
# Render revised.docx with Microsoft Word and inspect every page.
```

1. **Establish a reference.** Keep the editable source and a freshly compiled reference PDF. For a Word-origin project, retain the original DOCX as well.
2. **Read conversion reports.** Review unsupported and approximated content after each conversion. A completed conversion can still contain layout losses.
3. **Edit supported constructs.** Preserve semantic paragraph, heading, list, table, and image wrappers. Arbitrary Typst functions are not evaluated by the exporter.
4. **Compile before export.** Check markup, asset paths, fonts, clipping, and native layout before diagnosing Word-specific differences.
5. **Export and render in Word.** When targeting Microsoft Word, use Microsoft Word to repaginate and render the DOCX. A Typst PDF or another office renderer cannot establish Word fidelity.
6. **Compare every rendered page.** Inspect the final export at readable resolution. Contact sheets help locate issues, but small text, rules, and table insets require closer views.
7. **Measure and correct discrepancies.** Compare corresponding content rather than page numbers alone: natural pagination and DOCX-only breaks can place the same section on different pages.
8. **Verify after the final edit.** Render affected pages again, then review the final document for pagination regressions. Run relevant tests and commit the converter changes incrementally.

For Word-origin projects, the following command checks supported content in the generated Typst project:

```console
docx-typst-validate source.docx working
```

This is a content-preservation check, not a comparison of the final DOCX render. Run export and visual verification separately.

## Microsoft Word render and review

When Word is installed locally, a render script can use Word automation to open the exported DOCX, repaginate, update fields in every story range, repaginate again, and export a PDF. Close the document and the automation-owned Word instance in a `finally` block. Preserve the user's open Word sessions.

If saving the DOCX after updating its fields, verify that saved file too. Word can normalize XML and font data during a save. Do not claim a graphical click was tested when verification only used XML, PDF output, or automation selection.

Check all of the following against the reference:

- Header rule to first heading, body to footer rule, and footer rule to footer text.
- Heading to decoration, decoration to body, and spacing between standalone text blocks.
- Header and footer right edges on portrait and landscape pages.
- Table frame position, column widths, merged-cell widths, fills, border colors, visible insets, row heights, and wrapping.
- Image dimensions, SVG bounds, clipping, and spacing above and below artwork.
- Font family, weight, size, script fallback, hyperlink appearance, list restarts, and page totals.
- Page, column, and section break markers with formatting marks shown in Word.

Measure gaps between equivalent visible edges, using the same method in both renders. Normalize the position of the section heading before comparing content below it. Record the measured residual difference; passing tests alone does not establish visual equivalence.

## Native artwork and headings

Keep cover art and diagrams as their original native drawing code. Wrap static artwork in `docx-static` so rendering happens during DOCX export. Request SVG where vector output is required, and verify the exported bounds and Office fallback. Keep cover text outside static artwork when it must remain editable in Word.

```typst
#docx-static(width: 210mm, height: 297mm, format: "svg")[
  // Existing native artwork, unchanged.
]
```

Complex heading show rules that draw decorations are not supported heading conversions. Author the decoration explicitly after the heading using `docx-rule`. Its Word representation uses a paragraph border. Keep heading typography in the existing template; do not recreate the decoration as a custom inline Word shape.

Preserve the reference's rule width, thickness, color, and gaps. `docx-rule` supports native `before` and `after` values plus Word-only `word-before` and `word-after` values for measured line-box differences.

## Paragraph spacing and helper styles

A simple imported text helper uses an existing paragraph style when its name matches that style: `contact-heading` uses `Contact Heading`, for example. Ensure that the intended style exists in `docx-styles`; otherwise the helper may inherit Normal spacing. Check the style's Word-specific spacing as well as its visible font treatment.

Word paragraph boxes include padding that differs from Typst's text edges. Correct that mapping without reducing font sizes or compressing multiline baseline spacing. Page-start heading metrics apply only to single-line headings that fit the available width; wrapped headings retain their original multiline spacing.

Use measured Word-only adjustments when a shared spacing value cannot reproduce both formats. Preserve the native reference value:

```typst
// Example only: choose both values from the document's reference measurements.
#docx-space(12pt, word: 8pt)
```

`docx-space` renders its first argument in Typst and exports the `word` gap. When omitted, `word` uses the same amount. Keep these adjustments local to the affected call site or in its existing style definition. Avoid repeated guessed spacers or empty paragraphs used to force a page count.

## Single-table panels

Build a contact panel as one table with full-width merged rows for names and two columns for labels and values. Do not wrap a contact grid in another shaded table: nested tables introduce extra cell margins and end-paragraph spacing.

Use one shared source for fills, column widths, merges, and per-cell padding. `docx-table` supports individual edge insets:

```typst
#docx-table(
  columns: 2,
  rows: 2,
  widths: (117pt, 1fr),
  merges: ((0, 0, 1, 2),),
  stroke: none,
  inset: 0pt,
  fills: ((0, 0, "#FFF9E7"), (1, 0, "#FFF9E7"), (1, 1, "#FFF9E7")),
  cell-insets: (
    (0, 0, (top: 18pt, bottom: 18pt, left: 18pt, right: 18pt)),
    (1, 0, (top: 6pt, bottom: 18pt, left: 18pt, right: 14pt)),
    (1, 1, (top: 6pt, bottom: 18pt, left: 0pt, right: 18pt)),
  ),
  cells: ([Company name], [Phone], [Contact number]),
)
```

The numbers above illustrate the construct; inherit the actual document's values. Preserve the panel's text size when moving grid content into a table: an inherited table text rule can otherwise shrink it.

The exporter compensates cell padding for Word's text line boxes. Verify both the resulting row heights and visible padding, including cells with a zero inset on one edge. Verify that a merged row has the sum of its column widths and that long labels still fit. Native Word table shading has square corners; Typst clipping does not give Word tables rounded corners.

A table rule disabling justification must not override an explicitly right-aligned paragraph style, including header and footer labels. Native Typst links should retain their inherited appearance; imported Word hyperlink styling must still survive a roundtrip.

## Breaks and pagination

Page, column, and section break markers should occupy a full separate paragraph line in Word. Attaching a marker to body text can stretch the last justified line or make the marker appear beside the text. A tiny marker paragraph may still look attached; inspect with formatting marks enabled.

Use `docx-pagebreak()` when a break is required only in Word; it has no visible effect in Typst. Ordinary `pagebreak()` affects native pagination and export. Retain section changes required for page size, orientation, margins, or headers and footers.

Avoid promising identical page counts when Word's flow differs. Correct the underlying font, wrapping, spacing, or table error before changing deliberate pagination. Do not shrink typography to force a match.

## Fonts and final delivery

Font substitution can change both glyph size and row geometry, particularly for Khmer and other scripts. Check the font used in the Word render, not just the family named in the source. The exporter can embed available regular Khmer script fonts when their embedding rights permit it; verify that the font is actually used after Word opens and saves the document.

Before delivery, confirm that the requested source and DOCX were updated, the final render was inspected, relevant tests passed, and the working tree contains only intended changes. Report material remaining limitations and the scope of verification. Include the commit identifier when converter or workflow changes were committed.
