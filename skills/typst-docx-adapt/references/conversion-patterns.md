# Conversion patterns

## Map native content deliberately

| Existing Typst | Compatible representation and verification |
|---|---|
| Ordinary text and headings | Native supported markup or `docx-para` / `docx-heading` with existing `docx-styles`; verify style and spacing. |
| Imported simple text helper | Match its hyphenated name to the existing title-cased paragraph style; verify the helper is recognized and the style actually applies. |
| Arbitrary macro or show-rule body | Adapt to supported constructs; unknown bodies may flatten and unknown calls may become literal text. |
| Heading with a decorative rectangle | Keep the heading editable; author `docx-rule` separately and match its before/after gaps. |
| Cover drawings or diagrams | Preserve native code in `docx-static`; keep editable text outside the static layer. |
| Positioned cover text | Use `docx-canvas` and `docx-position` if the installed module supports them. Verify positions and source preservation after a Word save. |
| Shaded contact grid inside a panel | Use a single `docx-table`, with merged rows, cell fills, and `cell-insets`. |
| Native header/footer grids | Retain right-aligned paragraph styles and responsive widths; check portrait and landscape right edges. |

## Artwork stays native in the reconstruction

`docx-static` renders its body only during DOCX export. Do not replace drawing code with an exported PNG or SVG path in the editable source. The generated export may contain SVG plus an Office fallback; inspect both the intended display size and visible bounds.

```typst
#docx-static(width: 210mm, height: 297mm, format: "svg")[
  // Original cover drawing code goes here, unchanged.
]
```

Retain necessary top-level imports and correctly rebased asset paths. A fixed width and height reserve a bounded frame; verify that the artwork fits rather than clipping it to hide overflow. Do not rasterize an entire editable page to obtain a visual match.

## Styles and spacing

The reconstruction's `docx-styles` should reflect the reference, including font family, weight, fill, alignment, margins, and line spacing. Check both the visible native style and the Word metadata; a named style that the helper never uses cannot fix Word spacing.

`word-before` and `word-after` are Word-only style values. `docx-rule` also accepts these values in addition to native `before` and `after`. Use the appropriate scope: style-level for a shared treatment, local call-site adjustments for an isolated gap.

```typst
// Illustration only: derive these values from the actual reference.
#docx-space(12pt, word: 8pt)
```

The first value remains the native Typst gap. The second compensates Word's different line box. Do not reuse another document's calibration values. Page-start `word-page-before` and `word-page-line-spacing-pt` metrics apply only to native headings that fit on one line; preserve multiline baseline spacing.

## One-table contact panels

Use `docx-table` with full-width merged company-name rows followed by two contact columns. Set `stroke: none`, the reference fill on every logical cell, and explicit `cell-insets` for the four edges. The inset syntax is:

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

These are example dimensions, not required styling. Read the reference's actual values. A merged cell consumes one content entry; omit cells covered by the merge. Confirm that merged widths equal the column-width sum and long labels still fit after padding.

The Word writer compensates visible padding for font line boxes. Check row heights and text offsets, including zero-inset edges and wrapped addresses. Word table shading remains square even if the native Typst table is clipped to rounded corners; disclose that limitation.

## Pagination and fonts

- `pagebreak()` affects Typst and DOCX; `docx-pagebreak()` affects DOCX only. Confirm these helpers exist in the current module.
- Page and column markers belong on separate full paragraph lines. Section markers must also be separate to avoid stretching justified body text.
- Retain page size/orientation sections and inherited headers and footers. Page totals must be fields, not hard-coded text.
- Verify the font used by Word's render. Source font names do not prove that Word avoided substitution. Available regular Khmer fonts can be embedded when permitted; verify the actual result after a Word save.
- Keep links editable while retaining native text color and underline choices. Imported Word hyperlink formatting remains a separate roundtrip requirement.
