# Authoring patterns

## Initialize or integrate

```console
docx-typst init project
```

The starter includes `main.typ`, `docx-typst.typ`, and assets. Inspect it before editing. For an approved external template, add the helper module to a new filename and integrate the template's existing styles; do not replace the template with the starter design.

The following is a small working example for a document with no supplied visual reference. Replace its style values with the approved design rather than treating them as defaults for every document.

```typst
#import "docx-typst.typ": *

#let docx-styles = (
  "Normal": (font: "Calibri", size: 11pt, after: 8pt, line-spacing: 1.08),
  "Heading 1": (
    font: "Calibri", size: 18pt, bold: true,
    before: 18pt, after: 6pt, keep-with-next: true,
  ),
  "List Bullet": (font: "Calibri", size: 11pt, after: 3pt),
)

#show: docx-document.with(
  title: "Project brief",
  author: "Document author",
  page: (paper: "a4", margin: 72pt),
)

#docx-heading(styles: docx-styles, level: 1, style: "Heading 1")[Project brief]
#docx-para(styles: docx-styles, style: "Normal")[
  The purpose, scope, and expected result remain editable in Word.
]
#docx-item(styles: docx-styles, kind: "bullet", style: "List Bullet")[First deliverable]
#docx-item(styles: docx-styles, kind: "bullet", style: "List Bullet")[Second deliverable]

#docx-table(
  columns: 2,
  rows: 2,
  widths: (120pt, 1fr),
  header-rows: 1,
  stroke: 0.5pt + rgb("DCE3DF"),
  inset: 6pt,
  fills: ((0, 0, "#FFF9E7"), (0, 1, "#FFF9E7")),
  cells: ([*Deliverable*], [*Acceptance*], [Report], [Reviewed with the owner]),
)
```

## Text, headings, and spacing

Use named styles rather than repeating font overrides in body content. A simple imported helper named `contact-heading` uses an existing `Contact Heading` paragraph style; ensure that the style exists when relying on that mapping.

Place a heading's decoration explicitly after the heading with `docx-rule`. Give it the approved width, height, fill, and spacing. Its native `before`/`after` values can differ from Word-only `word-before`/`word-after` values when measurement demonstrates a line-box difference.

`word-before` and `word-after` also apply to paragraph style definitions. For an isolated explicit gap:

```typst
// Example only: measure the native and Word gaps before choosing values.
#docx-space(12pt, word: 8pt)
```

These metrics preserve the native layout. Do not tune font sizes or multiline leading to reproduce another document's page count. Use page-start heading metrics only when the current converter and actual layout support them.

## One table for a contact panel

Create merged full-width rows for the company name and script-language name, then two columns for contact labels and values. Each merged cell occupies one entry in `cells`; omit entries for the positions it covers. Fill every logical cell and set individual edge insets where needed.

```typst
#docx-table(
  columns: 2,
  rows: 3,
  widths: (117pt, 1fr),
  merges: ((0, 0, 1, 2),),
  stroke: none,
  inset: 0pt,
  keep-together: true,
  fills: (
    (0, 0, "#FFF9E7"),
    (1, 0, "#FFF9E7"), (1, 1, "#FFF9E7"),
    (2, 0, "#FFF9E7"), (2, 1, "#FFF9E7"),
  ),
  cell-insets: (
    (0, 0, (top: 18pt, bottom: 18pt, left: 18pt, right: 18pt)),
    (1, 0, (top: 6pt, bottom: 6pt, left: 18pt, right: 14pt)),
    (1, 1, (top: 6pt, bottom: 6pt, left: 0pt, right: 18pt)),
    (2, 0, (top: 6pt, bottom: 18pt, left: 18pt, right: 14pt)),
    (2, 1, (top: 6pt, bottom: 18pt, left: 0pt, right: 18pt)),
  ),
  cells: ([*Company name*], [*Phone*], [Contact number], [*Email*], [Email address]),
)
```

Choose the reference's actual dimensions and colors. Preserve text size across native table show rules and Word styles. Check row heights, visible padding, frame alignment, merged widths, and long-label wrapping in both renders. Word table shading has square corners.

## Native drawings and editable covers

Keep the original drawing calls directly in a `docx-static` body. With `format: "svg"`, native vector artwork is converted during DOCX export. Verify its size and Office fallback; explicit width and height reserve a bounded frame.

Keep editable titles and cover metadata outside the static art. Supported `docx-canvas`/`docx-position` constructs can combine page-positioned editable text and static decorative layers. Inspect the installed template for their exact parameters before using them.

## Review gates

Review the native PDF, export diagnostics, and Word-rendered PDF separately. Check every final page and inspect formatting marks for break markers. Compare equivalent visible edges and matching sections rather than assuming equal page counts.

Font families named in source can still be substituted by Word. Verify script glyphs and wrapping after a Word save; use supported font embedding only where available and permitted. Report any missing renderer, unsupported construct, font substitution, or visual limitation instead of claiming an unperformed review.
