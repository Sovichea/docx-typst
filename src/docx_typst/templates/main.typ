#import "docx-typst.typ": *

// Edit this table to change document-wide style defaults. The reverse
// converter reads the same table when exporting to DOCX.
#let docx-styles = (
  "Normal": (
    font: "Calibri",
    size: 11pt,
    after: 8pt,
    line-spacing: 1.08,
  ),
  "Heading 1": (
    font: "Calibri Light",
    size: 18pt,
    bold: true,
    fill: rgb("#2F5496"),
    before: 18pt,
    after: 6pt,
    keep-with-next: true,
  ),
  "List Bullet": (
    font: "Calibri",
    size: 11pt,
    after: 2pt,
    line-spacing: 1.08,
  ),
)

#show: docx-document.with(
  title: "A new document",
  author: "Your name",
  page: (paper: "a4", margin: 72pt),
)

#docx-heading(styles: docx-styles, level: 1, style: "Heading 1")[A new document]

#docx-para(styles: docx-styles, style: "Normal")[
  Edit this text, then export with
  #raw("typst-to-docx main.typ --output document.docx").
]

#docx-item(styles: docx-styles, kind: "bullet", level: 0, style: "List Bullet", marker: "•")[First item]

#docx-item(styles: docx-styles, kind: "bullet", level: 0, style: "List Bullet", marker: "•")[Second item]

#docx-table(
  columns: 2,
  header-rows: 1,
  cells: (
    [#docx-para(styles: docx-styles, style: "Normal")[Name]],
    [#docx-para(styles: docx-styles, style: "Normal")[Value]],
    [#docx-para(styles: docx-styles, style: "Normal")[Alpha]],
    [#docx-para(styles: docx-styles, style: "Normal")[1]],
  ),
)
