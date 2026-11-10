// Minimal, neutral document primitives used by generated docx-typst projects.
// Import with: #import "template.typ": *
//
// The module intentionally implements a static subset. It does not evaluate
// arbitrary Word behaviour; it renders the semantic document model produced by
// docx-typst.

#let docx-spec(styles, style) = {
  if style != none and style in styles { styles.at(style) } else { (:) }
}

#let docx-align(name) = {
  if name == "center" { center }
  else if name == "right" { right }
  else { left }
}

#let docx-leading(spec, size, line-spacing, line-spacing-pt) = {
  let multiple = if line-spacing != auto { line-spacing } else { spec.at("line-spacing", default: none) }
  let exact = if line-spacing-pt != auto { line-spacing-pt } else { spec.at("line-spacing-pt", default: none) }
  if exact != none { calc.max(0pt, exact - size * 1.2) }
  else if multiple != none { calc.max(0pt, (multiple - 1) * size * 1.2) }
  else { 0.65em }
}

#let docx-document(
  title: none,
  author: none,
  subject: none,
  keywords: none,
  page: (:),
  header: none,
  footer: none,
  body,
) = {
  set document(title: title, author: author)
  let page-settings = page
  if header != none { page-settings.insert("header", header) }
  if footer != none { page-settings.insert("footer", footer) }
  set std.page(..page-settings)
  set text(font: "Calibri", size: 11pt, fill: black)
  set par(leading: 0.65em)
  show heading: set block(above: 0pt, below: 0pt)
  body
}

#let docx-para(
  styles: (:),
  style: "Normal",
  align: auto,
  before: auto,
  after: auto,
  line-spacing: auto,
  line-spacing-pt: auto,
  left-indent: auto,
  right-indent: auto,
  first-line-indent: auto,
  keep-with-next: auto,
  keep-lines: auto,
  page-break-before: auto,
  bookmarks: (),
  body,
) = {
  let spec = docx-spec(styles, style)
  let font = spec.at("font", default: "Calibri")
  let size = spec.at("size", default: 11pt)
  let fill = spec.at("fill", default: black)
  let weight = if spec.at("bold", default: false) { "bold" } else { "regular" }
  let font-style = if spec.at("italic", default: false) { "italic" } else { "normal" }
  let align-name = if align != auto { align } else { spec.at("align", default: "left") }
  let before-value = if before != auto { before } else { spec.at("before", default: 0pt) }
  let after-value = if after != auto { after } else { spec.at("after", default: 8pt) }
  let left-value = if left-indent != auto { left-indent } else { spec.at("left-indent", default: 0pt) }
  let right-value = if right-indent != auto { right-indent } else { spec.at("right-indent", default: 0pt) }
  let first-value = if first-line-indent != auto { first-line-indent } else { spec.at("first-line-indent", default: 0pt) }
  let leading = docx-leading(spec, size, line-spacing, line-spacing-pt)
  let justify = spec.at("justify", default: false) or align-name == "justify"
  let break-before = if page-break-before != auto { page-break-before } else { spec.at("page-break-before", default: false) }
  let sticky = if keep-with-next != auto { keep-with-next } else { spec.at("keep-with-next", default: false) }
  let unbreakable = if keep-lines != auto { keep-lines } else { spec.at("keep-lines", default: false) }
  if break-before { pagebreak(weak: true) }
  set text(font: font, size: size, fill: fill, weight: weight, style: font-style)
  set par(justify: justify, leading: leading, first-line-indent: calc.max(0pt, first-value))
  let content = body
  if spec.at("underline", default: false) { content = underline(content) }
  if spec.at("strike", default: false) { content = strike(content) }
  if spec.at("small-caps", default: false) { content = smallcaps(content) }
  if spec.at("all-caps", default: false) { content = upper(content) }
  if bookmarks.len() > 0 {
    for name in bookmarks { docx-anchor(name) }
  }
  block(
    above: before-value,
    below: after-value,
    sticky: sticky,
    breakable: not unbreakable,
    width: 100%,
  )[
    #pad(left: left-value, right: right-value)[
      #std.align(docx-align(align-name))[#content]
    ]
  ]
}

#let docx-heading(
  styles: (:),
  level: 1,
  style: none,
  align: auto,
  before: auto,
  after: auto,
  line-spacing: auto,
  line-spacing-pt: auto,
  left-indent: auto,
  right-indent: auto,
  first-line-indent: auto,
  keep-with-next: auto,
  keep-lines: auto,
  page-break-before: auto,
  bookmarks: (),
  body,
) = {
  let spec = docx-spec(styles, style)
  let font = spec.at("font", default: "Calibri Light")
  let size = spec.at("size", default: (11pt + level * 2pt))
  let fill = spec.at("fill", default: black)
  let weight = if spec.at("bold", default: true) { "bold" } else { "regular" }
  let font-style = if spec.at("italic", default: false) { "italic" } else { "normal" }
  let align-name = if align != auto { align } else { spec.at("align", default: "left") }
  let before-value = if before != auto { before } else { spec.at("before", default: 12pt) }
  let after-value = if after != auto { after } else { spec.at("after", default: 6pt) }
  let left-value = if left-indent != auto { left-indent } else { spec.at("left-indent", default: 0pt) }
  let right-value = if right-indent != auto { right-indent } else { spec.at("right-indent", default: 0pt) }
  let leading = docx-leading(spec, size, line-spacing, line-spacing-pt)
  let break-before = if page-break-before != auto { page-break-before } else { spec.at("page-break-before", default: false) }
  if break-before { pagebreak(weak: true) }
  if bookmarks.len() > 0 {
    for name in bookmarks { docx-anchor(name) }
  }
  let heading-body = heading(level: level, outlined: true, bookmarked: true, {
    set text(font: font, size: size, fill: fill, weight: weight, style: font-style)
    set par(leading: leading)
    body
  })
  block(above: before-value, below: after-value, sticky: true)[
    #pad(left: left-value, right: right-value)[
      #std.align(docx-align(align-name))[#heading-body]
    ]
  ]
}

#let docx-item(
  styles: (:),
  kind: "bullet",
  level: 0,
  style: "List Bullet",
  marker: auto,
  num-id: none,
  before: auto,
  after: auto,
  line-spacing: auto,
  line-spacing-pt: auto,
  left-indent: auto,
  first-line-indent: auto,
  body,
) = {
  let spec = docx-spec(styles, style)
  let font = spec.at("font", default: "Calibri")
  let size = spec.at("size", default: 11pt)
  let fill = spec.at("fill", default: black)
  let weight = if spec.at("bold", default: false) { "bold" } else { "regular" }
  let font-style = if spec.at("italic", default: false) { "italic" } else { "normal" }
  let before-value = if before != auto { before } else { spec.at("before", default: 0pt) }
  let after-value = if after != auto { after } else { spec.at("after", default: 0pt) }
  let leading = docx-leading(spec, size, line-spacing, line-spacing-pt)
  let mark = if marker == auto {
    if kind == "number" { "1." } else { "\u{2022}" }
  } else { marker }
  let left-value = if left-indent != auto { left-indent } else { level * 18pt }
  let hanging-value = if first-line-indent != auto and first-line-indent < 0pt {
    -first-line-indent
  } else { 18pt }
  let marker-column = calc.max(0pt, left-value - hanging-value)
  set text(font: font, size: size, fill: fill, weight: weight, style: font-style)
  set par(leading: leading)
  block(above: before-value, below: after-value, width: 100%)[
    #pad(left: marker-column)[
      #grid(
        columns: (hanging-value, 1fr),
        align: (right, left),
        gutter: 4pt,
        text[#mark],
        body,
      )
    ]
  ]
}

#let docx-inline-image(path, width: auto, height: auto, alt: none) = {
  image(path, width: width, height: height, alt: alt)
}

#let docx-image(
  path,
  width: auto,
  height: auto,
  alt: none,
  align: center,
  caption: none,
) = {
  let picture = image(path, width: width, height: height, alt: alt)
  if caption != none {
    std.align(align)[#figure(picture, caption: caption)]
  } else {
    std.align(align)[#picture]
  }
}

#let docx-unsupported-image(alt, detail: none) = {
  box(
    width: 100%,
    height: 48pt,
    fill: luma(238),
    stroke: 0.5pt + luma(190),
    radius: 2pt,
  )[
    #align(center + horizon)[
      #text(size: 8pt, fill: luma(90))[Unsupported image: #alt]
    ]
  ]
}

#let docx-code(text, language: none, style: none) = {
  let args = (block: true)
  if language != none { args.insert("lang", language) }
  raw(text, ..args)
}

#let docx-math(text) = {
  text(font: "New Computer Modern Math")[#text]
}

#let docx-table(
  columns: 1,
  rows: auto,
  widths: auto,
  header-rows: 0,
  merges: (),
  style: none,
  align: "left",
  stroke: 0.5pt + luma(180),
  inset: 5pt,
  cells: (),
) = {
  let ncols = if type(columns) == int { columns } else { columns.len() }
  let cols = if widths == auto { columns } else { widths }
  let nrows = if rows == auto { calc.ceil(cells.len() / ncols) } else { rows }
  let merge-map = (:)
  let covered = (:)
  for m in merges {
    let row = m.at(0)
    let col = m.at(1)
    let rowspan = m.at(2)
    let colspan = m.at(3)
    merge-map.insert(str(row) + ":" + str(col), (rowspan, colspan))
    for r in range(row, row + rowspan) {
      for c in range(col, col + colspan) {
        if r != row or c != col {
          covered.insert(str(r) + ":" + str(c), true)
        }
      }
    }
  }
  let cell-align = docx-align(align)
  let rows = ()
  let index = 0
  for r in range(0, nrows) {
    let row-cells = ()
    for c in range(0, ncols) {
      let key = str(r) + ":" + str(c)
      if key in covered { continue }
      if index >= cells.len() { break }
      let cell = cells.at(index)
      index += 1
      let span = merge-map.at(key, default: (1, 1))
      if span.at(0) > 1 or span.at(1) > 1 {
        row-cells.push(table.cell(
          rowspan: span.at(0),
          colspan: span.at(1),
          align: cell-align,
        )[#cell])
      } else {
        row-cells.push(table.cell(align: cell-align)[#cell])
      }
    }
    rows.push(row-cells)
  }
  let header = ()
  let body = ()
  for (i, row) in rows.enumerate() {
    if i < header-rows { header = header + row } else { body = body + row }
  }
  let all = if header.len() > 0 {
    (table.header(repeat: true, ..header),) + body
  } else { body }
  table(columns: cols, stroke: stroke, inset: inset, ..all)
}

#let docx-page-number() = context counter(page).display()

#let docx-page-count() = context {
  let total = counter(page).final().first()
  [#total]
}

#let docx-field(name) = {
  [#name]
}

#let docx-allcaps(body) = upper(body)

#let docx-anchor(name) = {
  [#label(name)]
}

#let docx-anchor-link(name, body) = {
  link(label(name))[#body]
}

#let docx-unsupported(kind, detail: "") = {
  block(
    width: 100%,
    fill: luma(248),
    stroke: 0.5pt + luma(200),
    inset: 6pt,
  )[
    #text(size: 8pt, fill: luma(90))[Unsupported content: #kind #detail]
  ]
}
