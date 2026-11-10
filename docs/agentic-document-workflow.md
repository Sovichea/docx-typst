# Agentic document workflow

`docx-typst` is a file-based document IR for workflows where an agent or script needs to inspect and revise a Word document. The agent edits readable Typst source instead of manipulating OOXML package parts or driving Word through UI automation.

## Recommended loop

```console
docx-to-typst source.docx --output working
typst compile working/main.typ working/preview.pdf
# Inspect working/conversion_report.json.
# Edit working/main.typ and supported assets.
typst compile working/main.typ working/preview.pdf
typst-to-docx working/main.typ --output revised.docx
docx-typst-validate source.docx working
```

1. **Convert once.** Keep the generated project in version control so edits, style changes, and asset replacements are reviewable.
2. **Read the report before editing.** Unsupported or approximated objects will not behave like ordinary text.
3. **Edit semantic constructs.** Keep `docx-para`, `docx-heading`, `docx-item`, and `docx-table` wrappers intact when changing content.
4. **Compile before export.** Typst compilation catches markup and asset-path errors early.
5. **Export and validate.** The reverse converter writes its own `conversion_report.json` next to the DOCX.
6. **Inspect pages visually.** A successful conversion is not a pixel-equivalence guarantee.

## What agents can safely change

- paragraph and heading text inside content blocks
- inline formatting using Typst markup
- list items and levels
- table cell content
- image paths and dimensions
- document metadata, page size, margins, and columns
- generated style values in `docx-styles`

## What agents should not assume

- Arbitrary Typst macros, loops, or dynamic expressions run during export.
- Word comments, tracked changes, charts, SmartArt, OLE objects, and complex equations survive the roundtrip.
- Font substitution has no effect. Missing fonts change wrapping and alignment.
- A valid PDF proves Word layout equivalence. It does not.

## Editing patterns

Change a paragraph while preserving its style:

```typst
#docx-para(styles: docx-styles, style: "Body Text")[
  Revised text with *bold* and _italic_ runs.
]
```

Add a heading:

```typst
#docx-heading(styles: docx-styles, level: 2, style: "Heading 2")[New section]
```

Add a bullet:

```typst
#docx-item(styles: docx-styles, kind: "bullet", level: 0, style: "List Bullet", marker: "•")[New item]
```

Replace an image by updating its path and dimensions:

```typst
#docx-image("assets/media/image2_ab12cd34ef.png", width: 320pt, height: 180pt, alt: "Updated chart")
```

Edit the style table to change document-wide defaults:

```typst
#let docx-styles = (
  "Normal": (font: "Aptos", size: 11pt, after: 8pt, line-spacing: 1.08),
  "Heading 1": (font: "Aptos Display", size: 18pt, bold: true, before: 18pt, after: 6pt),
)
```

## Boundaries

The IR represents the accepted document, not revision history. It preserves supported content and reports the rest. Always review `conversion_report.json` and inspect the finished pages before delivery.
