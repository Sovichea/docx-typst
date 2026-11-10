# docx-typst

`docx-typst` is a bidirectional document conversion tool. It converts Word (`.docx`) documents into editable Typst projects and exports a documented, statically resolvable Typst subset back to editable Word documents.

It is designed for **file-based document workflows**, including AI-agent and version-control workflows. A document can be read and revised as readable Typst source, compiled to PDF for review, and exported to DOCX without a live Word session. Supported text, styles, lists, tables, images, headers, footers, and page setup remain editable. Conversion reports identify content that is approximated, flattened, or unsupported.

Version `0.1.0` is an alpha release for documents within the documented subset. It does not claim to reproduce every Word or Typst feature. Review `conversion_report.json` and inspect important pages before delivery.

- [Supported subset](docs/supported-subset.md)
- [Agentic document workflow](docs/agentic-document-workflow.md)

## Install

Python 3.10 through 3.13 is supported.

```console
python -m pip install docx-typst
```

The Python package performs conversion. Install the [Typst CLI](https://typst.app/) separately to compile generated projects to PDF. SVG-to-DOCX export needs Cairo native libraries through `cairosvg`; install the optional extra when needed:

```console
python -m pip install "docx-typst[svg]"
```

EMF/WMF rendering uses Inkscape when available, or the packaged Windows PowerShell/GDI+ path on Windows. Without a renderer the original asset is preserved and the conversion report records a warning.

## Word to Typst

```console
docx-to-typst report.docx --output report-typst
typst compile report-typst/main.typ report.pdf
```

The generated project contains:

```text
report-typst/
  main.typ
  template.typ
  asset_manifest.json
  conversion_report.json
  assets/media/
```

`main.typ` uses semantic constructs such as `docx-heading`, `docx-para`, `docx-item`, `docx-image`, and `docx-table`. Paragraph style names and effective formatting are stored in a generated `docx-styles` table, so the source stays readable while the reverse converter can recreate Word styles.

Useful options:

```console
docx-to-typst report.docx --output report-typst --no-media
docx-to-typst report.docx --output report-typst --drop-empty
docx-to-typst report.docx --output report-typst --strict
```

## Typst to Word

```console
typst-to-docx report-typst/main.typ --output report-roundtrip.docx
```

The reverse converter reads the documented static subset. It does not evaluate arbitrary Typst programs, dynamic expressions, loops, or user-defined macros. It writes `conversion_report.json` next to the output. Pass `--reference template.docx` to reuse an existing document's theme and styles, and `--strict` to fail when unsupported content is found.

A typical edit-and-export loop:

```console
docx-to-typst source.docx --output working
# Edit working/main.typ or replace assets.
typst compile working/main.typ working/preview.pdf
typst-to-docx working/main.typ --output revised.docx
```

## Start a Typst document

```console
docx-typst init my-document
typst compile my-document/main.typ my-document/preview.pdf
typst-to-docx my-document/main.typ --output my-document.docx
```

The starter includes a helper module, an authoring example, and an SVG asset. To write only the helper module:

```console
docx-typst module --output docx-typst.typ
```

Shortcut commands are also available:

```console
docx-typst to-typst document.docx --output document-typst
docx-typst to-docx document-typst/main.typ --output document.docx
```

## Supported subset

| Word feature | Word to Typst | Typst to Word |
|---|---|---|
| Headings and paragraph styles | Semantic headings plus a generated style table | Built-in and custom paragraph styles |
| Character formatting | Bold, italic, underline, strike, caps, super/subscript, color, highlight, font, size | The same supported properties |
| Lists | Bullet and numbered levels with resolved markers | List Bullet and List Number styles |
| Tables | Grid widths, header rows, horizontal/vertical merges, cell fills, common borders | Native tables, merges, header rows, cell fills |
| Images | PNG/JPEG/SVG retained; legacy formats converted or reported | PNG/JPEG embedded; SVG rasterized when Cairo is available |
| Links and bookmarks | External hyperlinks and named bookmarks | External hyperlinks and internal anchors |
| Headers and footers | Default section header/footer and page fields | Section headers, footers, page-number fields |
| Page setup | Paper size, margins, columns, orientation, sections | Paper size, margins, columns, orientation, sections |
| Footnotes | Rendered as Typst footnotes with a warning | Exported as superscript note text |
| Comments, charts, SmartArt, OLE, equations | Reported as unsupported or approximated | Not reconstructed |

The reverse converter also accepts a small native Typst subset, including headings, paragraphs, bullet and numbered lists, `*strong*`, `_emphasis_`, inline code, `#link`, `#image`, `#figure`, and `#table`.

## Reports and validation

Both directions write `conversion_report.json` with counts, status, and diagnostics. A successful command does not mean every source object was reproduced.

Compare a source DOCX with a generated Typst project:

```console
docx-typst-validate source.docx project --out project/validation.json
```

The validation report compares plain text, paragraph and table counts, images, styles, page size, and margins. Differences are informational; inspect the flagged content before delivery.

## Python API

```python
from docx_typst import docx_to_typst, typst_to_docx

forward = docx_to_typst("report.docx", "report-typst")
backward = typst_to_docx("report-typst/main.typ", "report-roundtrip.docx")
```

## Fidelity notes

- Typst normalizes whitespace, smart punctuation, and some spacing. Word's layout engine and Typst's layout engine will not produce identical line breaks.
- Fonts are referenced by name but are not embedded or installed. Missing fonts change wrapping and alignment.
- Floating shapes, text boxes, comments, tracked-change metadata, native chart data, SmartArt, OLE objects, and complex equations are outside the supported subset.
- The reverse converter does not evaluate arbitrary Typst code. Unsupported constructs are reported and, when possible, flattened into readable content.
- Review `conversion_report.json` and inspect important pages before using the output for delivery.

## Development

```console
python -m pip install -e ".[test]"
python -m pytest -q
python -m pip install --upgrade build twine
python -m build
python -m twine check --strict dist/*
```

Open an [issue](https://github.com/Sovichea/docx-typst/issues) for questions or bug reports.

## License

MIT. See [LICENSE](LICENSE).
