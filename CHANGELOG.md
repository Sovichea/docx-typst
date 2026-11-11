# Changelog

## 0.1.0

Initial alpha release for the documented Word/Typst subset.

- Editable text, styles, headings, lists, tables, links, images, page setup,
  sections, headers, footers, and page-number fields.
- Native Typst artwork exported as SVG with an Office bitmap fallback;
  supported positioned covers retain editable text.
- Explicit `docx-eval` evaluation for calculated text and generated tables.
  Formulas remain in the authored source; Word receives static values.
- Embedded source and local dependencies restore exact native source from
  unchanged exports, including supported Microsoft Word save normalization.
  Word edits trigger reconstruction instead of restoring stale source.
- Deduplicated source assets, per-cell borders and insets, currency alignment,
  table justification rules, and field formatting retained across Word updates.
- Starter project, conversion reports, validation CLI, and agent skills for
  adapting existing documents or authoring compatible documents.

Word and Typst use different layout engines. Review reports and rendered pages;
unsupported features and arbitrary unannotated Typst programs are not faithfully
reconstructed. Microsoft Word is needed for Word-specific visual verification,
not for running the converter.
