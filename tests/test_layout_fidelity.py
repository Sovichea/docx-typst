from pathlib import Path
import shutil

from docx import Document
from docx.oxml.ns import qn
from PIL import Image

from docx_typst import docx_to_typst, typst_to_docx

ROOT = Path(__file__).resolve().parents[1]


def project(tmp_path: Path, body: str) -> Path:
    shutil.copyfile(ROOT / "src/docx_typst/template.typ", tmp_path / "template.typ")
    (tmp_path / "main.typ").write_text(
        '#import "template.typ": *\n'
        '#show: docx-document.with(page: (paper: "a4", margin: 65pt))\n' + body,
        encoding="utf-8")
    return tmp_path / "main.typ"


def test_native_table_inherits_stroke_and_aligns_its_outer_border(tmp_path: Path) -> None:
    source = project(tmp_path, '''
#set table(inset: 6pt, stroke: 0.5pt + rgb("DCE3DF"))
#table(columns: (100pt, 200pt), [Header], [Value], [A], [B])
''')
    typst_to_docx(source, tmp_path / "out.docx")
    table = Document(tmp_path / "out.docx").tables[0]
    properties = table._tbl.tblPr
    assert properties.find(qn("w:tblInd")).get(qn("w:w")) == "120"
    assert properties.find(qn("w:tblCellMar")).find(qn("w:left")).get(qn("w:w")) == "120"
    borders = properties.find(qn("w:tblBorders"))
    assert len(borders) == 6
    assert all(border.get(qn("w:color")) == "DCE3DF" for border in borders)
    assert all(border.get(qn("w:sz")) == "4" for border in borders)
    docx_to_typst(tmp_path / "out.docx", tmp_path / "restored", restore_source=False)
    restored = (tmp_path / "restored/main.typ").read_text(encoding="utf-8")
    assert 'stroke: 0.5pt + rgb("#DCE3DF")' in restored
    assert "inset: 6pt" in restored


def test_borderless_table_overrides_word_style_borders(tmp_path: Path) -> None:
    source = project(tmp_path, '''
#docx-table(columns: 1, style: "Table Grid", stroke: none, cells: ([No border],))
''')
    typst_to_docx(source, tmp_path / "out.docx")
    borders = Document(tmp_path / "out.docx").tables[0]._tbl.tblPr.find(qn("w:tblBorders"))
    assert len(borders) == 6
    assert all(border.get(qn("w:val")) == "nil" for border in borders)
    docx_to_typst(tmp_path / "out.docx", tmp_path / "restored")
    assert "stroke: none" in (tmp_path / "restored/main.typ").read_text(encoding="utf-8")


def test_table_rules_remain_scoped_and_alignment_survives_wrappers(tmp_path: Path) -> None:
    source = project(tmp_path, '''
#set table(stroke: 0.5pt + rgb("DCE3DF"))
#box[
  #set table(stroke: 2pt + rgb("112233"))
  #table(columns: 1, [Red])
]
#align(center)[#table(columns: (100pt,), [Gray])]
''')
    typst_to_docx(source, tmp_path / "out.docx")
    tables = Document(tmp_path / "out.docx").tables
    red = tables[0]._tbl.tblPr.find(qn("w:tblBorders"))[0]
    gray = tables[1]._tbl.tblPr.find(qn("w:tblBorders"))[0]
    assert red.get(qn("w:color")) == "112233"
    assert red.get(qn("w:sz")) == "16"
    assert gray.get(qn("w:color")) == "DCE3DF"
    assert tables[1]._tbl.tblPr.find(qn("w:jc")).get(qn("w:val")) == "center"
    assert tables[1]._tbl.tblPr.find(qn("w:tblInd")) is None


def test_image_keeps_explicit_spacing_before_it(tmp_path: Path) -> None:
    Image.new("RGB", (20, 10), "white").save(tmp_path / "picture.png")
    source = project(tmp_path, '''
#set par(spacing: 0pt)
Before
#v(12pt)
#align(right)[#image("picture.png", width: 50pt)]
After
''')
    typst_to_docx(source, tmp_path / "out.docx")
    document = Document(tmp_path / "out.docx")
    paragraph = next(p for p in document.paragraphs if p._p.xpath(".//wp:inline"))
    assert paragraph.paragraph_format.space_before.pt == 12
    assert paragraph._p.pPr.find(qn("w:jc")).get(qn("w:val")) == "right"


def test_block_spacing_adds_explicit_gaps_and_does_not_leak_from_cells(tmp_path: Path) -> None:
    Image.new("RGB", (20, 10), "white").save(tmp_path / "picture.png")
    source = project(tmp_path, '''
#set par(spacing: 20pt)
Before
#v(7pt)
#table(columns: (100pt,), [Cell
#v(60pt)])
#v(9pt)
After table
#v(5pt)
#image("picture.png", width: 50pt)
#v(3pt)
After image
''')
    typst_to_docx(source, tmp_path / "out.docx")
    document = Document(tmp_path / "out.docx")
    paragraphs = document.paragraphs
    before = next(p for p in paragraphs if p.text == "Before")
    after_table = next(p for p in paragraphs if p.text == "After table")
    picture = next(p for p in paragraphs if p._p.xpath(".//wp:inline"))
    after_image = next(p for p in paragraphs if p.text == "After image")
    assert before.paragraph_format.space_after.pt == 27
    assert after_table.paragraph_format.space_before.pt == 29
    assert after_table.paragraph_format.space_after.pt == 0
    assert picture.paragraph_format.space_before.pt == 25
    assert after_image.paragraph_format.space_before.pt == 23
    assert document.tables[0].cell(0, 0).paragraphs[-1].paragraph_format.line_spacing.pt == 60


def test_adjacent_table_gaps_collapse_and_compact_wrapper_keeps_template_inset(tmp_path: Path) -> None:
    source = project(tmp_path, '''
#set par(spacing: 20pt)
#set table(inset: 6pt, stroke: 0.5pt + rgb("DCE3DF"))
#compact[#table(columns: (100pt,), [First])]
#table(columns: (100pt,), [Second])
''')
    # Use the template's usual newline-separated set-rule form.
    (tmp_path / "layout.typ").write_text(
        '#let compact(body) = {\n set table(inset: 4pt)\n body\n}', encoding="utf-8")
    source.write_text('#import "layout.typ": compact\n' + source.read_text(encoding="utf-8"), encoding="utf-8")
    typst_to_docx(source, tmp_path / "out.docx")
    document = Document(tmp_path / "out.docx")
    first, second = document.tables
    assert first._tbl.tblPr.find(qn("w:tblInd")).get(qn("w:w")) == "80"
    assert second._tbl.tblPr.find(qn("w:tblInd")).get(qn("w:w")) == "120"
    spacer = document.paragraphs[0]
    assert spacer.paragraph_format.line_spacing.pt == 20


def test_table_show_rule_disables_inherited_justification_and_keeps_explicit_alignment(tmp_path: Path) -> None:
    source = project(tmp_path, '''
#show table: set par(justify: false)
#table(columns: (100pt, 100pt), [Left cell], [#align(right)[Right cell]])
Body paragraph
''')
    source.write_text(source.read_text(encoding="utf-8").replace(
        '#show: docx-document',
        '#let docx-styles = ("Normal": (justify: true), "Table Text": (based-on: "Normal"))\n#show: docx-document'),
        encoding="utf-8")
    typst_to_docx(source, tmp_path / "out.docx")
    document = Document(tmp_path / "out.docx")
    left, right = document.tables[0].rows[0].cells
    assert left.paragraphs[0]._p.pPr.find(qn("w:jc")).get(qn("w:val")) == "left"
    assert right.paragraphs[0]._p.pPr.find(qn("w:jc")).get(qn("w:val")) == "right"
    assert document.styles["Normal"].paragraph_format.alignment == 3


def test_apostrophes_do_not_leak_code_or_comments_into_word(tmp_path: Path) -> None:
    source = project(tmp_path, '''
The hospital's original paragraph.
// Hidden conversion note
#table(columns: (100pt,), [The patient's record])
#text(size: 10pt)[The pharmacist's recovery decision.]
#image("picture.png", width: 30pt)
''')
    Image.new("RGB", (20, 10), "white").save(tmp_path / "picture.png")
    typst_to_docx(source, tmp_path / "out.docx")
    document = Document(tmp_path / "out.docx")
    assert document.tables[0].cell(0, 0).text == "The patient's record"
    text = "\n".join(p.text for p in document.paragraphs)
    assert "The pharmacist's recovery decision." in text
    assert "Hidden conversion note" not in text
    assert "#table" not in text and "#text" not in text
    assert len(document.inline_shapes) == 1


def test_native_table_padding_accounts_for_line_boxes_without_changing_typography(tmp_path: Path) -> None:
    source = project(tmp_path, '''
#set table(inset: 4pt)
#table(columns: (120pt,), [Native cell])
''')
    source.write_text(source.read_text(encoding="utf-8").replace(
        '#show: docx-document',
        '#let docx-styles = ("Table Text": (font: "Arial", size: 10pt, line-spacing-pt: 14pt))\n#show: docx-document'),
        encoding="utf-8")
    typst_to_docx(source, tmp_path / "out.docx")
    document = Document(tmp_path / "out.docx")
    table = document.tables[0]
    logical = table._tbl.tblPr.find(qn("w:tblCellMar"))
    assert logical.find(qn("w:top")).get(qn("w:w")) == "80"
    adjusted = table.cell(0, 0)._tc.tcPr.find(qn("w:tcMar"))
    assert 0 <= int(adjusted.find(qn("w:top")).get(qn("w:w"))) < 80
    assert 0 <= int(adjusted.find(qn("w:bottom")).get(qn("w:w"))) < 80
    assert adjusted.find(qn("w:left")) is None
    assert document.styles["Table Text"].font.size.pt == 10
    assert document.styles["Table Text"].paragraph_format.line_spacing.pt == 14
    docx_to_typst(tmp_path / "out.docx", tmp_path / "restored")
    assert "inset: 4pt" in (tmp_path / "restored/main.typ").read_text(encoding="utf-8")


def test_grid_gutters_are_real_empty_columns(tmp_path: Path) -> None:
    source = project(tmp_path, '#grid(columns: (1fr, 1fr), column-gutter: 24pt, [Left], [Right])')
    typst_to_docx(source, tmp_path / "out.docx")
    table = Document(tmp_path / "out.docx").tables[0]
    assert len(table.columns) == 3
    assert table.columns[1].width.pt == 24
    assert [cell.text for cell in table.rows[0].cells] == ["Left", "", "Right"]
    assert abs(table.columns[0].width.pt - table.columns[2].width.pt) < 0.1


def test_independent_enums_restart_without_breaking_contiguous_items(tmp_path: Path) -> None:
    source = project(tmp_path, '+ First\n+ Second\n\nA new procedure\n\n+ Third\n+ Fourth')
    typst_to_docx(source, tmp_path / "out.docx")
    document = Document(tmp_path / "out.docx")
    items = [p for p in document.paragraphs if p.style.name == "List Number"]
    ids = [p._p.pPr.numPr.numId.val for p in items]
    assert ids[0] == ids[1] and ids[2] == ids[3] and ids[0] != ids[2]
    for identity in set(ids):
        instance = document.part.numbering_part.element.xpath(f'./w:num[@w:numId="{identity}"]')[0]
        assert instance.xpath(".//w:startOverride")[0].get(qn("w:val")) == "1"


def test_explicit_rules_use_safe_paragraph_borders(tmp_path: Path) -> None:
    source = project(tmp_path, '''
#show heading.where(level: 1): it => {
  block[#it.body #rect(width: 38pt, height: 3pt, fill: rgb("F7EE16"))]
}
= Heading
#docx-rule(width: 38pt, height: 3pt, fill: rgb("F7EE16"), before: 14.8pt, word-before: 8.7pt, after: 20pt, word-after: 15pt)
Body
''')
    report = typst_to_docx(source, tmp_path / "out.docx")
    document = Document(tmp_path / "out.docx")
    assert report["counts"].get("images", 0) == 0
    assert not document.element.body.xpath('.//w:drawing')
    borders = document.element.body.xpath('.//w:pBdr/w:bottom')
    assert len(borders) == 1  # Complex heading show rules are not interpreted.
    assert borders[0].get(qn("w:sz")) == "24"
    assert borders[0].get(qn("w:color")) == "F7EE16"
    spacing = borders[0].getparent().getparent().find(qn("w:spacing"))
    assert spacing.get(qn("w:before")) == "174"
    assert spacing.get(qn("w:after")) == "300"
    docx_to_typst(tmp_path / "out.docx", tmp_path / "restored")
    assert '#docx-rule(width: 38pt, height: 3pt' in (tmp_path / "restored/main.typ").read_text(encoding="utf-8")


def test_template_text_helpers_preserve_price_size_and_weight(tmp_path: Path) -> None:
    (tmp_path / "layout.typ").write_text(
        '#let price(body) = text(font: "League Spartan", size: 25pt, weight: "semibold", body)', encoding="utf-8")
    source = project(tmp_path, '#price[USD 12,000]')
    source.write_text('#import "layout.typ": price\n' + source.read_text(encoding="utf-8"), encoding="utf-8")
    typst_to_docx(source, tmp_path / "out.docx")
    run = next(r for p in Document(tmp_path / "out.docx").paragraphs for r in p.runs if "12,000" in r.text)
    assert run.font.size.pt == 25
    assert run.font.name == "League Spartan SemiBold"
    assert run.bold is False


def test_total_field_targets_last_logical_page_and_roundtrips(tmp_path: Path) -> None:
    source = project(tmp_path, '#docx-page-count-padded()\n\nBody')
    typst_to_docx(source, tmp_path / "out.docx")
    document = Document(tmp_path / "out.docx")
    field = document.element.body.xpath(".//w:fldSimple")[0]
    assert 'PAGEREF _dxt_document_end' in field.get(qn("w:instr"))
    assert '"00"' in field.get(qn("w:instr"))
    assert document.element.body.xpath('.//w:bookmarkStart[@w:name="_dxt_document_end"]')
    docx_to_typst(tmp_path / "out.docx", tmp_path / "restored")
    assert '#docx-page-count-padded()' in (tmp_path / "restored/main.typ").read_text(encoding="utf-8")


def test_panel_accent_keeps_compound_color_expression(tmp_path: Path) -> None:
    source = project(tmp_path, '#docx-panel(stroke: (left: 3pt + rgb("F7EE16")))[Panel]')
    typst_to_docx(source, tmp_path / "out.docx")
    borders = Document(tmp_path / "out.docx").tables[0]._tbl.tblPr.find(qn("w:tblBorders"))
    assert borders.find(qn("w:left")).get(qn("w:color")) == "F7EE16"
    assert borders.find(qn("w:left")).get(qn("w:sz")) == "24"
    assert borders.find(qn("w:right")).get(qn("w:val")) == "nil"


def test_table_default_alignment_preserves_explicit_paragraph_style(tmp_path: Path) -> None:
    source = project(tmp_path, '''
#let docx-styles = ("Right Label": (align: "right", size: 8pt))
#show: docx-document.with(styles: docx-styles)
#show table: set par(justify: false)
#table(columns: (100pt, 100pt),
  [Left], [#docx-para(style: "Right Label")[Right]])
''')
    typst_to_docx(source, tmp_path / "out.docx")
    table = Document(tmp_path / "out.docx").tables[0]
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    assert table.cell(0, 1).paragraphs[0].alignment == WD_ALIGN_PARAGRAPH.RIGHT


def test_inline_artwork_is_not_clipped_by_exact_line_height(tmp_path: Path) -> None:
    Image.new("RGB", (28, 28), "yellow").save(tmp_path / "logo.png")
    source = project(tmp_path, '''
#let docx-styles = ("Logo": (line-spacing-pt: 12pt))
#docx-para(style: "Logo")[#docx-inline-image("logo.png", width: 28pt)]
''')
    typst_to_docx(source, tmp_path / "out.docx")
    document = Document(tmp_path / "out.docx")
    from docx.enum.text import WD_LINE_SPACING
    paragraph = document.paragraphs[0]
    assert paragraph.paragraph_format.line_spacing_rule != WD_LINE_SPACING.EXACTLY
    assert document.inline_shapes[0].width.pt == 28
    assert document.inline_shapes[0].height.pt == 28


def test_header_and_footer_rules_use_the_correct_paragraph_edges(tmp_path: Path) -> None:
    source = project(tmp_path, "Body")
    source.write_text(source.read_text(encoding="utf-8").replace(
        'margin: 65pt))',
        'margin: 65pt), header: [#line(length: 100%)], footer: [#line(length: 100%)])'), encoding="utf-8")
    typst_to_docx(source, tmp_path / "out.docx")
    section = Document(tmp_path / "out.docx").sections[0]
    assert section.header._element.xpath(".//w:pBdr/w:bottom")
    assert section.footer._element.xpath(".//w:pBdr/w:top")
    assert not section.footer._element.xpath(".//w:pBdr/w:bottom")


def test_single_contact_table_preserves_fill_padding_and_merged_width(tmp_path: Path) -> None:
    source = project(tmp_path, '''
#docx-table(columns: 2, rows: 3, widths: (117pt, 1fr),
  merges: ((0, 0, 1, 2),), stroke: none, inset: 0pt,
  fills: ((0, 0, "#FFF9E7"), (1, 0, "#FFF9E7"), (1, 1, "#FFF9E7"), (2, 0, "#FFF9E7"), (2, 1, "#FFF9E7")),
  cell-insets: (
    (0, 0, (left: 18pt, right: 18pt, top: 18pt, bottom: 18pt)),
    (1, 0, (left: 18pt, right: 14pt, top: 6pt, bottom: 6pt)),
    (1, 1, (left: 0pt, right: 18pt, top: 6pt, bottom: 6pt))),
  cells: ([Company], [Phone], [Number], [Registration], [1000061103]))
''')
    typst_to_docx(source, tmp_path / "out.docx")
    document = Document(tmp_path / "out.docx")
    assert len(document.tables) == 1
    table = document.tables[0]
    assert not table._tbl.xpath(".//w:tbl")
    assert len(table.rows) == 3
    assert abs(table.cell(0, 0).width.pt - sum(c.width.pt for c in table.columns)) < 0.1
    assert table.cell(1, 0).width.pt == 117
    assert table._tbl.tblPr.find(qn("w:tblInd")).get(qn("w:w")) == "360"
    assert table.cell(1, 0)._tc.tcPr.find(qn("w:tcMar")).find(qn("w:left")).get(qn("w:w")) == "360"
    for row in table.rows:
        for cell in row.cells:
            assert cell._tc.tcPr.find(qn("w:shd")).get(qn("w:fill")) == "FFF9E7"


def test_simple_imported_panel_resolves_template_palette(tmp_path: Path) -> None:
    (tmp_path / "layout.typ").write_text(
        '#let palette = (cream: rgb("FFF9E7"))\n'
        '#let panel(body) = block(fill: palette.cream, inset: 18pt, body)', encoding="utf-8")
    source = project(tmp_path, '#panel[Contact]')
    source.write_text('#import "layout.typ": panel\n' + source.read_text(encoding="utf-8"), encoding="utf-8")
    typst_to_docx(source, tmp_path / "out.docx")
    table = Document(tmp_path / "out.docx").tables[0]
    assert table.cell(0, 0)._tc.tcPr.find(qn("w:shd")).get(qn("w:fill")) == "FFF9E7"
    assert table._tbl.tblPr.find(qn("w:tblCellMar")).find(qn("w:left")).get(qn("w:w")) == "360"


def test_imported_text_helper_uses_existing_word_paragraph_style(tmp_path: Path) -> None:
    (tmp_path / "layout.typ").write_text(
        '#let contact-heading(body) = text(font: "Arial", size: 28pt, body)', encoding="utf-8")
    source = project(tmp_path, '#contact-heading[Company]')
    source.write_text('#import "layout.typ": contact-heading\n'
        '#let docx-styles = ("Contact Heading": (font: "Arial", size: 28pt, '
        'word-after: 36.2pt, line-spacing-pt: 30.48pt))\n'
        + source.read_text(encoding="utf-8"), encoding="utf-8")
    typst_to_docx(source, tmp_path / "out.docx")
    paragraph = Document(tmp_path / "out.docx").paragraphs[0]
    assert paragraph.style.name == "Contact Heading"
    assert abs(paragraph.paragraph_format.line_spacing.pt - 30.48) < 0.06
    assert abs(paragraph.style.paragraph_format.space_after.pt - 36.2) < 0.01


def test_word_specific_space_does_not_change_typst_spacing(tmp_path: Path) -> None:
    source = project(tmp_path, 'Before\n#docx-space(24pt, word: 6pt)\nAfter')
    typst_to_docx(source, tmp_path / "out.docx")
    after = next(p for p in Document(tmp_path / "out.docx").paragraphs if p.text == "After")
    assert after.paragraph_format.space_before.pt == 6
    assert '#docx-space(24pt, word: 6pt)' in source.read_text(encoding="utf-8")
