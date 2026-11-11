from pathlib import Path
import shutil

from docx import Document
from docx.oxml.ns import qn
import pytest
import typst

from docx_typst import typst_to_docx, docx_to_typst


def test_computations_export_values_and_restore_formulas(tmp_path: Path):
    shutil.copyfile(Path(__file__).parents[1] / "src/docx_typst/template.typ", tmp_path / "template.typ")
    library = tmp_path / "prices.typ"
    library.write_text('#let total(items, vat) = items.sum() * (1 + vat)')
    source = tmp_path / "quotation.typ"
    formula = '''#import "template.typ": *
#import "prices.typ": total
#let items = (
  100,
  200,
)
#let vat = 0.1
#let amount = calc.round(total(items, vat), digits: 2)
#docx-eval("amount", [*Total #amount*])
#docx-eval("rows", table(columns: (80pt, 1fr),
  fill: (x, y) => if x == 0 { luma(240) },
  stroke: (x, y) => (bottom: 0.5pt + luma(200)),
  align: (left, right), [VAT inclusive], [#amount]))
'''
    source.write_text(formula)
    for number, expected in ((200, '330'), (300, '440')):
        authored = formula.replace('  200,', f'  {number},')
        source.write_text(authored)
        output = tmp_path / f'{number}.docx'
        report = typst_to_docx(source, output)
        doc = Document(output)
        assert doc.paragraphs[0].text == f'Total {expected}'
        assert doc.tables[0].cell(0, 1).text == expected
        assert doc.tables[0].cell(0, 0)._tc.find('.//' + qn('w:shd')).get(qn('w:fill')) == 'F0F0F0'
        edge = doc.tables[0].cell(0, 1)._tc.find('.//' + qn('w:tcBorders') + '/' + qn('w:bottom'))
        assert edge.get(qn('w:color')) == 'C8C8C8'
        assert report['computed_values'] == ['amount', 'rows']
        assert source.read_text() == authored
        restored = tmp_path / f'restored-{number}'
        assert docx_to_typst(output, restored)['source_restoration'] == 'exact'
        assert (restored / 'main.typ').read_text() == authored
        assert (restored / 'prices.typ').read_bytes() == library.read_bytes()
        typst.compile(str(restored / 'main.typ'))


def test_opaque_computed_formatting_fails_explicitly(tmp_path: Path):
    shutil.copyfile(Path(__file__).parents[1] / "src/docx_typst/template.typ", tmp_path / "template.typ")
    source = tmp_path / 'unsupported.typ'
    source.write_text('#import "template.typ": *\n#docx-eval("styled", text(fill: red)[Value])')
    with pytest.raises(ValueError, match='unsupported computed content: styled'):
        typst_to_docx(source, tmp_path / 'out.docx')

