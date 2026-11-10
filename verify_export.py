from pathlib import Path
import re
import zipfile
import xml.etree.ElementTree as ET

root = Path(__file__).parent
ns = {'x': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
LONG = [20, 200, 400, 120, 60, 8, 60, 1100, 120, 50, 1000, 110, 22,
        100, 400, 65, 285, 10, 10, 70, 200, 180, 450, 45, 135, 120, 300,
        180, 860, 1300, 540, 330, 2080, 560, 930]
GASKET_SIZES = ['SIZE: DN400 3MM THK Full-Face with Bolt Holes', 'SIZE: DN200 3MM THK Full-Face with Bolt Holes',
                'SIZE: DN200 3MM THK Full-Face with Bolt Holes', 'SIZE: DN150 5MM THK Full-Face with Bolt Holes',
                'SIZE: DN100 3MM THK Full-Face with Bolt Holes']


def load(z, part):
    return ET.fromstring(z.read(part))


def text(cell, strings):
    # Shared, formula-string (t="str") and inline strings; rich text keeps a non-breaking space.
    v = cell.find('x:v', ns)
    if cell.attrib.get('t') == 's':
        value = strings[int(v.text)]
    elif v is not None:
        value = v.text
    else:
        value = ''.join(t.text or '' for t in cell.iter('{%s}t' % ns['x']))
    return value.replace('\xa0', ' ')


# Fastener orders keep quantities in F with prices in H and amounts in I; gasket orders use G, I and J.
for sub, project, material, filename, (qty_col, price_col, amount_col), expected in [
    ('ss316-test', 'C3D1', 'SS316', 'C3D1 SS316 Fastener Order.xlsx', ('F', 'H', 'I'), [2880, 8640, 5760, 17280]),
    ('hdg-test', 'C3D2', 'HDG', 'C3D2 HDG Fastener Order.xlsx', ('F', 'H', 'I'), [50, 150, 200, 50, 150, 200]),
    ('long-test', 'C3D1', 'SS316', 'C3D1 SS316 Fastener Order.xlsx', ('F', 'H', 'I'), LONG),
    ('table-test', 'C3D1', 'SS316', 'C3D1 SS316 Fastener Order.xlsx', ('F', 'H', 'I'), [630, 630, 2016, 378, 378, 160, 1016, 2032, 488, 368, 340, 410, 256, 96, 64, 64, 128, 112]),
    ('gasket-test', 'C3D1', 'EPDM', 'C3D1 EPDM Gasket Order.xlsx', ('G', 'I', 'J'), [10, 50, 5, 4, 8]),
]:
    with zipfile.ZipFile(root / 'outputs' / sub / filename) as z:
        assert z.testzip() is None
        wb = load(z, 'xl/workbook.xml')
        assert [s.attrib['name'] for s in wb.findall('x:sheets/x:sheet', ns)] == [project]
        strings = [''.join(t.text or '' for t in si.iter('{%s}t' % ns['x']))
                   for si in load(z, 'xl/sharedStrings.xml').findall('x:si', ns)] if 'xl/sharedStrings.xml' in z.namelist() else []
        sh = load(z, 'xl/worksheets/sheet1.xml')
        cells = {c.attrib['r']: c for c in sh.findall('.//x:sheetData/x:row/x:c', ns)}
        assert not any(c.attrib.get('t') == 'e' for c in cells.values())
        qty = {a: c for a, c in cells.items()
               if a.startswith(qty_col) and c.find('x:v', ns) is not None and c.attrib.get('t') != 's'}
        actual = [float(c.find('x:v', ns).text) for c in qty.values()]
        assert actual == expected, (actual, expected)
        assert cells[f'{amount_col}8'].find('x:f', ns).text == f'IF({price_col}8="","",{qty_col}8*{price_col}8)'
        assert 'B4' in cells
        # Excel recalculates every formula when the file opens, so any formula the generator did not
        # write (such as a leftover template shared formula) would silently change the order.
        rows = [int(a[1:]) for a in qty]
        last = rows[-1] + 1
        amounts = {f'{amount_col}{r}': f'IF({price_col}{r}="","",{qty_col}{r}*{price_col}{r})' for r in rows}
        totals = 0
        for a, c in cells.items():
            f = c.find('x:f', ns)
            if f is None:
                continue
            assert not f.attrib and f.text, (a, f.attrib, f.text)
            if a in amounts:
                assert f.text == amounts[a], (a, f.text)
            elif a in qty:
                assert re.fullmatch(r'SUM\(\d+(,\d+)+\)', f.text), (a, f.text)
                assert sum(map(int, re.findall(r'\d+', f.text))) == float(qty[a].find('x:v', ns).text), (a, f.text)
            else:
                assert f.text == f'IF(COUNT({price_col}8:{price_col}{last})=0,"",SUM({amount_col}8:{amount_col}{last}))', (a, f.text)
                totals += 1
        assert totals == 1 and all(a in cells and cells[a].find('x:f', ns) is not None for a in amounts)
        if material == 'EPDM':
            # Every gasket row carries the template's size wording, standard, reinforcement, PN and type.
            assert [text(cells[f'B{r}'], strings) for r in rows] == GASKET_SIZES, [text(cells[f'B{r}'], strings) for r in rows]
            assert {text(cells[f'C{r}'], strings) for r in rows} == {'BS 1154 Class Y3'}
            assert {text(cells[f'D{r}'], strings) for r in rows} == {'Single layers of fabric in accordance with BS EN 751'}
            assert [float(cells[f'E{r}'].find('x:v', ns).text) for r in rows] == [10, 10, 16, 10, 16]
            assert {text(cells[f'F{r}'], strings) for r in rows} == {'EPDM'}
            assert text(cells['B7'], strings) == 'Title: WDL12 Pump Gallery 2 and 3 PSD'
            assert 'A7' not in cells or cells['A7'].find('x:v', ns) is None
            assert text(cells['B2'], strings) == 'RFQ for Gasket EPDM'
    print(f'{project} {material} ({sub}): saved ZIP/XML, sheet name, quantities and formulas verified')

# Weekly summary list written by test.mjs: a header plus one row per item, plain values only.
summary = root / 'outputs' / 'summary-test' / 'Order Summary 2026-09-14 to 2026-09-20.xlsx'
with zipfile.ZipFile(summary) as z:
    sh = load(z, 'xl/worksheets/sheet1.xml')
rows = sh.findall('x:sheetData/x:row', ns)
assert len(rows) == 16 and not sh.findall('.//x:f', ns)
qty = [float(c.find('x:v', ns).text) for row in rows[1:] for c in row.findall('x:c', ns) if c.attrib['r'].startswith('I')]
assert qty == [2880, 8640, 5760, 17280, 50, 150, 200, 50, 150, 200, 10, 50, 5, 4, 8], qty
print('Weekly summary list: saved ZIP/XML, one row per item, quantities verified')
