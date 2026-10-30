import sys
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

NS = {'m': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

def col_letter(idx):
    s = ''
    while idx > 0:
        idx, r = divmod(idx - 1, 26)
        s = chr(65 + r) + s
    return s

def load_shared_strings(z):
    try:
        root = ET.fromstring(z.read('xl/sharedStrings.xml'))
    except KeyError:
        return []
    out = []
    for si in root.findall('m:si', NS):
        texts = [t.text or '' for t in si.findall('.//m:t', NS)]
        out.append(''.join(texts))
    return out

def cell_value(c, shared):
    t = c.get('t')
    v = c.find('m:v', NS)
    if v is None or v.text is None:
        return ''
    if t == 's':
        return shared[int(v.text)]
    return v.text

def inspect_sheet(z, sheet_path, shared, rows=range(1, 16)):
    root = ET.fromstring(z.read(sheet_path))
    cells = {}
    for row in root.findall('.//m:sheetData/m:row', NS):
        r = int(row.get('r'))
        if r not in rows:
            continue
        for c in row.findall('m:c', NS):
            ref = c.get('r')
            col = ''.join(ch for ch in ref if ch.isalpha())
            cells[(col, r)] = cell_value(c, shared)
    print('---', sheet_path, '---')
    for col_idx in range(3, 30):  # C..AC
        col = col_letter(col_idx)
        vals = [cells.get((col, r), '') for r in rows]
        if any(vals):
            print(f'{col}: ' + ' | '.join(f'R{r}={vals[i]!r}' for i, r in enumerate(rows)))

def main():
    path = Path(sys.argv[1] if len(sys.argv) > 1 else r'c:\Users\robgo\EXCEL\fluxo_caixa_crud\bin\Release\Fluxo_Caixa.xlsx')
    if not path.exists():
        print('NOT FOUND', path)
        return 1
    with zipfile.ZipFile(path) as z:
        shared = load_shared_strings(z)
        wb = ET.fromstring(z.read('xl/workbook.xml'))
        rels = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
        relmap = {r.get('Id'): r.get('Target') for r in rels}
        for sh in wb.findall('m:sheets/m:sheet', NS):
            name = sh.get('name')
            if name in ('ALS', '_CONFIG'):
                rid = sh.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                target = relmap[rid]
                if not target.startswith('xl/'):
                    target = 'xl/' + target.lstrip('/')
                print('SHEET', name)
                inspect_sheet(z, target, shared)
    return 0

if __name__ == '__main__':
    raise SystemExit(main())
