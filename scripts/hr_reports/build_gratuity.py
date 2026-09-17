import openpyxl, re, subprocess, sys, os
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from collections import defaultdict

FILE = 'Comprehensive_Salons_and_Staff_Register_EN_with_Payroll_12_REVIEWED.xlsx'
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

def recalc(path):
    result = subprocess.run([sys.executable, os.path.join(SCRIPT_DIR, 'recalc.py'), path, '270'],
                             capture_output=True, text=True)
    print('recalc:', result.stdout.strip() or result.stderr.strip())

recalc(FILE)

wbv = openpyxl.load_workbook(FILE, data_only=True)
salons = ['ALEKSANDRA', 'UNIQUE YOU', 'NISANTASI', 'THE LAB - Branch', 'THE LAB - HQ', 'HAIR TAG']

def norm(s):
    s = re.sub(r'[^A-Z ]', '', str(s).upper())
    return re.sub(r'\s+', ' ', s).strip()

def clean(raw):
    return re.sub(r'^🔴\s*(ABSCONDED|TERMINATED)\s*[—-]\s*', '', str(raw).strip(), flags=re.IGNORECASE).strip()

records = []
for s in salons:
    ws = wbv[s]
    for r in range(5, ws.max_row + 1):
        name = ws.cell(row=r, column=3).value
        if not name:
            continue
        cname = clean(name)
        records.append({
            'salon': s, 'row': r, 'raw_name': name, 'name': cname, 'norm': norm(cname),
            'passport': ws.cell(row=r, column=2).value,
            'profession': ws.cell(row=r, column=4).value,
            'csd': ws.cell(row=r, column=33).value,
            'basic': ws.cell(row=r, column=15).value,
            'commission': ws.cell(row=r, column=32).value,
            'notes': ws.cell(row=r, column=30).value or '',
            'absconded': 'ABSCONDED' in str(name).upper(),
            'terminated': 'TERMINATED' in str(name).upper(),
        })

# dedupe by passport (fallback to salon+name) to get one entry per real person
by_person = defaultdict(list)
for rec in records:
    key = rec['passport'] or (rec['salon'], rec['norm'])
    by_person[key].append(rec)

already_in_sheet = {'SHAHNAZA MATYAKUBOVA', 'ZOYA MAULENOVA'}  # existing rows 4 & 5

# people with real Join Date + Basic Salary on at least one of their records, not already in the sheet
active_with_data = []
for key, recs in by_person.items():
    if recs[0]['name'] in already_in_sheet or recs[0]['absconded'] or recs[0]['terminated']:
        continue
    dated = [r for r in recs if r['csd'] and r['basic']]
    if dated:
        # prefer the highest basic salary on file - avoids picking a nominal
        # AED-100 commission-registration placeholder over a real contracted wage
        dated.sort(key=lambda r: -r['basic'])
        active_with_data.append((recs, dated[0]))

print(f"{len(active_with_data)} active employees have a real Join Date + Basic Salary on file.")
for recs, d in active_with_data:
    other_salons = [r['salon'] for r in recs if r is not d]
    print(' -', d['name'], '|', d['salon'], '| CSD:', d['csd'], '| Basic:', d['basic'],
          ('| also at: ' + ', '.join(other_salons) if other_salons else ''))

# ---------------- write to workbook ----------------
wb = openpyxl.load_workbook(FILE, data_only=False)
ws = wb['Gratuity Calculator']

FONT_NAME = 'Arial'
normal_font = Font(name=FONT_NAME, size=10)
note_font = Font(name=FONT_NAME, size=9, italic=True, color='9C5700')

row = 6
for recs, d in active_with_data:
    other = [r['salon'] for r in recs if r is not d]
    notes = []
    if len(recs) > 1:
        notes.append(f"Also holds work permit(s) at: {', '.join(other)} - flagged as a duplicate/multi-salon "
                      f"permit (see 'Duplicate Employees' tab). Confirm which salary should be used for gratuity "
                      f"before finalizing; do not double-count service.")
    if (d['commission'] and str(d['commission']).upper().startswith('Y')) or 'commission' in d['notes'].lower():
        notes.append("Commission-based - the Basic Salary on file may be a nominal MOHRE/WPS registration figure, "
                      "not real take-home pay. Confirm the actual average monthly salary (e.g. from WPS transfers) "
                      "before finalizing this gratuity figure.")
    note = ' '.join(notes)
    ws.cell(row=row, column=1, value=d['name'])
    ws.cell(row=row, column=2, value=d['salon'] + (' (+ others)' if other else ''))
    ws.cell(row=row, column=3, value=d['profession'])
    ws.cell(row=row, column=4, value=d['csd'])
    ws.cell(row=row, column=5, value='=TODAY()')
    ws.cell(row=row, column=6, value=f'=IF(OR(D{row}="",E{row}=""),"",(E{row}-D{row})/365)')
    ws.cell(row=row, column=7, value=d['basic'])
    ws.cell(row=row, column=8, value=f'=IF(G{row}="","",G{row}*12/365)')
    ws.cell(row=row, column=9, value='Still employed - accrued to date (not a final payout)' + (' ' + note if note else ''))
    ws.cell(row=row, column=10, value=f'=IF(OR(F{row}="",F{row}=0),"",IF(F{row}<1,0,IF(F{row}<=5,F{row}*21,5*21+(F{row}-5)*30)))')
    ws.cell(row=row, column=11, value=f'=IF(OR(J{row}="",H{row}=""),"",J{row}*H{row})')
    ws.cell(row=row, column=12, value=f'=IF(G{row}="","",G{row}*24)')
    ws.cell(row=row, column=13, value=f'=IF(OR(K{row}="",L{row}=""),"",MIN(K{row},L{row}))')
    ws.cell(row=row, column=14, value=0)
    ws.cell(row=row, column=15, value=f'=IF(M{row}="","",M{row}-N{row})')
    for c in range(1, 16):
        cell = ws.cell(row=row, column=c)
        cell.font = normal_font
        if c in (4, 5):
            cell.number_format = 'dd/mm/yyyy'
    row += 1

last_data_row = row - 1
print('Gratuity Calculator: wrote rows 6 to', last_data_row)

# ---------------- new sheet: full staff gratuity status ----------------
if 'All Staff - Gratuity Status' in wb.sheetnames:
    del wb['All Staff - Gratuity Status']
ws2 = wb.create_sheet('All Staff - Gratuity Status')
ws2.sheet_view.showGridLines = False

header_font = Font(name=FONT_NAME, bold=True, color='FFFFFF', size=11)
header_fill = PatternFill('solid', fgColor='305496')
title_font = Font(name=FONT_NAME, bold=True, size=14, color='305496')
sub_font = Font(name=FONT_NAME, italic=True, size=10, color='595959')
thin = Side(style='thin', color='BFBFBF')
border = Border(left=thin, right=thin, top=thin, bottom=thin)

ws2['A1'] = 'All Staff - Gratuity / Join Date Status'
ws2['A1'].font = title_font
ws2['A2'] = ('One row per employee. "Calculated" = a real Join Date and Basic Salary are on file and their '
             'gratuity accrual is in the Gratuity Calculator tab. "Awaiting Contract" = no Join Date on file yet - '
             'upload their signed labour contract (or the physical work permit card, which sometimes states the '
             'issue date) and I will fill it in. Anyone with more than one work permit gets a row per salon, so a '
             'second real contract shows as a second Join Date once you provide it.')
ws2['A2'].font = sub_font
ws2.merge_cells('A1:H1')
ws2.merge_cells('A2:H2')
ws2.row_dimensions[2].height = 44
ws2['A2'].alignment = Alignment(wrap_text=True, vertical='top')

headers = ['Salon', 'Name', 'Profession', 'Join Date', 'Basic Salary (AED)', 'Status', 'Duplicate/Multi-Permit?', 'Notes']
hr = 4
for i, h in enumerate(headers, start=1):
    ws2.cell(row=hr, column=i, value=h)
for c in range(1, len(headers) + 1):
    cell = ws2.cell(row=hr, column=c)
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = Alignment(vertical='center', wrap_text=True)
    cell.border = border

calculated_names = {d['name'] for _, d in active_with_data} | already_in_sheet
STATUS_COLORS = {'Calculated': 'C6EFCE', 'Awaiting Contract': 'FFF2CC', 'Absconded': 'C00000', 'Terminated': '808080'}

r = hr + 1
for rec in sorted(records, key=lambda x: (x['salon'], x['name'])):
    dup = len(by_person[rec['passport'] or (rec['salon'], rec['norm'])]) > 1
    if rec['absconded']:
        status = 'Absconded'
        note = 'Absconded - gratuity may be forfeited under Article 42 of UAE labour law; get legal advice before any payout.'
    elif rec['terminated']:
        status = 'Terminated'
        note = 'See Gratuity Calculator tab - already calculated as part of final settlement.'
    elif rec['name'] in calculated_names and rec['csd']:
        status = 'Calculated'
        note = 'In Gratuity Calculator tab (accrued to date).'
    else:
        status = 'Awaiting Contract'
        note = 'No Join Date on file - upload contract/work permit card to complete.'
    vals = [rec['salon'], rec['name'], rec['profession'], rec['csd'], rec['basic'], status,
            'YES' if dup else '', note]
    for c, v in enumerate(vals, start=1):
        cell = ws2.cell(row=r, column=c, value=v)
        cell.font = normal_font
        cell.border = border
        cell.alignment = Alignment(vertical='top', wrap_text=True)
        if c == 4 and v:
            cell.number_format = 'dd/mm/yyyy'
    status_cell = ws2.cell(row=r, column=6)
    status_cell.fill = PatternFill('solid', fgColor=STATUS_COLORS[status])
    status_cell.font = Font(name=FONT_NAME, size=10, bold=True,
                             color='FFFFFF' if status in ('Absconded', 'Terminated') else '000000')
    if dup:
        ws2.cell(row=r, column=7).fill = PatternFill('solid', fgColor='FFC7CE')
        ws2.cell(row=r, column=7).font = Font(name=FONT_NAME, size=10, bold=True)
    r += 1

widths = [16, 26, 16, 13, 16, 18, 18, 46]
for i, w in enumerate(widths, start=1):
    ws2.column_dimensions[get_column_letter(i)].width = w
ws2.freeze_panes = f'A{hr + 1}'
ws2.auto_filter.ref = f"A{hr}:H{r - 1}"

from collections import Counter
counts = Counter()
for rec in records:
    dup = len(by_person[rec['passport'] or (rec['salon'], rec['norm'])]) > 1
    if rec['absconded']:
        counts['Absconded'] += 1
    elif rec['terminated']:
        counts['Terminated'] += 1
    elif rec['name'] in calculated_names and rec['csd']:
        counts['Calculated'] += 1
    else:
        counts['Awaiting Contract'] += 1

sr = r + 1
ws2.cell(row=sr, column=1, value='Summary').font = Font(name=FONT_NAME, bold=True, size=10)
sr += 1
for k in ['Calculated', 'Awaiting Contract', 'Terminated', 'Absconded']:
    if counts.get(k):
        ws2.cell(row=sr, column=1, value=k).font = normal_font
        ws2.cell(row=sr, column=2, value=counts[k]).font = normal_font
        sr += 1

# move new sheet right after 'Gratuity Calculator'
order = wb.sheetnames
idx = order.index('Gratuity Calculator') + 1
order.remove('All Staff - Gratuity Status')
order.insert(idx, 'All Staff - Gratuity Status')
wb._sheets = [wb[s] for s in order]

wb.save(FILE)
recalc(FILE)
print('Done. Counts:', dict(counts))
