import openpyxl
from datetime import datetime
from openpyxl.styles import Font, PatternFill

FILE = 'Comprehensive_Salons_and_Staff_Register_EN_with_Payroll_12_REVIEWED.xlsx'
wb = openpyxl.load_workbook(FILE, data_only=False)
FONT_NAME = 'Arial'
normal_font = Font(name=FONT_NAME, size=10)
bold_font = Font(name=FONT_NAME, bold=True, size=10)
TODAY = datetime(2026, 9, 17)

# ---------- 1. Add confirmed entries to MOHRE Verification (establishment 1278237) ----------
ws = wb['MOHRE Verification']
legend_row = None
for r in range(1, ws.max_row + 1):
    if ws.cell(row=r, column=1).value == 'Legend:':
        legend_row = r
        break

new_rows = [
    ('__HEADER__', 'Individual work-permit documents received 17/09/2026 for the last 2 outstanding THE LAB - HQ '
                    'employees - establishment 1278237 again. Vladislav Oblonskii confirmed by the owner as NOT '
                    'actually working here - her work permit needs cancellation, not confirmation (see Action '
                    'Tracker and her row notes).'),
    ('THE LAB - HQ', 'MOHAMAD SAMIR HAMAMA', '✓ Confirmed', '127325391', '15/Jun/2027',
     'Work permit card and ORIGINAL employment contract on file - establishment 1278237. '
     'Join date confirmed 03/06/2025.'),
    ('THE LAB - HQ', 'GLORY FELLY SERENIO ANTESODA', '✓ Confirmed', '136466330', '21/Oct/2027',
     'Work permit card and RENEWED contract on file (effective 22/10/2025) - establishment 1278237. '
     'Original hire date not on file - only the renewal contract was provided.'),
]

insert_at = legend_row
ws.insert_rows(insert_at, amount=len(new_rows) + 1)
r = insert_at
for row in new_rows:
    if row[0] == '__HEADER__':
        ws.cell(row=r, column=1, value=row[1]).font = bold_font
        r += 1
        continue
    salon, name, status, card, exp, detail = row
    ws.cell(row=r, column=1, value=salon)
    ws.cell(row=r, column=2, value=name)
    ws.cell(row=r, column=3, value=status)
    ws.cell(row=r, column=4, value=card)
    ws.cell(row=r, column=5, value=exp)
    ws.cell(row=r, column=6, value=detail)
    for c in range(1, 7):
        ws.cell(row=r, column=c).font = normal_font
    r += 1
print(f'Inserted {len(new_rows)} rows into MOHRE Verification (before row {insert_at}).')

# ---------- 2. Join Date / notes on THE LAB - HQ ----------
ws_hq = wb['THE LAB - HQ']
join_dates_hq = {'MOHAMAD SAMIR HAMAMA': datetime(2025, 6, 3)}
notes_append_hq = {
    'GLORY FELLY SERENIO ANTESODA': 'Renewed contract on file, effective 22/10/2025 - original hire date not on '
                                     'file (only the renewal was provided); do not use this date for gratuity.',
}
for r in range(5, ws_hq.max_row + 1):
    name = ws_hq.cell(row=r, column=3).value
    if not name:
        continue
    name_u = str(name).upper().strip()
    if name_u in join_dates_hq:
        cell = ws_hq.cell(row=r, column=33, value=join_dates_hq[name_u])
        cell.font = normal_font
        cell.number_format = 'dd/mm/yyyy'
        print('HQ Join Date set:', name, join_dates_hq[name_u])
    if name_u in notes_append_hq:
        notes_cell = ws_hq.cell(row=r, column=30)
        existing = notes_cell.value or ''
        notes_cell.value = (existing + (' ' if existing else '') + notes_append_hq[name_u]).strip()
        notes_cell.font = normal_font

# ---------- 3. Vladislav Oblonskii - not actually working here, cancel her work permit ----------
CANCEL_FILL = PatternFill('solid', fgColor='FFC7CE')
for r in range(5, ws_hq.max_row + 1):
    name = ws_hq.cell(row=r, column=3).value
    if name and str(name).upper().strip() == 'VLADISLAV OBLONSKII':
        ws_hq.cell(row=r, column=7, value='Cancellation').font = normal_font  # Action Type
        notes_cell = ws_hq.cell(row=r, column=30)
        existing = notes_cell.value or ''
        note = ('🔴 CONFIRMED BY OWNER (17/09/2026): does NOT actually work at this salon - work permit '
                f'(Card 127834345) registered but never used. Process MOHRE work permit cancellation with PRO. '
                f'Passport 663346172.')
        notes_cell.value = (existing + (' ' if existing else '') + note).strip()
        notes_cell.font = normal_font
        for c in range(1, ws_hq.max_column + 1):
            ws_hq.cell(row=r, column=c).fill = CANCEL_FILL
        print('Flagged Vladislav Oblonskii for cancellation at row', r)

# ---------- 4. Action Tracker item for Vladislav's cancellation ----------
wsa = wb['Action Tracker']
last_num = 0
r = 5
while wsa.cell(row=r, column=1).value is not None:
    v = wsa.cell(row=r, column=1).value
    if isinstance(v, (int, float)):
        last_num = max(last_num, v)
    r += 1
new_row = r
new_num = int(last_num) + 1
border_font = Font(name=FONT_NAME, size=10)
vals = [
    new_num, TODAY, 'Work Permit', 'THE LAB - HQ', 'Vladislav Oblonskii',
    'Owner confirmed she does not actually work here - cancel work permit (Card 127834345) with MOHRE/PRO.',
    'Urgent', 'PRO', 'Not Started', None, 'Passport 663346172.'
]
priority_fill = PatternFill('solid', fgColor='FCE4D6')
for c, v in enumerate(vals, start=1):
    cell = wsa.cell(row=new_row, column=c, value=v)
    cell.font = border_font
    if c == 2:
        cell.number_format = 'dd/mm/yyyy'
wsa.cell(row=new_row, column=7).fill = priority_fill
print('Added Action Tracker item', new_num, 'for Vladislav Oblonskii cancellation')

wb.save(FILE)
print('Saved.')
