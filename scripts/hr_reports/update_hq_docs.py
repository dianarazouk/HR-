import openpyxl
from datetime import datetime
from openpyxl.styles import Font

FILE = 'Comprehensive_Salons_and_Staff_Register_EN_with_Payroll_12_REVIEWED.xlsx'
wb = openpyxl.load_workbook(FILE, data_only=False)
FONT_NAME = 'Arial'
normal_font = Font(name=FONT_NAME, size=10)
bold_font = Font(name=FONT_NAME, bold=True, size=10)

# ---------- 1. Add confirmed entries to MOHRE Verification (establishment 1278237) ----------
ws = wb['MOHRE Verification']
legend_row = None
for r in range(1, ws.max_row + 1):
    if ws.cell(row=r, column=1).value == 'Legend:':
        legend_row = r
        break

new_rows = [
    ('__HEADER__', 'Individual work-permit documents received 17/09/2026 for 5 THE LAB - HQ employees - all '
                    'show establishment 1278237 ("THE LAB GENTS SALON AND SPA L.L.C"), NOT 2063109 (the Branch '
                    'establishment) as previously stated by the owner. This is a partial confirmation only (5 of '
                    '23 HQ staff) - not a full official roster for 1278237, so the remaining 18 stay "Awaiting '
                    'Official MOHRE List" rather than being flagged as confirmed missing.'),
    ('THE LAB - HQ', 'ABDULRAHMAN SAMIR HALLAK', '✓ Confirmed', '133075884', '22/Nov/2027',
     'Work permit card and employment contract on file - establishment 1278237. Join date confirmed 24/11/2023.'),
    ('THE LAB - HQ', 'AHMED IBRAHIM ALIBRAHIM', '✓ Confirmed', '133319297', '18/Oct/2027',
     'Work permit card and RENEWED contract on file (effective 19/10/2025) - establishment 1278237. '
     'Original hire date not on file - only the renewal contract was provided.'),
    ('THE LAB - HQ', 'ALI FAISAL ALOBEID', '✓ Confirmed', '133103152', '19/Oct/2027',
     'Work permit card and RENEWED contract on file (effective 20/10/2025) - establishment 1278237. '
     'Original hire date not on file - only the renewal contract was provided.'),
    ('THE LAB - HQ', 'ANALIZA DOMONDON FLORES', '✓ Confirmed', '133968179', '20/Nov/2027',
     'Work permit card and RENEWED contract on file (approx. effective 01/12/2025 - scan partially illegible) '
     '- establishment 1278237. Original hire date not on file.'),
    ('THE LAB - HQ', 'CHEA NILDA ATIENZA CANILLAS', '✓ Confirmed', '123496449', '25/Feb/2027',
     'Work permit card and employment contract on file - establishment 1278237. Join date confirmed 26/02/2025.'),
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

# ---------- 2. Fill in real Join Dates (Contract Start Date, column AG=33) on THE LAB - HQ ----------
ws2 = wb['THE LAB - HQ']
join_dates = {
    'ABDULRAHMAN SAMIR HALLAK': datetime(2023, 11, 24),
    'CHEA NILDA ATIENZA CANILLAS': datetime(2025, 2, 26),
}
notes_append = {
    'AHMED IBRAHIM ALIBRAHIM': 'Renewed contract on file, effective 19/10/2025 - original hire date not on file '
                                '(only the renewal was provided); do not use this date for gratuity.',
    'ALI FAISAL ALOBEID': 'Renewed contract on file, effective 20/10/2025 - original hire date not on file '
                           '(only the renewal was provided); do not use this date for gratuity.',
    'ANALIZA DOMONDON FLORES': 'Renewed contract on file, approx. effective 01/12/2025 (scan partially illegible) '
                                '- original hire date not on file; do not use this date for gratuity.',
}
updated = []
for r in range(5, ws2.max_row + 1):
    name = ws2.cell(row=r, column=3).value
    if not name:
        continue
    name_u = str(name).upper().strip()
    if name_u in join_dates:
        cell = ws2.cell(row=r, column=33, value=join_dates[name_u])
        cell.font = normal_font
        cell.number_format = 'dd/mm/yyyy'
        updated.append((name, join_dates[name_u]))
    if name_u in notes_append:
        notes_cell = ws2.cell(row=r, column=30)
        existing = notes_cell.value or ''
        notes_cell.value = (existing + (' ' if existing else '') + notes_append[name_u]).strip()
        notes_cell.font = normal_font
        updated.append((name, 'note added'))

wb.save(FILE)
print('Join Date / notes updated for:', updated)
