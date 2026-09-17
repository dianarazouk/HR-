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
    ('__HEADER__', 'Individual work-permit documents received 17/09/2026 for 3 more THE LAB - HQ employees - '
                    'establishment 1278237 again. Mohamad Mourad Eid\'s documents matched his existing confirmed '
                    'THE LAB - Branch record (2063109) - no change needed. Gladys Balbiran Rampas\'s document was '
                    'an Emirates ID only (no work permit card or contract) - not enough to confirm; still needs '
                    'her actual MOHRE work permit card or labour contract.'),
    ('THE LAB - HQ', 'KARIM MOUNIR MOHAMED HAFEZ MALEK', '✓ Confirmed', '138166665', '07/Apr/2027',
     'Work permit card and ORIGINAL employment contract on file - establishment 1278237. '
     'Join date confirmed 08/04/2025.'),
    ('THE LAB - HQ', 'MOHAMAD WALID ABBAS', '✓ Confirmed', '135250181', '25/Nov/2027',
     'Work permit card and RENEWED contract on file (effective 26/11/2025) - establishment 1278237. '
     'Original hire date not on file - only the renewal contract was provided.'),
    ('THE LAB - HQ', 'MOHAMED ELHARMEL', '✓ Confirmed', '126525421', '19/Mar/2027',
     'Work permit card and RENEWED contract on file (effective 22/05/2025) - establishment 1278237. '
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

# ---------- 2. Fill in real Join Dates / notes on THE LAB - HQ ----------
ws_hq = wb['THE LAB - HQ']
join_dates_hq = {'KARIM MOUNIR MOHAMED HAFEZ MALEK': datetime(2025, 4, 8)}
notes_append_hq = {
    'MOHAMAD WALID ABBAS': 'Renewed contract on file, effective 26/11/2025 - original hire date not on file '
                            '(only the renewal was provided); do not use this date for gratuity.',
    'MOHAMED ELHARMEL': 'Renewed contract on file, effective 22/05/2025 - original hire date not on file '
                         '(only the renewal was provided); do not use this date for gratuity.',
    'GLADYS BALBIRAN RAMPAS': 'Emirates ID confirms employer as THE LAB GENTS SALON AND SPA L.L.C, but no work '
                               'permit card or labour contract has been provided yet - MOHRE status still unconfirmed.',
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

wb.save(FILE)
print('Saved.')
