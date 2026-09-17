import openpyxl
from datetime import datetime
from openpyxl.styles import Font, PatternFill

FILE = 'Comprehensive_Salons_and_Staff_Register_EN_with_Payroll_12_REVIEWED.xlsx'
wb = openpyxl.load_workbook(FILE, data_only=False)
FONT_NAME = 'Arial'
normal_font = Font(name=FONT_NAME, size=10)
NO_FILL = PatternFill(fill_type=None)
TODAY = datetime(2026, 9, 17)

# ---------- Correction: Gladys Balbiran Rampas is ACTIVE, not left ----------
# Owner clarified 17/09/2026: she does work at THE LAB - HQ; only her Emirates ID is on
# file (employer field confirms THE LAB GENTS SALON AND SPA L.L.C as sponsor) - no separate
# MOHRE work permit card or labour contract has been provided. She also works part-time
# with a "partner house" (a related/partner business) - worth verifying for a duplicate-permit
# situation like the other multi-salon cases already flagged.
ws_hq = wb['THE LAB - HQ']
for r in range(5, ws_hq.max_row + 1):
    name = ws_hq.cell(row=r, column=3).value
    if name and 'GLADYS BALBIRAN RAMPAS' in str(name).upper():
        ws_hq.cell(row=r, column=3, value='GLADYS BALBIRAN RAMPAS').font = normal_font  # remove LEFT prefix
        ws_hq.cell(row=r, column=7, value=None)  # clear Action Type ("Left Company")
        notes_cell = ws_hq.cell(row=r, column=30)
        notes_cell.value = (
            'CORRECTED (17/09/2026): confirmed ACTIVELY working here, not left - only her Emirates ID is on '
            'file (employer field confirms THE LAB GENTS SALON AND SPA L.L.C as sponsor); no separate MOHRE '
            'work permit card or labour contract received yet. Also works PART-TIME with a "partner house" '
            '(related/partner business) - verify this does not create a duplicate work permit situation, '
            'same as other multi-salon staff already flagged.'
        )
        notes_cell.font = normal_font
        for c in range(1, ws_hq.max_column + 1):
            ws_hq.cell(row=r, column=c).fill = NO_FILL
        print('Reverted Gladys Balbiran Rampas to active status at row', r)
        break

# remove the earlier MOHRE Verification header line that described her as unconfirmed-only
# (still true that no WP card exists, so leave the substantive "not enough to confirm" note
# in the earlier header as-is - it remains accurate; just make sure no "left" language lingers)

# ---------- Action Tracker: follow up on her part-time "partner house" work ----------
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
vals = [
    new_num, TODAY, 'Work Permit', 'THE LAB - HQ', 'Gladys Balbiran Rampas',
    'Confirmed active (not left) - only Emirates ID on file, no MOHRE work permit card/contract yet. Also works '
    'part-time with a "partner house" - get her actual work permit card and verify the part-time arrangement '
    "doesn't create a duplicate-permit issue.",
    'Normal', 'PRO', 'Not Started', None, 'Passport P8628752A.'
]
priority_fill = PatternFill('solid', fgColor='FFF2CC')
for c, v in enumerate(vals, start=1):
    cell = wsa.cell(row=new_row, column=c, value=v)
    cell.font = normal_font
    if c == 2:
        cell.number_format = 'dd/mm/yyyy'
wsa.cell(row=new_row, column=7).fill = priority_fill
print('Added Action Tracker item', new_num, 'for Gladys follow-up')

wb.save(FILE)
print('Saved.')
