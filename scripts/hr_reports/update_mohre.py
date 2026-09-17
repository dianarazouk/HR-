import openpyxl
from datetime import datetime
from openpyxl.styles import Font

FILE = 'Comprehensive_Salons_and_Staff_Register_EN_with_Payroll_12_REVIEWED.xlsx'
wb = openpyxl.load_workbook(FILE, data_only=False)
ws = wb['MOHRE Verification']

FONT_NAME = 'Arial'
normal_font = Font(name=FONT_NAME, size=10)
bold_font = Font(name=FONT_NAME, bold=True, size=10)

# find current Legend row
legend_row = None
for r in range(1, ws.max_row + 1):
    if ws.cell(row=r, column=1).value == 'Legend:':
        legend_row = r
        break

new_rows = []
new_rows.append(('__HEADER__', 'Updated from official MOHRE establishment lists received 17/09/2026 '
                                '(ALEKSANDRA RP247180927AE, UNIQUE YOU RP247180977AE, NISANTASI RP247181016AE, '
                                'THE LAB - Branch RP247181061AE/RP247181098AE).', None, None, None, None))

# ALEKSANDRA - previously not on file, now confirmed on the official list
new_rows += [
    ('ALEKSANDRA', 'JENNALYN GUADALUPE GEONIGO', '✓ Added', '141042347', '29/07/2028',
     'On official MOHRE list - was missing from internal register, now added.'),
    ('ALEKSANDRA', 'MUSTAFA GULEC', '✓ Added', '142417353', '04/11/2026',
     'On official MOHRE list (Pre-Approval for Work Permit) - was missing from internal register, now added.'),
    ('ALEKSANDRA', 'MAHYM NURYYEVA', '✓ Added', '141294323', '04/10/2026',
     'On official MOHRE list (Pre-Approval for Work Permit) - was missing from internal register, now added.'),
]

# NISANTASI - previously not on file, now confirmed
new_rows += [
    ('NISANTASI', 'EMRE MISIRLI', '✓ Updated', '142110309', '23/04/2028',
     'Card renewed - new card 142110309 (was 138096442), same expiry 23/04/2028.'),
    ('NISANTASI', 'AIZHAN KENZHIBAEVA', '✓ Added', '140768362', '21/07/2028',
     'On official MOHRE list - was missing from internal register, now added.'),
    ('NISANTASI', 'OMER SAYAN', '✓ Added', '141083391', '28/09/2026',
     'On official MOHRE list (Pre-Approval for Work Permit) - was missing from internal register, now added.'),
    ('NISANTASI', 'ALTYNAY GURDOVA', '✓ Added', '141337569', '05/10/2026',
     'On official MOHRE list (Pre-Approval for Work Permit) - was missing from internal register, now added.'),
    ('NISANTASI', 'ALBINA MEDETOVNA ABYLGAZIEVA', '✓ Added', '142303820', '01/11/2026',
     'On official MOHRE list (Pre-Approval for Work Permit) - was missing from internal register, now added.'),
    ('NISANTASI', 'AYLAR ATAMYRADOVA', '✓ Added', '141120426', '29/09/2026',
     'On official MOHRE list (Pre-Approval for Work Permit) - was missing from internal register, now added.'),
]

# THE LAB - Branch - brand new establishment verification (RP247181061AE / establishment 2063109)
lab_branch = [
    ('ABDUL SAMEEH MAMMADKUNHI PAKIYARE ABDULLA', '123058316', '08/01/2027'),
    ('AHMED HELMY HANAFY ATIA', '139134701', '14/04/2027'),
    ('HUSAM ALI MAHMOUD FAYYAD', '141307640', '06/08/2028'),
    ('HUSSEIN MAHMOUD BASAL', '136911104', '10/03/2028'),
    ('AISHA TWINOMUGISHA', '119582194', '04/11/2026'),
    ('ABEGAIL YANSON CEDILLO', '130450781', '03/09/2027'),
    ('JENICYL DIAMA ALQUISA', '126559522', '21/05/2027'),
    ('RECHELLE MEDINA SAMONTE', '137734406', '09/04/2028'),
    ('RONALYN MEDINA SAMONTE', '119799337', '10/11/2026'),
    ('ABDULKAFI ABDULRAZAK ALJAAFAR', '138611107', '13/03/2028'),
    ('BASEL MUSTAFA KAMAL AL HAMDAN', '137587539', '12/03/2028'),
    ('MOHAMMAD KAMAL AHMAD ALQUDAH', '139410526', '10/06/2028'),
    ('JUARNIE CAMARIN LUYAHAN', '137014888', '14/02/2028'),
    ('MOHAMAD MOURAD EID', '138692398', '05/03/2028'),
    ('PAVEL KUROCHKIN', '141806904', '26/12/2027'),
]
for name, card, exp in lab_branch:
    new_rows.append(('THE LAB - Branch', name, '✓ Confirmed', card, exp,
                      'Matches register - official MOHRE list received 17/09/2026 (establishment 2063109).'))

# insert before the Legend row
insert_at = legend_row
ws.insert_rows(insert_at, amount=len(new_rows) + 1)  # +1 for a blank spacer row

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
# blank spacer row
r += 1

wb.save(FILE)
print(f'Inserted {len(new_rows)} rows into MOHRE Verification (before row {insert_at}).')
print('New Legend row is now at', legend_row + len(new_rows) + 1)
