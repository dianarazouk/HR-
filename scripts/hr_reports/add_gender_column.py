import openpyxl
import re
from openpyxl.styles import Font, PatternFill, Alignment

FILE = 'Comprehensive_Salons_and_Staff_Register_EN_with_Payroll_12_REVIEWED.xlsx'
wb = openpyxl.load_workbook(FILE, data_only=False)
FONT_NAME = 'Arial'
header_font = Font(name=FONT_NAME, bold=True, size=10)
normal_font = Font(name=FONT_NAME, size=10)
CONFIRMED_FILL = PatternFill('solid', fgColor='C6EFCE')
INFERRED_FILL = PatternFill('solid', fgColor='FFF2CC')

# 'M'/'F' = confirmed from an Emirates ID / MOHRE work permit card actually seen.
# 'M?'/'F?' = inferred from name/nationality only - not confirmed from a document, please verify.
GENDER = {
    'THE LAB - Branch': {
        'ABDUL SAMEEH MAMMADKUNHI PAKIYARE ABDULLA': 'M?',
        'AHMED HELMY HANAFY ATIA': 'M',              # Barber (Men's grooming role)
        'HUSSEIN MAHMOUD BASAL': 'M',                 # Barber
        'AISHA TWINOMUGISHA': 'F?',
        'ABEGAIL YANSON CEDILLO': 'F?',
        'JENICYL DIAMA ALQUISA': 'F?',
        'RECHELLE MEDINA SAMONTE': 'F?',
        'RONALYN MEDINA SAMONTE': 'F?',
        'ABDULKAFI ABDULRAZAK ALJAAFAR': 'M',          # Barber
        'BASEL MUSTAFA KAMAL AL HAMDAN': 'M',          # Barber
        'MOHAMMAD KAMAL AHMAD ALQUDAH': 'M',           # Barber + ID confirmed
        'JUARNIE CAMARIN LUYAHAN': '',                 # unclear from name - left blank
        'MOHAMAD MOURAD EID': 'M',                     # ID/passport confirmed
        'PAVEL KUROCHKIN': 'M?',
        'HUSAM ALI MAHMOUD FAYYAD': 'M',               # Barber
    },
    'THE LAB - HQ': {
        'HUSAM ALI MAHMOUD FAYYAD': 'M',
        'FLOVELYN IRIS ESCUBIO': 'F',                  # ID confirmed (Sex: F)
        'KARIM MOUNIR MOHAMED HAFEZ MALEK': 'M',       # Barber + ID confirmed
        'MOHAMAD SAMIR HAMAMA': 'M',                   # Barber + ID confirmed
        'CHEA NILDA ATIENZA CANILLAS': 'F',            # ID confirmed
        'JENNIFER MON GASPAR': 'F',                    # ID confirmed (Sex: F)
        'VLADISLAV OBLONSKII': 'F',                    # per owner's own reference ("she") - no ID on file
        'ANALIZA DOMONDON FLORES': 'F',                # ID/photo confirmed
        'ABDULRAHMAN SAMIR HALLAK': 'M',               # Barber + ID confirmed (Sex: M)
        'AHMED IBRAHIM ALIBRAHIM': 'M',                # Barber + ID/photo confirmed
        'ALI FAISAL ALOBEID': 'M',                     # Barber + ID confirmed
        'HAITHEM TRABELSI': 'M',                       # Barber + ID confirmed (Sex: M)
        'MOHAMAD WALID ABBAS': 'M',                    # Barber + ID confirmed
        'MOHAMED ELHARMEL': 'M',                       # Barber + ID confirmed (Sex: M)
        'SALWAN BASHAR AL HAKIM': 'M',                 # Barber + ID confirmed (Sex: M)
        'GINA BANYARES DE MESA': 'F',                  # ID confirmed
        'GLORY FELLY SERENIO ANTESODA': 'F',           # ID confirmed (Sex: F)
        'MOHAMMED AHMED ABDALLAH ALJUNEIDI': 'M',      # ID confirmed
        'GLADYS BALBIRAN RAMPAS': 'F',                 # ID confirmed (Sex: F) - left company
    },
}

GENDER_COL = 37  # column AK, first free column after Entry Permit Status (AJ)

for sname, mapping in GENDER.items():
    ws = wb[sname]
    ws.cell(row=4, column=GENDER_COL, value='Gender (M/F)').font = header_font
    ws.column_dimensions[openpyxl.utils.get_column_letter(GENDER_COL)].width = 12
    matched = 0
    for r in range(5, ws.max_row + 1):
        name = ws.cell(row=r, column=3).value
        if not name:
            continue
        clean_name = re.sub(r'^🔴\s*(ABSCONDED|TERMINATED|LEFT)\s*[—-]\s*', '', str(name).strip(), flags=re.IGNORECASE).strip()
        name_u = clean_name.upper()
        val = mapping.get(name_u)
        if val is None:
            print(f'  WARNING: no gender mapping for "{name}" in {sname} row {r}')
            continue
        cell = ws.cell(row=r, column=GENDER_COL, value=val if val else None)
        cell.font = normal_font
        cell.alignment = Alignment(horizontal='center')
        if val.endswith('?'):
            cell.fill = INFERRED_FILL
        elif val:
            cell.fill = CONFIRMED_FILL
        matched += 1
    print(sname, ': set gender for', matched, 'rows')

wb.save(FILE)
print('Saved.')
