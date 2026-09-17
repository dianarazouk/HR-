import openpyxl, re, shutil, subprocess, sys, os
from collections import defaultdict
from difflib import SequenceMatcher
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

def recalc(path):
    """Recalculate all formulas via LibreOffice so cached values are fresh.
    openpyxl strips cached formula values on every save, so this MUST run
    right before any data_only=True read and again after the final save."""
    result = subprocess.run([sys.executable, os.path.join(SCRIPT_DIR, 'recalc.py'), path, '270'],
                             capture_output=True, text=True)
    print('recalc:', result.stdout.strip() or result.stderr.strip())

SRC = sys.argv[1] if len(sys.argv) > 1 else 'input.xlsx'
OUT = sys.argv[2] if len(sys.argv) > 2 else 'Comprehensive_Salons_and_Staff_Register_EN_with_Payroll_12_REVIEWED.xlsx'
if os.path.abspath(SRC) != os.path.abspath(OUT):
    shutil.copy(SRC, OUT)
recalc(OUT)

wb = openpyxl.load_workbook(OUT, data_only=True)  # read cached values for analysis
salons = ['ALEKSANDRA', 'UNIQUE YOU', 'NISANTASI', 'THE LAB - Branch', 'THE LAB - HQ', 'HAIR TAG']
verified_salons = {'ALEKSANDRA', 'UNIQUE YOU', 'NISANTASI', 'THE LAB - Branch', 'THE LAB - HQ'}

def norm(s):
    if not s:
        return ''
    s = str(s).upper()
    s = re.sub(r'[^A-Z ]', '', s)
    s = re.sub(r'\s+', ' ', s).strip()
    return s

def salon_norm(s):
    return norm(s)

# ---------- 1. Load employee records from all 6 salon sheets ----------
records = []
for sname in salons:
    ws = wb[sname]
    header = [ws.cell(row=4, column=c).value for c in range(1, ws.max_column + 1)]
    for r in range(5, ws.max_row + 1):
        row = [ws.cell(row=r, column=c).value for c in range(1, ws.max_column + 1)]
        if row[2] is None:
            continue
        d = dict(zip(header, row))
        d['__salon'] = sname
        d['__row'] = r
        raw_name = str(d['Name'])
        d['__absconded'] = 'ABSCONDED' in raw_name.upper()
        d['__terminated'] = 'TERMINATED' in raw_name.upper()
        clean_name = re.sub(r'^🔴\s*(ABSCONDED|TERMINATED)\s*[—-]\s*', '', raw_name.strip(), flags=re.IGNORECASE).strip()
        d['__cleanname'] = clean_name
        d['__norm'] = norm(clean_name)
        records.append(d)

# ---------- 2. Load MOHRE Verification sheet into a lookup ----------
mws = wb['MOHRE Verification']
mohre = defaultdict(list)  # (salon_norm, name_norm) -> list of entries
for r in range(5, mws.max_row + 1):
    salon_m = mws.cell(row=r, column=1).value
    if salon_m == 'Legend:':
        break
    row = [mws.cell(row=r, column=c).value for c in range(1, 7)]
    salon_m, name_m, status, card, expiry, detail = row
    if not salon_m or not name_m or salon_m == 'Salon':
        continue
    key = (salon_norm(salon_m), norm(name_m))
    mohre[key].append({'status': status, 'card': card, 'expiry': expiry, 'detail': detail})

mohre_by_salon = defaultdict(list)  # salon_norm -> [(name_norm, entry), ...]
for (sn, nn), entries in mohre.items():
    for e in entries:
        mohre_by_salon[sn].append((nn, e))

def token_sort(s):
    return ' '.join(sorted(s.split()))

def name_ratio(a, b):
    return max(SequenceMatcher(None, a, b).ratio(),
               SequenceMatcher(None, token_sort(a), token_sort(b)).ratio())

# THE LAB - Branch and THE LAB - HQ are two physical locations under one MOHRE
# establishment license (2063109) - confirmed by the salon owner 17/09/2026. One
# official work-permit list covers staff at both locations, so match HQ against it too.
MOHRE_SALON_ALIAS = {'THE LAB - HQ': 'THE LAB - Branch'}

def find_mohre_entry(salon, name_norm):
    """Exact match first, else best fuzzy match (>=0.84) within the same salon - catches
    spelling differences ('ZHARKYNAI' vs 'ZHARKYNAIL') and first/last name order swaps
    ('MUATOVA DIANA' vs 'DIIANA MURATOVA', 'AIZHAN KENZHIBAEVA' vs 'KENZHIBAEVA AIZHA')."""
    salon = MOHRE_SALON_ALIAS.get(salon, salon)
    key = (salon_norm(salon), name_norm)
    if key in mohre:
        return mohre[key][0], True
    best, best_ratio = None, 0.0
    for nn, e in mohre_by_salon[salon_norm(salon)]:
        ratio = name_ratio(name_norm, nn)
        if ratio > best_ratio:
            best, best_ratio = e, ratio
    if best is not None and best_ratio >= 0.84:
        return best, False
    return None, False

# ---------- 3. Duplicate detection (same passport across salon sheets) ----------
bypass = defaultdict(list)
for d in records:
    p = d.get('Passport No.')
    if p:
        bypass[str(p).strip().upper()].append(d)

dup_groups = {p: v for p, v in bypass.items() if len(v) > 1}

for d in records:
    p = d.get('Passport No.')
    d['__dup_group'] = str(p).strip().upper() if p and str(p).strip().upper() in dup_groups else None

# near-duplicate name pairs (different passport, similar name) - informational only
names_by_norm = defaultdict(list)
for d in records:
    if d['__norm']:
        names_by_norm[d['__norm']].append(d)
unique_norms = list(names_by_norm.keys())
near_dup_pairs = []
for i in range(len(unique_norms)):
    for j in range(i + 1, len(unique_norms)):
        a, b = unique_norms[i], unique_norms[j]
        if SequenceMatcher(None, a, b).ratio() > 0.82:
            pa = {x.get('Passport No.') for x in names_by_norm[a]}
            pb = {x.get('Passport No.') for x in names_by_norm[b]}
            if not (pa & pb - {None}):
                near_dup_pairs.append((a, b))

near_dup_lookup = defaultdict(set)
for a, b in near_dup_pairs:
    near_dup_lookup[a].add(b)
    near_dup_lookup[b].add(a)

# ---------- 4. Categorize each record ----------
def categorize(d):
    notes = (d.get('Notes') or '')
    action = (d.get('Action Type') or '')
    wps = (d.get('Work Permit Status') or '')
    combined = f"{notes} {action} {wps}".lower()

    mentry, exact = find_mohre_entry(d['__salon'], d['__norm'])
    d['__mohre_exact'] = exact

    if d['__absconded']:
        return ('Absconded', 'Flagged absconded / left with no notice - see Absconded tab.')

    if d['__terminated']:
        return ('Terminated', 'Employment ended (termination) - see Warnings / Gratuity Calculator for settlement status.')

    if 'confirmed by manager' in combined and 'no work permit' in combined:
        return ('Undocumented - No Permit (Manager Confirmed)',
                'Manager has confirmed this person works with NO MOHRE work permit at all. Urgent legal/PRO action needed.')

    if mentry and 'own visa' in str(mentry['status']).lower():
        return ('Own Visa - No Salon Work Permit',
                f"MOHRE list confirms: NOT on official salon work-permit list. {mentry['detail'] or ''}".strip())

    if 'independent residency' in combined or 'own visa' in combined:
        return ('Own Visa - No Salon Work Permit',
                'Internal register notes independent residency / own visa - no salon-sponsored work permit.')

    if mentry:
        status = str(mentry['status'])
        spelling_note = '' if exact else ' (matched by similar spelling - confirm same person.)'
        if 'not confirmed' in status.lower():
            return ('MOHRE Mismatch - Needs Review',
                    f"On internal register but NOT found on official MOHRE list under this salon.{spelling_note} {mentry['detail'] or ''}".strip())
        if 'row added' in status.lower():
            return ('MOHRE Mismatch - Needs Review',
                    f"On official MOHRE list but was missing from internal register (now added).{spelling_note} {mentry['detail'] or ''}".strip())
        if any(k in status.lower() for k in ['confirmed', 'updated', 'upgraded', 'added', 'in process']):
            return ('MOHRE Confirmed', (mentry['detail'] or 'Matches official MOHRE establishment list.') + spelling_note)

    if d['__salon'] not in verified_salons:
        if 'no work permit' in combined:
            return ('No Work Permit - Salon Not Yet MOHRE-Verified',
                    'No work permit on internal register, and this salon\'s official MOHRE list has not been received yet - verify manually.')
        return ('Awaiting Official MOHRE List',
                'This salon\'s official MOHRE establishment list has not been received yet - cannot cross-check.')

    # Verified salon (ALEKSANDRA / UNIQUE YOU / NISANTASI) with no MOHRE match at all -
    # this person does not appear on the official government work-permit list under this salon.
    if 'no work permit' in combined:
        return ('No Work Permit - Not on MOHRE List',
                'No work permit on internal register and no matching entry found on the official MOHRE list.')
    return ('MOHRE Mismatch - Needs Review',
            'Not found on the official MOHRE work-permit list for this salon (checked by name, incl. similar spellings) - '
            'verify whether their work permit/residency is genuinely sponsored by this salon.')

for d in records:
    cat, detail = categorize(d)
    d['__category'] = cat
    d['__category_detail'] = detail

# ---------- 5. Build output workbook (append review sheets, keep originals untouched) ----------
wb_out = openpyxl.load_workbook(OUT)  # fresh load, keep formulas intact for original sheets

FONT_NAME = 'Arial'
header_font = Font(name=FONT_NAME, bold=True, color='FFFFFF', size=11)
header_fill = PatternFill('solid', fgColor='305496')
title_font = Font(name=FONT_NAME, bold=True, size=14, color='305496')
sub_font = Font(name=FONT_NAME, italic=True, size=10, color='595959')
normal_font = Font(name=FONT_NAME, size=10)
bold_font = Font(name=FONT_NAME, bold=True, size=10)
thin = Side(style='thin', color='BFBFBF')
border = Border(left=thin, right=thin, top=thin, bottom=thin)

CAT_COLORS = {
    'Absconded': 'C00000',
    'Terminated': '808080',
    'Undocumented - No Permit (Manager Confirmed)': 'FF0000',
    'Own Visa - No Salon Work Permit': 'BDD7EE',
    'MOHRE Mismatch - Needs Review': 'FFFF00',
    'MOHRE Confirmed': 'C6EFCE',
    'No Work Permit - Salon Not Yet MOHRE-Verified': 'FCE4D6',
    'Awaiting Official MOHRE List': 'D9D9D9',
    'No Work Permit - Not on MOHRE List': 'FFC7CE',
    'OK - No Flag': 'FFFFFF',
}
CAT_TEXT_WHITE = {'Absconded', 'Terminated', 'Undocumented - No Permit (Manager Confirmed)'}

def style_header_row(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(vertical='center', wrap_text=True)
        cell.border = border

# ===== Sheet A: Duplicate Employees =====
if 'Duplicate Employees' in wb_out.sheetnames:
    del wb_out['Duplicate Employees']
ws_dup = wb_out.create_sheet('Duplicate Employees')
ws_dup.sheet_view.showGridLines = False
ws_dup['A1'] = 'Duplicate Employees - Cross-Salon Matches'
ws_dup['A1'].font = title_font
ws_dup['A2'] = ('Same person (matched by passport number) found registered under more than one salon / work permit. '
                'Review each group below - only one active work permit per person is normally valid under UAE labour law.')
ws_dup['A2'].font = sub_font
ws_dup.merge_cells('A1:H1')
ws_dup.merge_cells('A2:H2')
ws_dup.row_dimensions[2].height = 30
ws_dup['A2'].alignment = Alignment(wrap_text=True, vertical='top')

headers = ['Group #', 'Passport No.', 'Name', 'Salon', 'Work Permit Card No.', 'Work Permit Status', 'Basic Salary', 'Notes / Action Needed']
hr = 4
for i, h in enumerate(headers, start=1):
    ws_dup.cell(row=hr, column=i, value=h)
style_header_row(ws_dup, hr, len(headers))

r = hr + 1
group_num = 0
for passport, group in sorted(dup_groups.items(), key=lambda x: x[1][0]['__cleanname']):
    group_num += 1
    for idx, d in enumerate(group):
        note = ''
        if idx == 0:
            note = 'DUPLICATE - same passport appears under multiple salons. Confirm which salon holds the genuine, current work permit; cancel/investigate the others.'
        ws_dup.cell(row=r, column=1, value=group_num if idx == 0 else None)
        ws_dup.cell(row=r, column=2, value=passport)
        ws_dup.cell(row=r, column=3, value=d['__cleanname'])
        ws_dup.cell(row=r, column=4, value=d['__salon'])
        ws_dup.cell(row=r, column=5, value=d.get('Work Permit Card No.'))
        ws_dup.cell(row=r, column=6, value=d.get('Work Permit Status'))
        ws_dup.cell(row=r, column=7, value=d.get('Basic Salary (Monthly)'))
        ws_dup.cell(row=r, column=8, value=note)
        fill = PatternFill('solid', fgColor='FFC7CE' if group_num % 2 else 'FFE6E6')
        for c in range(1, len(headers) + 1):
            cell = ws_dup.cell(row=r, column=c)
            cell.font = normal_font
            cell.fill = fill
            cell.border = border
            cell.alignment = Alignment(vertical='top', wrap_text=True)
        r += 1
    r += 1  # blank spacer row between groups

# near-duplicate names (different passports) as a caution note section
r += 1
ws_dup.cell(row=r, column=1, value='Similar Names (different passport numbers - NOT duplicates, but verify no confusion):')
ws_dup.cell(row=r, column=1).font = bold_font
r += 1
seen = set()
for a, b in near_dup_pairs:
    pair = tuple(sorted([a, b]))
    if pair in seen:
        continue
    seen.add(pair)
    for nm in pair:
        for d in names_by_norm[nm]:
            ws_dup.cell(row=r, column=3, value=d['__cleanname'])
            ws_dup.cell(row=r, column=4, value=d['__salon'])
            ws_dup.cell(row=r, column=2, value=d.get('Passport No.'))
            for c in range(1, len(headers) + 1):
                ws_dup.cell(row=r, column=c).font = normal_font
                ws_dup.cell(row=r, column=c).border = border
            r += 1

widths = [8, 14, 26, 16, 20, 22, 12, 55]
for i, w in enumerate(widths, start=1):
    ws_dup.column_dimensions[get_column_letter(i)].width = w
ws_dup.freeze_panes = 'A5'

# ===== Sheet B: MOL vs Employee List Match =====
if 'MOL vs Employee Match' in wb_out.sheetnames:
    del wb_out['MOL vs Employee Match']
ws_m = wb_out.create_sheet('MOL vs Employee Match')
ws_m.sheet_view.showGridLines = False
ws_m['A1'] = 'MOL (MOHRE) vs Employee List - Matching & Residency Review'
ws_m['A1'].font = title_font
ws_m['A2'] = ('Every staff member across all 6 salons, matched against the official MOHRE work-permit list '
              '(MOHRE Verification tab, updated 14/09/2026). Use the Status column to see who is fully confirmed, '
              'who only has residency/visa without a genuine salon work permit, who needs review, and who is at a salon whose '
              'official MOHRE list has not been received yet.')
ws_m['A2'].font = sub_font
ws_m.merge_cells('A1:I1')
ws_m.merge_cells('A2:I2')
ws_m.row_dimensions[2].height = 44
ws_m['A2'].alignment = Alignment(wrap_text=True, vertical='top')

# legend
legend_items = [
    ('MOHRE Confirmed', 'On both the official MOHRE list and the salon register - matches.'),
    ('Own Visa - No Salon Work Permit', 'Works at the salon but residency/visa is NOT sponsored by the salon (independent, spouse, or MOHRE list shows "Own visa"). No genuine company work permit exists for them here.'),
    ('MOHRE Mismatch - Needs Review', 'Discrepancy between the official MOHRE list and the salon register (missing row, or not found on official list) - needs your review.'),
    ('Undocumented - No Permit (Manager Confirmed)', 'Manager has confirmed this person works with NO work permit at all. Highest priority / legal risk.'),
    ('Absconded', 'Flagged as absconded / left with no notice.'),
    ('Terminated', 'Employment ended (termination) - not absconding; check Warnings/Gratuity for settlement status.'),
    ('No Work Permit - Salon Not Yet MOHRE-Verified', "No work permit on file, and this salon's official MOHRE list has not been received yet."),
    ('Awaiting Official MOHRE List', "This salon's official MOHRE list has not been received yet - cannot cross-check."),
    ('No Work Permit - Not on MOHRE List', 'No work permit on file and no match on the official MOHRE list (verified salons only).'),
]
lr = 3
ws_m.cell(row=lr, column=1, value='Legend:').font = bold_font
lr += 1
for cat, desc in legend_items:
    c1 = ws_m.cell(row=lr, column=1, value='  ')
    c1.fill = PatternFill('solid', fgColor=CAT_COLORS[cat])
    c2 = ws_m.cell(row=lr, column=2, value=f'{cat} - {desc}')
    c2.font = Font(name=FONT_NAME, size=9, italic=True, color='595959')
    ws_m.merge_cells(start_row=lr, start_column=2, end_row=lr, end_column=9)
    lr += 1
lr += 1

headers = ['Salon', 'Name', 'Passport No.', 'Employee-List Work Permit Status',
           'MOHRE Verification Status', 'MOHRE Card No.', 'Duplicate?', 'Status / Flag', 'Notes - Action Needed']
hr = lr
for i, h in enumerate(headers, start=1):
    ws_m.cell(row=hr, column=i, value=h)
style_header_row(ws_m, hr, len(headers))

r = hr + 1
for d in sorted(records, key=lambda x: (x['__salon'], x['__cleanname'])):
    mentry, exact = find_mohre_entry(d['__salon'], d['__norm'])
    if mentry:
        mohre_status = mentry['status'] + ('' if exact else ' (spelling match)')
    else:
        mohre_status = 'Not verified yet' if d['__salon'] not in verified_salons else 'Not on official list'
    mohre_card = mentry['card'] if mentry else ''

    vals = [
        d['__salon'],
        d['__cleanname'],
        d.get('Passport No.'),
        d.get('Work Permit Status'),
        mohre_status,
        mohre_card,
        'YES' if d['__dup_group'] else '',
        d['__category'],
        d['__category_detail'],
    ]
    for c, v in enumerate(vals, start=1):
        cell = ws_m.cell(row=r, column=c, value=v)
        cell.font = normal_font
        cell.border = border
        cell.alignment = Alignment(vertical='top', wrap_text=True)
    fill_color = CAT_COLORS.get(d['__category'], 'FFFFFF')
    status_cell = ws_m.cell(row=r, column=8)
    status_cell.fill = PatternFill('solid', fgColor=fill_color)
    status_cell.font = Font(name=FONT_NAME, size=10, bold=True,
                             color='FFFFFF' if d['__category'] in CAT_TEXT_WHITE else '000000')
    if d['__dup_group']:
        ws_m.cell(row=r, column=7).fill = PatternFill('solid', fgColor='FFC7CE')
        ws_m.cell(row=r, column=7).font = bold_font
    r += 1

widths = [16, 26, 13, 24, 20, 14, 10, 34, 55]
for i, w in enumerate(widths, start=1):
    ws_m.column_dimensions[get_column_letter(i)].width = w
ws_m.freeze_panes = ws_m.cell(row=hr + 1, column=1).coordinate
ws_m.auto_filter.ref = f"A{hr}:I{r-1}"

# summary counts
sr = r + 1
ws_m.cell(row=sr, column=1, value='Summary').font = bold_font
sr += 1
from collections import Counter
counts = Counter(d['__category'] for d in records)
for cat in CAT_COLORS:
    if counts.get(cat):
        ws_m.cell(row=sr, column=1, value=cat).font = normal_font
        ws_m.cell(row=sr, column=2, value=counts[cat]).font = normal_font
        ws_m.cell(row=sr, column=1).fill = PatternFill('solid', fgColor=CAT_COLORS[cat])
        sr += 1
ws_m.cell(row=sr, column=1, value='Duplicate records (cross-salon)').font = normal_font
ws_m.cell(row=sr, column=2, value=sum(len(v) for v in dup_groups.values())).font = normal_font

# move the two new sheets right after 'MOHRE Verification'
order = wb_out.sheetnames
idx = order.index('MOHRE Verification') + 1
for sheet_name in ['Duplicate Employees', 'MOL vs Employee Match']:
    order.remove(sheet_name)
    order.insert(idx, sheet_name)
    idx += 1
wb_out._sheets = [wb_out[s] for s in order]

wb_out.save(OUT)
recalc(OUT)  # openpyxl just stripped every cached formula value again - restore them
print('Saved', OUT)
print('Duplicate groups:', len(dup_groups), 'records involved:', sum(len(v) for v in dup_groups.values()))
print(Counter(d['__category'] for d in records))
