import openpyxl
import sys
import os
import subprocess
from datetime import datetime, timedelta
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
                                 PageBreak, HRFlowable)
from reportlab.lib.enums import TA_LEFT, TA_CENTER

FILE = sys.argv[1] if len(sys.argv) > 1 else 'Comprehensive_Salons_and_Staff_Register_EN_with_Payroll_12_REVIEWED.xlsx'

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
result = subprocess.run([sys.executable, os.path.join(SCRIPT_DIR, 'recalc.py'), FILE, '270'],
                         capture_output=True, text=True)
print('recalc:', result.stdout.strip() or result.stderr.strip())

wb = openpyxl.load_workbook(FILE, data_only=True)
TODAY = datetime.strptime(sys.argv[2], '%Y-%m-%d') if len(sys.argv) > 2 else datetime.now()

# ---------- styles ----------
styles = getSampleStyleSheet()
NAVY = colors.HexColor('#305496')
LIGHTBLUE = colors.HexColor('#D9E2F3')
RED = colors.HexColor('#C00000')
ORANGE = colors.HexColor('#ED7D31')
GREY = colors.HexColor('#595959')

title_style = ParagraphStyle('TitleX', parent=styles['Title'], textColor=NAVY, fontSize=20, spaceAfter=4)
sub_style = ParagraphStyle('SubX', parent=styles['Normal'], textColor=GREY, fontSize=10, spaceAfter=14)
h2 = ParagraphStyle('H2X', parent=styles['Heading2'], textColor=NAVY, fontSize=13, spaceBefore=16, spaceAfter=6)
h3 = ParagraphStyle('H3X', parent=styles['Heading3'], textColor=colors.HexColor('#1F3864'), fontSize=11, spaceBefore=10, spaceAfter=4)
body = ParagraphStyle('BodyX', parent=styles['Normal'], fontSize=9.5, leading=13)
small = ParagraphStyle('SmallX', parent=styles['Normal'], fontSize=8.5, leading=11, textColor=GREY)
cell = ParagraphStyle('CellX', parent=styles['Normal'], fontSize=8.5, leading=11)
cell_b = ParagraphStyle('CellBX', parent=styles['Normal'], fontSize=8.5, leading=11, fontName='Helvetica-Bold')

def table_style(header_bg=NAVY, zebra=True):
    ts = [
        ('BACKGROUND', (0, 0), (-1, 0), header_bg),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 8.5),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#BFBFBF')),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('LEFTPADDING', (0, 0), (-1, -1), 4),
        ('RIGHTPADDING', (0, 0), (-1, -1), 4),
        ('TOPPADDING', (0, 0), (-1, -1), 3),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
    ]
    if zebra:
        ts.append(('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#F2F2F2')]))
    return TableStyle(ts)

def P(text, style=cell):
    return Paragraph(str(text) if text is not None else '', style)

# ---------- load data ----------
def load_sheet(name, header_row=4):
    ws = wb[name]
    header = [ws.cell(row=header_row, column=c).value for c in range(1, ws.max_column + 1)]
    rows = []
    for r in range(header_row + 1, ws.max_row + 1):
        row = [ws.cell(row=r, column=c).value for c in range(1, ws.max_column + 1)]
        if row[0] is None:
            continue
        rows.append(dict(zip(header, row)))
    return rows

actions = load_sheet('Action Tracker')
warnings_all = [w for w in load_sheet('Warnings') if w.get('Employee Name')]

# category counts from the match sheet (recomputed quickly, robust to header offset)
ws_match = wb['MOL vs Employee Match']
cat_counts = {}
header_row_idx = None
for r in range(1, ws_match.max_row + 1):
    if ws_match.cell(row=r, column=1).value == 'Salon' and ws_match.cell(row=r, column=8).value == 'Status / Flag':
        header_row_idx = r
        break
total_staff = 0
if header_row_idx:
    r = header_row_idx + 1
    while ws_match.cell(row=r, column=1).value is not None:
        cat = ws_match.cell(row=r, column=8).value
        cat_counts[cat] = cat_counts.get(cat, 0) + 1
        total_staff += 1
        r += 1

ws_dupsheet = wb['Duplicate Employees']
dup_people = []
r = 5
while r < 200:
    col_a = ws_dupsheet.cell(row=r, column=1).value
    if isinstance(col_a, str) and 'Similar Names' in col_a:
        break  # stop before the "similar but different people" appendix
    name = ws_dupsheet.cell(row=r, column=3).value
    passport = ws_dupsheet.cell(row=r, column=2).value
    salon = ws_dupsheet.cell(row=r, column=4).value
    if name:
        dup_people.append((name, salon, passport))
    r += 1

def in_period(d, start, end):
    if not isinstance(d, datetime):
        return False
    return start <= d <= end

PRIORITY_COLOR = {'Critical': RED, 'Urgent': ORANGE, 'Normal': colors.HexColor('#BFBFBF')}

def build_story(period_label, start, end, is_monthly):
    story = []
    story.append(Paragraph('Comprehensive Salons &amp; Staff Register', sub_style))
    story.append(Paragraph(f"{'Monthly' if is_monthly else 'Weekly'} HR Report to Management", title_style))
    story.append(Paragraph(f"Period: {period_label}   |   Prepared: {TODAY.strftime('%d/%m/%Y')}   |   Prepared by: HR (Diana Razouk)", sub_style))
    story.append(HRFlowable(width='100%', color=NAVY, thickness=1))

    # ---- 1. Headcount / compliance snapshot ----
    story.append(Paragraph('1. Staff & Work-Permit Compliance Snapshot', h2))
    data = [['Status', 'Count', 'What it means']]
    legend = {
        'MOHRE Confirmed': 'Matches the official government work-permit list.',
        'MOHRE Mismatch - Needs Review': 'Discrepancy vs. the official MOHRE list - needs HR review.',
        'Own Visa - No Salon Work Permit': 'Works at the salon on independent/own residency - no company permit.',
        'Undocumented - No Permit (Manager Confirmed)': 'Working with NO permit at all - urgent legal risk.',
        'Absconded': 'Flagged as absconded / left with no notice.',
        'Terminated': 'Employment ended (termination) - not absconding.',
        'Left Company': 'Confirmed left the company by the salon manager.',
        'Needs Cancellation - Not Actually Working': 'Owner confirmed no real employment - permit needs MOHRE cancellation.',
        'No Work Permit - Salon Not Yet MOHRE-Verified': "No permit on file; salon's official list not yet received.",
        'Awaiting Official MOHRE List': "Official MOHRE list not yet received for this salon.",
    }
    order = ['MOHRE Confirmed', 'MOHRE Mismatch - Needs Review', 'Own Visa - No Salon Work Permit',
             'Undocumented - No Permit (Manager Confirmed)', 'Absconded', 'Terminated', 'Left Company',
             'Needs Cancellation - Not Actually Working',
             'No Work Permit - Salon Not Yet MOHRE-Verified', 'Awaiting Official MOHRE List']
    for cat in order:
        if cat_counts.get(cat):
            data.append([P(cat, cell_b), P(cat_counts[cat]), P(legend.get(cat, ''), small)])
    data.append([P('Total staff on register', cell_b), P(total_staff, cell_b), ''])
    t = Table(data, colWidths=[5.5 * cm, 1.8 * cm, 9.2 * cm])
    t.setStyle(table_style())
    story.append(t)

    # ---- 2. Duplicate work permits ----
    story.append(Paragraph('2. Duplicate Work Permits (Same Person, Multiple Salons)', h2))
    story.append(Paragraph(
        f'{len(set(n for n, s, p in dup_people))} people were found registered under more than one salon / work permit card. '
        'Only one active work permit per person is normally valid under UAE labour law - each should be reviewed and the extra permit(s) cancelled.',
        body))
    data = [['Name', 'Passport No.', 'Salon(s)']]
    seen = {}
    for name, salon, passport in dup_people:
        seen.setdefault((name, passport), []).append(salon)
    for (name, passport), salons in seen.items():
        data.append([P(name), P(passport), P(', '.join(salons))])
    t = Table(data, colWidths=[6 * cm, 3.5 * cm, 7 * cm])
    t.setStyle(table_style(header_bg=RED))
    story.append(t)

    # ---- 3. Warnings issued & actions taken ----
    period_warnings = [w for w in warnings_all if in_period(w.get('Date of Warning'), start, end)]
    all_relevant_warnings = warnings_all if is_monthly else period_warnings
    story.append(Paragraph('3. Staff Warnings Issued & Action Taken' + ('' if is_monthly else ' (this week)'), h2))
    if all_relevant_warnings:
        data = [['Date', 'Salon', 'Employee', 'Level', 'Reason', 'Deduction', 'Ack.']]
        for w in sorted(all_relevant_warnings, key=lambda x: x.get('Date of Warning') or TODAY):
            d = w.get('Date of Warning')
            d = d.strftime('%d/%m/%Y') if isinstance(d, datetime) else ''
            ded = w.get('Salary Deduction (AED)')
            ded = f"AED {ded:,.0f}" if isinstance(ded, (int, float)) else '-'
            ack = w.get('Acknowledged (Y/N)') or 'Pending'
            if isinstance(ack, datetime):
                ack = ack.strftime('%d/%m/%Y')
            data.append([P(d), P(w.get('Salon')), P(w.get('Employee Name')), P(w.get('Warning Level')),
                         P((w.get('Reason / Incident') or '')[:110], small), P(ded), P(ack)])
        t = Table(data, colWidths=[1.8 * cm, 2.6 * cm, 3.2 * cm, 2.2 * cm, 5.5 * cm, 2 * cm, 1.9 * cm])
        t.setStyle(table_style())
        story.append(t)
        notes = [w.get('Notes') for w in all_relevant_warnings if w.get('Notes') and ('Termina' in str(w.get('Notes')) or 'settlement' in str(w.get('Notes')).lower())]
        for n in notes:
            story.append(Spacer(1, 4))
            story.append(Paragraph(f"Note: {n}", small))
    else:
        story.append(Paragraph('No new warnings issued in this period.', body))

    # ---- 4. Action tracker: open, in progress, resolved ----
    story.append(Paragraph('4. Action Tracker - Follow-Up Items', h2))
    period_actions = [a for a in actions if in_period(a.get('Date Added'), start, end)]
    urgent_open = [a for a in actions if a.get('Priority') in ('Critical', 'Urgent') and a.get('Status') != 'Resolved']

    if not is_monthly:
        story.append(Paragraph(f'New items logged this week: {len(period_actions)}', h3))
        src = period_actions if period_actions else []
    else:
        story.append(Paragraph(f'All open/tracked items this month: {len(actions)} total, '
                                f'{sum(1 for a in actions if a.get("Status")=="Resolved")} resolved, '
                                f'{sum(1 for a in actions if a.get("Status")!="Resolved")} still open.', h3))
        src = actions

    if src:
        data = [['#', 'Date', 'Category', 'Salon', 'Subject', 'Description', 'Priority', 'Status']]
        for a in src:
            d = a.get('Date Added')
            d = d.strftime('%d/%m/%Y') if isinstance(d, datetime) else ''
            data.append([P(a.get('#')), P(d), P(a.get('Category')), P(a.get('Salon')), P(a.get('Employee / Subject')),
                         P((a.get('Description') or '')[:130], small), P(a.get('Priority')), P(a.get('Status'))])
        t = Table(data, colWidths=[0.8 * cm, 1.8 * cm, 2.3 * cm, 2.3 * cm, 2.6 * cm, 5.8 * cm, 1.8 * cm, 1.9 * cm])
        ts = table_style()
        for i, a in enumerate(src, start=1):
            pc = PRIORITY_COLOR.get(a.get('Priority'))
            if pc:
                ts.add('TEXTCOLOR', (6, i), (6, i), pc)
                ts.add('FONTNAME', (6, i), (6, i), 'Helvetica-Bold')
        t.setStyle(ts)
        story.append(t)

    # ---- 5. Urgent / needs follow-up now ----
    story.append(Paragraph('5. URGENT - Needs Follow-Up Now', h2))
    if urgent_open:
        data = [['Priority', 'Salon', 'Subject', 'What is needed', 'Assigned To']]
        for a in sorted(urgent_open, key=lambda x: 0 if x.get('Priority') == 'Critical' else 1):
            data.append([P(a.get('Priority'), cell_b), P(a.get('Salon')), P(a.get('Employee / Subject')),
                         P((a.get('Description') or '')[:130], small), P(a.get('Assigned To'))])
        t = Table(data, colWidths=[2 * cm, 2.6 * cm, 3.2 * cm, 6.8 * cm, 2.4 * cm])
        ts = table_style(header_bg=RED)
        for i, a in enumerate(sorted(urgent_open, key=lambda x: 0 if x.get('Priority') == 'Critical' else 1), start=1):
            pc = PRIORITY_COLOR.get(a.get('Priority'))
            if pc:
                ts.add('BACKGROUND', (0, i), (0, i), colors.HexColor('#FCE4D6') if a.get('Priority') == 'Urgent' else colors.HexColor('#FFC7CE'))
        t.setStyle(ts)
        story.append(t)
        story.append(Spacer(1, 6))
        story.append(Paragraph('These items are open, Critical/Urgent priority, and have not been marked Resolved - please review and confirm next steps.', small))
    else:
        story.append(Paragraph('No open Critical/Urgent items at this time.', body))

    story.append(Spacer(1, 16))
    story.append(HRFlowable(width='100%', color=colors.HexColor('#BFBFBF'), thickness=0.5))
    story.append(Paragraph('Generated automatically from the Comprehensive Salons & Staff Register. '
                            'Full detail available in the "Action Tracker", "Warnings", "MOL vs Employee Match" and '
                            '"Duplicate Employees" tabs of the register.', small))
    return story

def make_pdf(path, period_label, start, end, is_monthly):
    doc = SimpleDocTemplate(path, pagesize=A4, topMargin=1.6 * cm, bottomMargin=1.6 * cm,
                             leftMargin=1.4 * cm, rightMargin=1.4 * cm,
                             title=f"HR {'Monthly' if is_monthly else 'Weekly'} Report - {period_label}")
    story = build_story(period_label, start, end, is_monthly)
    doc.build(story)
    print('Wrote', path)

week_start = TODAY - timedelta(days=6)
make_pdf(f"Weekly_HR_Report_{week_start.strftime('%Y-%m-%d')}_to_{TODAY.strftime('%Y-%m-%d')}.pdf",
          f"{week_start.strftime('%d %b %Y')} - {TODAY.strftime('%d %b %Y')}", week_start, TODAY, False)

month_start = TODAY.replace(day=1)
make_pdf(f"Monthly_HR_Report_{TODAY.strftime('%B_%Y')}.pdf",
          f"{month_start.strftime('%B %Y')} (to {TODAY.strftime('%d %b %Y')})", month_start, TODAY, True)
