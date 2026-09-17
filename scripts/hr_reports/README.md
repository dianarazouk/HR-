# HR Weekly / Monthly Reports

Generates the Weekly and Monthly HR PDF reports (compliance snapshot, duplicate work
permits, warnings issued, action tracker, urgent follow-ups) from the master
"Comprehensive Salons & Staff Register" workbook.

## Setup

Google Drive holds the live data and the document library:

- **HR - Salon Staff Documents** (root folder) - keep the latest master register
  `.xlsx` here. Each report run downloads the most recently modified `.xlsx` in this
  folder.
- **ALEKSANDRA / UNIQUE YOU / NISANTASI / THE LAB - Branch / THE LAB - HQ / HAIR TAG** -
  one folder per salon for employee documents (passport, visa, EID, contract scans).
  Name files `SURNAME FIRSTNAME - DocType.ext` (e.g. `HEMRAYEVA MAYSA - Passport.pdf`).
- **Reports - Weekly & Monthly** - generated PDF reports land here.

`build_review_tabs.py` adds the "Duplicate Employees" and "MOL vs Employee Match"
review tabs to the register (matches the internal salon registers against the
"MOHRE Verification" tab, flags duplicate work permits, own-visa/no-permit staff,
and MOHRE mismatches). Run this first if the register doesn't already have those
tabs, or after adding a new official MOHRE list.

`generate_reports.py` reads the register (with those tabs present) and produces:

- `Weekly_HR_Report_<start>_to_<end>.pdf` - warnings issued and actions logged in
  the last 7 days, plus all open Critical/Urgent items.
- `Monthly_HR_Report_<Month_Year>.pdf` - full month snapshot: compliance counts,
  duplicates, all warnings, full action tracker status, urgent items.

```bash
pip install -r requirements.txt
python3 build_review_tabs.py            # only if review tabs are missing/stale
python3 generate_reports.py <register.xlsx> <YYYY-MM-DD>   # date = "today" for the report
```

Both PDFs are then uploaded to the **Reports - Weekly & Monthly** Drive folder, and
(once a Gmail/Outlook connector is authorized) emailed to management.

## Recipient

Reports and urgent alerts go to **Admin-HR Salon <hr.beautysalons@gmail.com>**
(also the Google Drive account these folders live in).
