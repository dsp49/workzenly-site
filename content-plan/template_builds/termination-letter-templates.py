#!/usr/bin/env python3
"""Build the downloads for /templates/termination-letter-templates/.

Run from the site root:  python content-plan/template_builds/termination-letter-templates.py
Files go to assets/templates/termination-letter-templates/. Letters use [brackets] for the parts to fill in;
the filled-in examples live in the page previews.
"""
import sys
import pathlib

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent.parent))
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
import tplkit
# tplkit's fixed .foot overlaps the last text line when a page is full; print the footer in the
# bottom page margin instead (Chrome page-margin boxes). Local to this build; tplkit is unchanged.
tplkit._CSS += (".foot { display: none; }\n@page { margin-bottom: 18mm; @bottom-left { content: "
                "\"Free template from workzenly.com\"; font-family: Arial, sans-serif; font-size: 8pt; color: #888; } }\n")
from tplkit import out, xlsx_sheet, docx_doc, pdf_doc

d = out("termination-letter-templates")
thin = Side(style="thin", color="C9C9C3")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)

RECEIPT = [
    ("h", "Acknowledgment of receipt"),
    ("p", "I received a copy of this letter. My signature confirms receipt only, not agreement."),
    ("sign", ["Employee signature", "Company representative"]),
]


def letter(name, title, text, extra=None):
    blocks = [("letter", text)] + (extra if extra is not None else RECEIPT)
    docx_doc(d / f"{name}.docx", title, blocks)
    pdf_doc(d / f"{name}.pdf", title, blocks)


FINAL = """Final pay and benefits
- Final paycheck: You will receive your final pay, including all hours worked through [termination date][ and accrued, unused vacation or PTO as required by our policy or state law], on [date and method].
- Health insurance: Your coverage ends on [date]. [If the company has 20 or more employees or your state has a continuation law:] You will receive a separate notice explaining how to continue coverage under COBRA[ or state continuation coverage].
- 401(k) or retirement plan: [Plan administrator] will send information about your account options.
- Other benefits: [Life insurance, FSA or HSA, commissions, bonus: state what happens to each.]

Company property
Please return [keys, badge, laptop, phone, uniforms, credit card] by [date]. [Explain how: drop off or prepaid shipping label.]

[State notices, if required: California employers attach the EDD pamphlet "For Your Benefit" (DE 2320). New York employers state the exact date benefits end, within five working days of termination.]

If you have questions about your final pay or benefits, contact [Name] at [phone] or [email]."""

# 1. Simple termination letter (at-will, no reason detail) -------------------------------------
letter("simple-termination-letter", "Termination letter", f"""[Company name]
[Address]

[Date]

[Employee full name]
[Employee address]

Dear [Employee first name],

This letter confirms that your employment with [Company name] as [Job title] ends effective [termination date].

{FINAL}

We thank you for your work with us and wish you the best.

Sincerely,

[Name]
[Title]""")

# 2. Termination for performance -------------------------------------------------------------
letter("termination-for-performance-letter", "Termination letter: performance", f"""[Company name]
[Address]

[Date]

PERSONAL AND CONFIDENTIAL

Dear [Employee first name],

This letter confirms our conversation today. Your employment with [Company name] as [Job title] ends effective [termination date].

The reason is that your job performance has not met the standards we discussed with you. In particular:
- [Specific standard not met, with facts: e.g., order accuracy averaged 91% from date to date against a required 98%]
- [Second issue, with facts]

We discussed these issues with you on [dates], including [a written warning on date / a performance improvement plan from date to date]. Despite that support, the required improvement did not happen.

{FINAL}

Sincerely,

[Name]
[Title]""")

# 3. Termination for misconduct or policy violation ---------------------------------------------
letter("termination-for-misconduct-letter", "Termination letter: misconduct or policy violation", f"""[Company name]
[Address]

[Date]

PERSONAL AND CONFIDENTIAL

Dear [Employee first name],

This letter confirms that your employment with [Company name] as [Job title] is terminated effective immediately, [date].

On [date], [describe what happened in one or two factual sentences, with no opinions or labels]. This violates our [policy name] in the [employee handbook section], which you acknowledged on [date].

[We reviewed (video, records, written statements) and met with you on (date) to hear your account before making this decision.]

{FINAL}

Sincerely,

[Name]
[Title]""")

# 4. Job abandonment / no call, no show --------------------------------------------------------
letter("job-abandonment-termination-letter", "Termination letter: job abandonment (no call, no show)", f"""[Company name]
[Address]

[Date]

Sent by [mail and email / certified mail] to [address and email on file]

Dear [Employee first name],

You were scheduled to work on [date], [date] and [date], and you did not report to work or contact your manager on any of those days. We tried to reach you by [phone on date, text on date, email on date] and did not receive a response.

Our attendance policy, in the [handbook section] you acknowledged on [date], states that [__] consecutive scheduled shifts missed without notice are treated as a voluntary resignation. As a result, your employment as [Job title] ended effective [date].

If you were unable to contact us because of an emergency, [a medical reason or other leave you may be entitled to], please contact [Name] at [phone] by [date] so we can review the situation.

{FINAL}

Sincerely,

[Name]
[Title]""")

# 5. Layoff / position elimination -----------------------------------------------------------
letter("layoff-position-elimination-letter", "Layoff letter: position elimination", f"""[Company name]
[Address]

[Date]

Dear [Employee first name],

We are sorry to tell you that, because of [reason, such as the closing of the location / lower sales / a reorganization], your position as [Job title] is being eliminated. Your last day of work will be [date]. This decision is not a reflection of your performance.

[Notice period: You will be paid for your regular hours through (date) whether or not you are asked to work.]
[Severance: The company will offer you (amount or formula) in severance pay if you sign the enclosed separation agreement by (date). You have (__) days to review it, and we encourage you to talk to a lawyer before signing.]
[Recall: If a similar position opens within (__) months, we will contact you first.]

{FINAL}

You may be eligible for unemployment benefits. You can apply through [state agency name and website].

[Name], [title], is available to answer your questions at [phone] or [email]. Thank you for your work with us.

Sincerely,

[Name]
[Title]""")

# 6. Probationary period termination ----------------------------------------------------------
letter("probationary-termination-letter", "Termination letter: during the introductory period", f"""[Company name]
[Address]

[Date]

Dear [Employee first name],

Thank you for meeting with me today. As discussed, your employment as [Job title] will end effective [date], during your [90-day] introductory period.

[Optional, one sentence: The role requires (skill or standard), and over the past (weeks) it has not been the right fit.]

Employment with [Company name] is at will, both during and after the introductory period.

{FINAL}

We appreciate the effort you put in and wish you success in your next role.

Sincerely,

[Name]
[Title]""")

# 7. End of temporary or seasonal assignment ---------------------------------------------------
letter("end-of-temporary-assignment-letter", "End of seasonal or temporary assignment letter", f"""[Company name]
[Address]

[Date]

Dear [Employee first name],

Thank you for working with us this [season or project] as [Job title]. As described in your offer letter, your seasonal assignment ends on [date], which will be your last scheduled shift.

{FINAL}

We would be glad to have you back. [If you are interested in next season / open regular positions, reply to this letter or apply at (link) by (date), and we will consider you first.]

Thank you again,

[Name]
[Title]""")

# 8. Notice of separation form (California and New York ready) -------------------------------
sep_blocks = [
    ("p", "Give this notice to every employee who is discharged, laid off or placed on leave, on or before the last "
          "day of work, and keep a copy in the personnel file."),
    ("h", "Employee and employer"),
    ("fields", ["Employee name and job title", "Employer name", "Employer address and phone"]),
    ("h", "Change in employment"),
    ("checks", ["Discharge", "Layoff", "Leave of absence", "Change in status (explain)"]),
    ("fields", ["Last day worked", "Effective date of separation (exact date)", "Expected return date (leave only)"]),
    ("h", "Benefits"),
    ("fields", ["Health, dental and vision end on (exact date)", "Other benefits end on (exact date)"]),
    ("checks", ["COBRA or state continuation notice will be sent by the plan administrator", "No group health plan"]),
    ("h", "Final pay"),
    ("fields", ["Final paycheck date and method", "Accrued vacation or PTO paid (hours)"]),
    ("h", "Enclosures"),
    ("checks", ["California: EDD For Your Benefit pamphlet (DE 2320)", "State unemployment notice", "COBRA information"]),
    ("checks", ["Separation or severance agreement", "Return of company property list"]),
    ("sign", ["Employer representative", "Employee (acknowledges receipt only)"]),
]
docx_doc(d / "notice-of-separation-form.docx", "Notice of separation", sep_blocks)
pdf_doc(d / "notice-of-separation-form.pdf", "Notice of separation", sep_blocks)

# 9. Termination checklist and final pay worksheet (Excel) -------------------------------------
chk_headers = ["Step", "Owner", "Due", "Done (Y/N)", "Notes"]
chk_rows = [
    ["Confirm reason, documents and prior warnings are in the file", "Manager + HR", "Before meeting", "", ""],
    ["Check state final paycheck deadline and PTO payout rule", "HR / payroll", "Before meeting", "", ""],
    ["Prepare termination letter and state notices", "HR", "Before meeting", "", ""],
    ["Hold the termination meeting with a witness", "Manager", "Day of", "", ""],
    ["Turn off building, system, email and app access", "IT / owner", "Day of", "", ""],
    ["Collect keys, badge, devices, uniforms, cards", "Manager", "Day of", "", ""],
    ["Run final paycheck (see Final pay sheet)", "Payroll", "Per state law", "", ""],
    ["Notify the health plan administrator (COBRA)", "HR", "Within 30 days", "", ""],
    ["Remove employee from future schedules", "Manager", "Day of", "", ""],
    ["Update personnel file and save signed copies", "HR", "Within 1 week", "", ""],
    ["Respond to unemployment claim notice, if received", "HR", "By agency deadline", "", ""],
]
chk_path = d / "termination-checklist-and-final-pay-worksheet.xlsx"
xlsx_sheet(chk_path, title="Employee termination checklist",
           subtitle="Employee: ____________   Last day worked: ________   Manager: ____________",
           headers=chk_headers, rows=chk_rows, blank_rows=4, widths=[52, 16, 18, 12, 30],
           sheet_name="Checklist", landscape=True,
           notes="Final paycheck deadlines and PTO payout rules vary by state. Check your state labor department.")

# Final pay worksheet on a second sheet with formulas
wb = load_workbook(chk_path)
ws = wb.create_sheet("Final pay")
ws["A1"] = "Final pay worksheet"
ws["A1"].font = Font(bold=True, size=16)
ws["A2"] = "Enter the yellow cells. Gross pay before taxes and deductions."
ws["A2"].font = Font(size=11, color="444444")
rows = [
    ("Item", "Hours", "Rate", "Amount"),
    ("Regular hours since last paycheck", "", "", "=B5*C5"),
    ("Overtime hours (1.5x)", "", "=C5*1.5", "=B6*C6"),
    ("Accrued PTO or vacation to pay out", "", "=C5", "=B7*C7"),
    ("Commission or bonus earned", "", "", ""),
    ("Expense reimbursements owed", "", "", ""),
    ("Total gross final pay", "", "", "=SUM(D5:D9)"),
]
for i, row in enumerate(rows):
    r = 4 + i
    for c, v in enumerate(row, 1):
        cell = ws.cell(row=r, column=c, value=v)
        cell.border = BOX
        cell.alignment = Alignment(vertical="center", wrap_text=True)
        if i == 0:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="161616")
        elif isinstance(v, str) and v.startswith("="):
            cell.fill = PatternFill("solid", fgColor="F2FFD9")
        elif c > 1 and v == "" and not (i == 6):
            cell.fill = PatternFill("solid", fgColor="FFF6CC")
    if i > 0:
        for c in (3, 4):
            ws.cell(row=r, column=c).number_format = '"$"#,##0.00'
ws.cell(row=10, column=1).font = Font(bold=True)
ws.cell(row=10, column=4).font = Font(bold=True)
ws["A12"] = "Final paycheck due date (check state law):"
ws["A13"] = "Paid on:"
ws["A14"] = "Paid by (check, direct deposit, mail):"
for r in (12, 13, 14):
    ws.cell(row=r, column=1).font = Font(bold=True)
    ws.cell(row=r, column=2).fill = PatternFill("solid", fgColor="FFF6CC")
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
ws["A16"] = "Free template from workzenly.com"
ws["A16"].font = Font(size=9, color="888888")
for col, w in zip("ABCD", (40, 12, 12, 16)):
    ws.column_dimensions[col].width = w
wb.save(chk_path)

pdf_doc(d / "termination-checklist-and-final-pay-worksheet.pdf", "Employee termination checklist", [
    ("fields", ["Employee", "Last day worked", "Manager"]),
    ("table", chk_headers, chk_rows),
    ("h", "Final pay worksheet"),
    ("table", ["Item", "Hours", "Rate", "Amount"], [
        ["Regular hours since last paycheck", "", "", ""],
        ["Overtime hours (1.5x)", "", "", ""],
        ["Accrued PTO or vacation to pay out", "", "", ""],
        ["Commission or bonus earned", "", "", ""],
        ["Expense reimbursements owed", "", "", ""],
        ["Total gross final pay", "", "", ""]]),
    ("fields", ["Final paycheck due date (check state law)", "Paid on", "Paid by"]),
    ("p", "Final paycheck deadlines and PTO payout rules vary by state. Check your state labor department."),
])

for f in sorted(d.iterdir()):
    print(f"{f.name:55s} {f.stat().st_size:>8,d} bytes")
