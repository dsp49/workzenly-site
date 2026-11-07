#!/usr/bin/env python3
"""Build the downloads for /templates/onboarding-checklist-templates/.

Run from the site root:  python content-plan/template_builds/onboarding-checklist-templates.py
Files go to assets/templates/onboarding-checklist-templates/. Files are blank and ready to fill in
(the master checklist and the 30-60-90 plan come with editable sample tasks);
the filled-in examples live in the page previews.
"""
import sys
import pathlib

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent.parent))
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.worksheet.datavalidation import DataValidation
from tplkit import out, xlsx_sheet, docx_doc, pdf_table, pdf_doc

d = out("onboarding-checklist-templates")
thin = Side(style="thin", color="C9C9C3")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)


def both(name, title, blocks):
    docx_doc(d / f"{name}.docx", title, blocks)
    pdf_doc(d / f"{name}.pdf", title, blocks)


def box(c, fill=None, bold=False):
    c.border = BOX
    c.alignment = Alignment(vertical="center", wrap_text=True)
    if fill:
        c.fill = PatternFill("solid", fgColor=fill)
    if bold:
        c.font = Font(bold=True)


# 1. Master new hire onboarding checklist (Excel) -----------------------------------------------
m_tasks = [
    ("Before day 1", "Send signed offer letter and start-date details", "Hiring manager", -10),
    ("Before day 1", "Send W-4, state withholding and direct deposit forms", "HR / payroll", -7),
    ("Before day 1", "Invite to complete I-9 Section 1 (no earlier than offer acceptance)", "HR", -5),
    ("Before day 1", "Add to schedule and time clock; create app login", "Manager", -3),
    ("Before day 1", "Order uniform, keys, badge or equipment", "Manager", -3),
    ("Before day 1", "Assign a buddy or trainer", "Manager", -2),
    ("Day 1", "Welcome, tour, introductions", "Manager", 0),
    ("Day 1", "Employee completes I-9 Section 1 (if not done)", "Employee", 0),
    ("Day 1", "Review handbook; collect signed acknowledgment", "HR", 0),
    ("Day 1", "Show how to clock in and out, breaks and call-out rules", "Manager", 0),
    ("Day 1", "Safety training required for the job", "Trainer", 0),
    ("Week 1", "Employer completes I-9 Section 2 (within 3 business days of start)", "HR", 3),
    ("Week 1", "Job training on core tasks", "Trainer", 4),
    ("Week 1", "First check-in: questions, schedule, pay date", "Manager", 4),
    ("Week 1", "Confirm first paycheck and pay stub access", "Payroll", 7),
    ("Week 3", "Report new hire to the state (federal limit 20 days; many states are shorter)", "HR / payroll", 14),
    ("Day 30", "30-day check-in", "Manager", 30),
    ("Day 60", "60-day check-in; benefits enrollment confirmed", "Manager / HR", 60),
    ("Day 90", "90-day review", "Manager", 90),
]
m_headers = ["Phase", "Task", "Owner", "Days from start", "Due date", "Done (Y/N)", "Date done", "Notes"]
m_path = d / "new-hire-onboarding-checklist.xlsx"
xlsx_sheet(m_path, title="New hire onboarding checklist",
           subtitle="Employee: ____________   Job title: ____________   Manager: ____________",
           headers=m_headers, rows=[[p, t, o, n, "", "", "", ""] for p, t, o, n in m_tasks], blank_rows=6,
           widths=[13, 58, 16, 12, 13, 11, 13, 30],
           formulas={"Due date": '=IF(OR($J$5="",D{r}=""),"",$J$5+D{r})'},
           sheet_name="Checklist", landscape=True,
           notes="Type the start date in J5. Due dates are calendar days from the start date, except I-9 Section 2, "
                 "which counts 3 weekdays. Check your state's new hire reporting deadline.")
wb = load_workbook(m_path)
ws = wb.active
ws.column_dimensions["I"].width = 22
ws.column_dimensions["J"].width = 14
cells = [("I4", "Summary", True), ("I5", "Start date", True), ("I6", "Tasks", True), ("I7", "Done", True),
         ("I8", "% complete", True), ("I9", "Overdue", True)]
for ref, val, b in cells:
    ws[ref].value = val
    box(ws[ref], bold=b)
ws["I4"].font = Font(bold=True, color="FFFFFF")
ws["I4"].fill = PatternFill("solid", fgColor="161616")
ws["J4"].fill = PatternFill("solid", fgColor="161616")
box(ws["J4"])
last = 4 + len(m_tasks) + 6
ws["J5"].number_format = "mm/dd/yyyy"
ws["J6"].value = f'=COUNTA(B5:B{last})'
ws["J7"].value = f'=COUNTIF(F5:F{last},"Y")'
ws["J8"].value = '=IF(J6=0,"",J7/J6)'
ws["J8"].number_format = "0%"
ws["J9"].value = f'=SUMPRODUCT((E5:E{last}<>"")*(E5:E{last}<TODAY())*(F5:F{last}<>"Y"))'
for r in range(5, 10):
    box(ws[f"J{r}"], fill="FFFFFF" if r == 5 else "F2FFD9")
for r in range(5, last + 1):
    ws[f"E{r}"].number_format = "mm/dd/yyyy"
    ws[f"G{r}"].number_format = "mm/dd/yyyy"
# I-9 Section 2 is due in business days, not calendar days
i9_row = 5 + next(i for i, t in enumerate(m_tasks) if "Section 2" in t[1])
ws[f"E{i9_row}"].value = '=IF($J$5="","",WORKDAY($J$5,3))'
ws[f"D{i9_row}"].value = "3 business"
dv = DataValidation(type="list", formula1='"Y,N"', allow_blank=True)
ws.add_data_validation(dv)
dv.add(f"F5:F{last}")
wb.save(m_path)
pdf_table(d / "new-hire-onboarding-checklist.pdf", "New hire onboarding checklist",
          ["Phase", "Task", "Owner", "Due date", "Done", "Notes"],
          [[p, t, o, "", "☐", ""] for p, t, o, n in m_tasks], blank_rows=3,
          subtitle="Employee: ____________   Start date: ________   Manager: ____________", landscape=True)

# 2. New hire paperwork checklist ---------------------------------------------------------------
both("new-hire-paperwork-checklist", "New hire paperwork checklist (US)", [
    ("fields", ["Employee name", "Start date (first day of work for pay)", "Work state", "HR contact"]),
    ("h", "Federal forms"),
    ("table", ["Form", "Deadline or rule", "Done", "Date"], [
        ["Form W-4 (current year version from irs.gov)", "Before the first payroll", "☐", ""],
        ["Form I-9, Section 1 (employee)", "No later than the first day of employment", "☐", ""],
        ["Form I-9, Section 2 (employer reviews documents)", "Within 3 business days of the first day", "☐", ""],
        ["E-Verify case (if enrolled or required by your state)", "After the I-9 is complete", "☐", ""],
    ]),
    ("h", "State and payroll"),
    ("table", ["Item", "Deadline or rule", "Done", "Date"], [
        ["State withholding certificate (states with income tax)", "Before the first payroll", "☐", ""],
        ["State new hire report", "Within 20 days of hire under federal law; many states require sooner", "☐", ""],
        ["State-required notices at hire (for example, pay rate notice)", "At hire; check your state", "☐", ""],
        ["Direct deposit authorization (if the employee chooses it)", "Before the first payroll", "☐", ""],
        ["Pay rate, pay schedule and overtime status confirmed in writing", "Before the first payroll", "☐", ""],
    ]),
    ("h", "Company documents"),
    ("table", ["Item", "Notes", "Done", "Date"], [
        ["Signed offer letter", "", "☐", ""],
        ["Employee handbook acknowledgment", "", "☐", ""],
        ["Emergency contact form", "", "☐", ""],
        ["Benefits enrollment or waiver", "", "☐", ""],
        ["Confidentiality or other agreements (if used)", "", "☐", ""],
    ]),
    ("p", "Keep each Form I-9 for 3 years after hire or 1 year after employment ends, whichever is later."),
    ("sign", ["Completed by"]),
])

# 3. First-day checklist for managers -----------------------------------------------------------
both("first-day-onboarding-checklist", "First-day onboarding checklist for managers", [
    ("fields", ["New hire", "Start date and shift", "Manager", "Buddy or trainer"]),
    ("h", "Before they arrive"),
    ("checks", ["Workstation, locker or truck ready", "Uniform, badge, keys ready", "Login and time clock set up",
                "Team told who is starting", "First week schedule posted"]),
    ("h", "First hour"),
    ("checks", ["Welcome and tour (restrooms, break room, exits)", "Introduce the team and the buddy",
                "Paperwork: I-9 Section 1, W-4, direct deposit", "Walk through the first week schedule"]),
    ("h", "During the shift"),
    ("checks", ["Show how to clock in and out and record breaks", "Explain call-out and time off requests",
                "Required safety training", "Shadow the buddy on core tasks", "Lunch with the team or buddy"]),
    ("h", "End of the day"),
    ("checks", ["Ask what was confusing", "Confirm tomorrow's shift", "Explain payday and how to view pay stubs"]),
    ("fields", ["Questions to follow up on"]),
    ("sign", ["Manager signature"]),
])

# 4. 30-60-90 day onboarding plan (Excel) -------------------------------------------------------
p_headers = ["Phase", "Goal", "How we will know", "Owner", "Target date", "Status", "Notes"]
p_rows = [
    ["Days 1-30: Learn", "", "", "", "", "", ""], ["Days 1-30: Learn", "", "", "", "", "", ""],
    ["Days 1-30: Learn", "", "", "", "", "", ""],
    ["Days 31-60: Contribute", "", "", "", "", "", ""], ["Days 31-60: Contribute", "", "", "", "", "", ""],
    ["Days 31-60: Contribute", "", "", "", "", "", ""],
    ["Days 61-90: Own", "", "", "", "", "", ""], ["Days 61-90: Own", "", "", "", "", "", ""],
    ["Days 61-90: Own", "", "", "", "", "", ""],
]
p_path = d / "30-60-90-day-onboarding-plan.xlsx"
xlsx_sheet(p_path, title="30-60-90 day onboarding plan",
           subtitle="Employee: ____________   Job title: ____________   Manager: ____________   Start date: ________",
           headers=p_headers, rows=p_rows, blank_rows=3, widths=[22, 40, 34, 14, 13, 14, 28],
           sheet_name="Plan", landscape=True,
           notes="Status options: Not started, On track, At risk, Done. Review the plan at the 30, 60 and 90-day "
                 "check-ins.")
wb = load_workbook(p_path)
ws = wb.active
dv = DataValidation(type="list", formula1='"Not started,On track,At risk,Done"', allow_blank=True)
ws.add_data_validation(dv)
dv.add("F5:F16")
for r in range(5, 17):
    ws[f"E{r}"].number_format = "mm/dd/yyyy"
wb.save(p_path)
pdf_table(d / "30-60-90-day-onboarding-plan.pdf", "30-60-90 day onboarding plan",
          ["Phase", "Goal", "How we will know", "Owner", "Target date", "Status"],
          [r[:6] for r in p_rows], blank_rows=2,
          subtitle="Employee: ____________   Job title: ____________   Start date: ________", landscape=True)

# 5. Hourly and shift worker onboarding checklist -----------------------------------------------
both("hourly-employee-onboarding-checklist", "Hourly employee onboarding checklist", [
    ("p", "For restaurants, retail stores, warehouses, clinics and field crews. Complete in the first week."),
    ("fields", ["Employee name", "Position", "Location", "Start date", "Hourly rate", "Manager"]),
    ("h", "Pay and time"),
    ("checks", ["Pay rate, payday and pay period explained", "Overtime rule explained (over 40 hours a week, plus any state rule)",
                "Clock in and out shown, including on the phone app or kiosk",
                "Meal and rest break rules for this state explained", "How to fix a missed punch",
                "Tip policy and tip pool explained (tipped jobs)"]),
    ("h", "Schedule"),
    ("checks", ["How and when schedules are posted", "How to request time off", "How to swap or drop a shift",
                "Call-out rule (who to contact and how early)"]),
    ("h", "Job and safety"),
    ("checks", ["Uniform and dress code", "Safety training for this job", "Food handler or other card (where required)",
                "Equipment training (POS, scanner, forklift awareness, etc.)", "Where to report an injury or concern"]),
    ("h", "People"),
    ("checks", ["Met the team and shift leads", "Buddy assigned for first five shifts", "Knows who to ask for help"]),
    ("sign", ["Employee signature", "Manager signature"]),
])

# 6. Remote employee onboarding checklist -------------------------------------------------------
both("remote-employee-onboarding-checklist", "Remote employee onboarding checklist", [
    ("fields", ["Employee name", "Job title", "Home state (for payroll tax)", "Start date", "Manager"]),
    ("h", "Before day 1"),
    ("checks", ["Laptop and accessories shipped with tracking", "Accounts created (email, chat, apps)",
                "I-9 plan confirmed: authorized representative or the DHS alternative procedure (E-Verify employers only)",
                "Payroll set up for the employee's work state", "Welcome email with day 1 agenda"]),
    ("h", "Day 1"),
    ("checks", ["Video welcome with manager", "IT setup call and security training", "Team introduction call",
                "Working hours, time tracking and response time expectations", "Handbook and remote work policy acknowledged"]),
    ("h", "Weeks 1 to 4"),
    ("checks", ["Daily 15-minute check-in for week 1", "Buddy calls twice a week", "First project assigned",
                "Home office reimbursement rules explained (if any)", "30-day check-in scheduled"]),
    ("fields", ["Equipment serial numbers"]),
    ("sign", ["Employee signature", "Manager signature"]),
])

# 7. Onboarding tracker for HR (Excel) ----------------------------------------------------------
t_headers = ["Employee", "Position", "Location", "Start date", "W-4 received", "I-9 Section 1", "I-9 Section 2 due",
             "I-9 Section 2 done", "State report due (20-day max)", "State report filed", "Handbook signed",
             "90-day review", "Status"]
t_path = d / "onboarding-tracker.xlsx"
xlsx_sheet(t_path, title="Onboarding tracker",
           subtitle="One row per new hire. Enter dates; due dates and status fill in automatically.",
           headers=t_headers, rows=[], blank_rows=30,
           widths=[20, 16, 14, 12, 11, 11, 13, 12, 15, 12, 11, 12, 18],
           formulas={"I-9 Section 2 due": '=IF(D{r}="","",WORKDAY(D{r},3))',
                     "State report due (20-day max)": '=IF(D{r}="","",D{r}+20)',
                     "90-day review": '=IF(D{r}="","",D{r}+90)',
                     "Status": '=IF(D{r}="","",IF(AND(E{r}<>"",H{r}<>"",J{r}<>"",K{r}<>""),"Paperwork complete",'
                               'IF(AND(H{r}="",TODAY()>G{r}),"I-9 overdue",IF(AND(J{r}="",TODAY()>I{r}),'
                               '"Report overdue","In progress"))))'},
           sheet_name="Tracker", landscape=True,
           notes="I-9 Section 2 due counts 3 weekdays after the start date (WORKDAY skips Saturdays and Sundays). "
                 "The state report due date uses the federal 20-day limit; many states require sooner, so change "
                 "+20 to your state's deadline.")
wb = load_workbook(t_path)
ws = wb.active
for r in range(5, 35):
    for col in "DEFGHIJKL":
        ws[f"{col}{r}"].number_format = "mm/dd/yyyy"
wb.save(t_path)
pdf_table(d / "onboarding-tracker.pdf", "Onboarding tracker",
          ["Employee", "Position", "Start date", "W-4", "I-9 Sec. 1", "I-9 Sec. 2 due", "I-9 Sec. 2 done",
           "State report due", "Report filed", "Handbook", "90-day review"], blank_rows=16,
          subtitle="I-9 Section 2 due within 3 business days of start. State new hire report within 20 days "
                   "of hire (sooner in many states).", landscape=True)

print("built", sorted(p.name for p in d.iterdir()))
