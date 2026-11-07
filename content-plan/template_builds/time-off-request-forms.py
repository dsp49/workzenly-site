#!/usr/bin/env python3
"""Build the downloads for /templates/time-off-request-forms/.

Run from the site root:  python content-plan/template_builds/time-off-request-forms.py
Files go to assets/templates/time-off-request-forms/. Files are blank and ready to fill in;
the filled-in examples live in the page previews.
"""
import sys
import pathlib

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent.parent))
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from tplkit import out, xlsx_sheet, docx_doc, pdf_table, pdf_doc

d = out("time-off-request-forms")
LEAVE_TYPES = ["Vacation", "Personal", "Sick", "Bereavement", "Jury duty", "Unpaid", "Other: ________"]
thin = Side(style="thin", color="C9C9C3")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)


def both(name, title, blocks):
    docx_doc(d / f"{name}.docx", title, blocks)
    pdf_doc(d / f"{name}.pdf", title, blocks)


def summary_cells(path, cells, col_label="H", col_value="I", start_row=4, label_width=30, value_width=14):
    """Add a small labeled summary box with formulas to the right of a tplkit sheet."""
    wb = load_workbook(path)
    ws = wb.active
    ws.column_dimensions[col_label].width = label_width
    ws.column_dimensions[col_value].width = value_width
    for i, (label, value) in enumerate(cells):
        r = start_row + i
        a = ws[f"{col_label}{r}"]
        b = ws[f"{col_value}{r}"]
        a.value, b.value = label, value
        a.font = Font(bold=True, color="FFFFFF" if i == 0 else "111111")
        if i == 0:
            a.fill = PatternFill("solid", fgColor="161616")
            b.fill = PatternFill("solid", fgColor="161616")
        else:
            b.fill = PatternFill("solid", fgColor="F2FFD9" if str(value).startswith("=") else "FFFFFF")
        for c in (a, b):
            c.border = BOX
            c.alignment = Alignment(vertical="center", wrap_text=True)
    wb.save(path)


# 1. Simple time off request form -------------------------------------------------------------
both("simple-time-off-request-form", "Time off request form", [
    ("p", "Fill in the top half and give this form to your manager. Submit vacation and other planned time off "
          "at least two weeks before the first day off, or as your company policy requires."),
    ("h", "Employee"),
    ("fields", ["Employee name", "Employee ID", "Job title", "Department or location", "Manager", "Date of request"]),
    ("h", "Time off requested"),
    ("checks", LEAVE_TYPES),
    ("fields", ["First day off", "Last day off", "Date I return to work", "Total hours requested"]),
    ("fields", ["Comments (optional)"]),
    ("sign", ["Employee signature"]),
    ("h", "Manager use only"),
    ("checks", ["Approved", "Approved with changes", "Denied"]),
    ("fields", ["Changes or reason if denied", "Hours to record as paid", "Hours to record as unpaid"]),
    ("sign", ["Manager signature"]),
])

# 2. PTO request form with balance check (Excel) -----------------------------------------------
pto_headers = ["Date", "Day", "Leave type", "Hours requested", "Notes"]
pto_path = d / "pto-request-form-with-balance.xlsx"
xlsx_sheet(pto_path, title="PTO request form",
           subtitle="Employee: ____________   Manager: ____________   Date of request: ________",
           headers=pto_headers, rows=[], blank_rows=10, widths=[14, 10, 22, 16, 30],
           sheet_name="PTO request", landscape=False,
           notes="Enter one row per day off. Only rows marked PTO come out of the balance. "
                 "Paid holidays and unpaid days are recorded for the schedule but do not use PTO.")
# table header is row 4, data rows 5-14
summary_cells(pto_path, [
    ("Summary", ""),
    ("PTO balance before this request (hours)", ""),
    ("Total hours requested", "=SUM(D5:D14)"),
    ("Hours taken from PTO", '=SUMIF(C5:C14,"PTO",D5:D14)'),
    ("PTO balance after this request", "=I5-I7"),
    ("Approved? (Yes / No)", ""),
    ("Manager initials and date", ""),
], col_label="G", col_value="I")
wb = load_workbook(pto_path)
ws = wb.active
ws.column_dimensions["H"].width = 2
for r in range(4, 11):  # merge G:H so labels read as one cell
    ws.merge_cells(f"G{r}:H{r}")
wb.save(pto_path)
pdf_doc(d / "pto-request-form-with-balance.pdf", "PTO request form", [
    ("fields", ["Employee name", "Manager", "Date of request"]),
    ("table", pto_headers, [[""] * 5] * 8),
    ("h", "PTO balance"),
    ("fields", ["Balance before this request (hours)", "Hours taken from PTO", "Balance after this request (hours)"]),
    ("sign", ["Employee signature"]),
    ("h", "Manager use only"),
    ("checks", ["Approved", "Approved with changes", "Denied"]),
    ("sign", ["Manager signature"]),
])

# 3. Vacation request form with coverage plan --------------------------------------------------
both("vacation-request-form-with-coverage-plan", "Vacation request form", [
    ("p", "Use this form for vacations of three days or more. List who will handle your work so your manager can "
          "approve the dates and the coverage at the same time."),
    ("fields", ["Employee name", "Job title", "Manager", "Date of request"]),
    ("h", "Vacation dates"),
    ("fields", ["First day off", "Last day off", "Return to work", "Workdays off", "PTO hours requested"]),
    ("checks", ["Not reachable while away", "Reachable for emergencies only", "Checking email once a day"]),
    ("h", "Coverage plan"),
    ("table", ["Task or account", "Covered by", "Handoff notes"], [["", "", ""]] * 5),
    ("fields", ["Handoff meeting date", "Out-of-office message set (yes/no)"]),
    ("sign", ["Employee signature", "Covering coworker(s)"]),
    ("h", "Manager use only"),
    ("checks", ["Approved", "Approved with changes", "Denied"]),
    ("sign", ["Manager signature"]),
])

# 4. Partial-day time off request --------------------------------------------------------------
both("partial-day-time-off-request-form", "Partial-day time off request", [
    ("p", "For time off of less than a full shift: appointments, school events, a late start or leaving early."),
    ("fields", ["Employee name", "Position", "Manager", "Date of request"]),
    ("h", "Time requested"),
    ("fields", ["Date", "Scheduled shift (start to end)", "Leaving at", "Returning at (or not returning)",
                "Hours off"]),
    ("h", "How should these hours be paid?"),
    ("checks", ["Use PTO", "Unpaid", "Make up the hours in the same workweek (hourly employees)"]),
    ("fields", ["If making up hours: date and time"]),
    ("sign", ["Employee signature"]),
    ("h", "Manager use only"),
    ("checks", ["Approved", "Denied"]),
    ("sign", ["Manager signature"]),
])

# 5. Sick leave notice form --------------------------------------------------------------------
both("sick-leave-notice-form", "Sick leave notice", [
    ("p", "Use this form to record a sick day. Complete it when you call out or when you return. You do not need "
          "to describe your illness or give medical details."),
    ("fields", ["Employee name", "Position", "Location", "Manager"]),
    ("h", "Absence"),
    ("table", ["Date", "Scheduled shift", "Hours missed", "Paid sick hours used"], [["", "", "", ""]] * 4),
    ("h", "How you gave notice"),
    ("fields", ["Person notified", "Date and time notified"]),
    ("checks", ["Phone call", "Text", "Scheduling app", "Email", "In person"]),
    ("h", "Leave used for (optional)"),
    ("checks", ["My own health", "Care of a family member", "Other reason allowed under our sick leave policy"]),
    ("sign", ["Employee signature"]),
    ("h", "Payroll use only"),
    ("fields", ["Sick leave balance before", "Hours used", "Balance after"]),
    ("sign", ["Recorded by"]),
])

# 6. Multi-day leave request with day-by-day table (Excel) -------------------------------------
multi_headers = ["Date", "Day", "Scheduled hours", "Leave type", "Hours requested", "Paid?"]
multi_path = d / "multi-day-leave-request-form.xlsx"
xlsx_sheet(multi_path, title="Multi-day leave request",
           subtitle="Employee: ____________   Manager: ____________   Date of request: ________",
           headers=multi_headers, rows=[], blank_rows=14, widths=[14, 10, 16, 18, 16, 10],
           sheet_name="Leave request", landscape=True,
           notes="Leave type: PTO, Sick, Bereavement, Holiday, Jury duty or Unpaid. Totals update as you type.")
summary_cells(multi_path, [
    ("Totals", ""),
    ("Total hours requested", "=SUM(E5:E18)"),
    ("PTO hours", '=SUMIF(D5:D18,"PTO",E5:E18)'),
    ("Sick hours", '=SUMIF(D5:D18,"Sick",E5:E18)'),
    ("Bereavement hours", '=SUMIF(D5:D18,"Bereavement",E5:E18)'),
    ("Holiday hours", '=SUMIF(D5:D18,"Holiday",E5:E18)'),
    ("Jury duty hours", '=SUMIF(D5:D18,"Jury duty",E5:E18)'),
    ("Unpaid hours", '=SUMIF(D5:D18,"Unpaid",E5:E18)'),
    ("Approved? (Yes / No)", ""),
    ("Manager initials and date", ""),
])
pdf_doc(d / "multi-day-leave-request-form.pdf", "Multi-day leave request", [
    ("fields", ["Employee name", "Manager", "Date of request"]),
    ("p", "Leave type: PTO, Sick, Bereavement, Holiday, Jury duty or Unpaid. One row per scheduled day."),
    ("table", multi_headers, [[""] * 6] * 10),
    ("h", "Totals (hours)"),
    ("table", ["PTO", "Sick", "Bereavement", "Holiday", "Jury duty", "Unpaid", "All hours"], [[""] * 7]),
    ("sign", ["Employee signature"]),
    ("checks", ["Approved", "Approved with changes", "Denied"]),
    ("sign", ["Manager signature"]),
])

# 7. Time off request log for managers (Excel) ------------------------------------------------
log_headers = ["Req #", "Date received", "Employee", "Leave type", "First day", "Last day", "Hours",
               "Status", "Decided by", "Date decided", "Notes"]
log_path = d / "time-off-request-log.xlsx"
xlsx_sheet(log_path, title="Time off request log",
           subtitle="Team or location: ____________   Year: ______",
           headers=log_headers, rows=[], blank_rows=40,
           widths=[8, 14, 20, 14, 12, 12, 9, 12, 14, 14, 28],
           sheet_name="Request log", landscape=True,
           notes="Status: Pending, Approved, Denied or Cancelled. Log requests in the order received.")
summary_cells(log_path, [
    ("Summary", ""),
    ("Pending", '=COUNTIF(H5:H44,"Pending")'),
    ("Approved", '=COUNTIF(H5:H44,"Approved")'),
    ("Denied", '=COUNTIF(H5:H44,"Denied")'),
    ("Approved hours", '=SUMIF(H5:H44,"Approved",G5:G44)'),
], col_label="M", col_value="N", label_width=18, value_width=10)
pdf_table(d / "time-off-request-log.pdf", "Time off request log", log_headers, blank_rows=18,
          subtitle="Team or location: ______________   Year: ______",
          notes="Status: Pending, Approved, Denied or Cancelled. Log requests in the order received.")

# 8. Shift worker time off and coverage request -------------------------------------------------
both("shift-coverage-time-off-request-form", "Time off and shift coverage request", [
    ("p", "For employees on a posted schedule. List each shift you need off and, if your workplace allows it, the "
          "coworker who agreed to cover it. A manager must approve before the schedule changes."),
    ("fields", ["Employee name", "Position", "Location", "Date of request"]),
    ("h", "Shifts affected"),
    ("table", ["Date", "Shift (start to end)", "Position", "Covered by", "Coworker initials"], [["", "", "", "", ""]] * 4),
    ("checks", ["Vacation or PTO", "Personal", "Unpaid", "Shift swap (I will work the coworker's shift)"]),
    ("fields", ["If swapping: shift I will work instead"]),
    ("sign", ["Employee signature", "Covering coworker signature"]),
    ("h", "Manager use only"),
    ("checks", ["Approved", "Denied", "Schedule updated"]),
    ("sign", ["Manager signature"]),
])

# 9. Leave of absence request (extended or FMLA-qualifying) ------------------------------------
both("leave-of-absence-request-form", "Leave of absence request", [
    ("p", "Use this form to request leave of a week or more, or intermittent leave, for medical, family, military "
          "or personal reasons. Give it to HR at least 30 days ahead when the need is foreseeable, or as soon as you can."),
    ("fields", ["Employee name", "Job title", "Department", "Manager", "Phone", "Date of request"]),
    ("h", "Reason for leave"),
    ("checks", ["My own serious health condition", "Care of a spouse, child or parent with a serious health condition"]),
    ("checks", ["Birth, adoption or foster placement", "Military family leave", "Personal leave of absence", "Other"]),
    ("h", "Dates and schedule"),
    ("fields", ["Expected first day of leave", "Expected return date"]),
    ("checks", ["Continuous", "Intermittent", "Reduced schedule"]),
    ("fields", ["If intermittent or reduced: expected pattern", "Accrued PTO to use during leave (hours)"]),
    ("sign", ["Employee signature"]),
    ("h", "HR use only"),
    ("fields", ["Date received", "Eligibility notice given (date)", "Certification requested (yes/no, due date)",
                "Designation notice given (date)", "Notes"]),
    ("sign", ["HR signature"]),
])

# 10. Time off request email templates ---------------------------------------------------------
EMAIL_TEXT = """Subject: Time off request: [dates]

Hi [manager name],

I would like to request time off from [first day] through [last day], returning on [return date]. That is [number] workdays, or [hours] hours, and I would like to use [PTO / unpaid time] for it.

[Coworker name] has agreed to cover [tasks or shifts], and I will finish [project or task] before I leave.

Please let me know if these dates work or if you need anything else from me.

Thank you,
[Your name]"""

SHORT_NOTICE = """Subject: Short-notice time off request for [date]

Hi [manager name],

Something has come up and I need to request [the day off / the afternoon off] on [date], from [time] to [time]. I am sorry for the short notice. I have [arranged coverage with coworker / left notes for the team] and can [make up the hours on date / use PTO].

Please let me know if that works.

Thank you,
[Your name]"""

SICK = """Subject: Out sick today, [date]

Hi [manager name],

I am sick and will not be able to work my [start time] to [end time] shift today. I would like to use paid sick time for these hours. [Optional: I have let coworker know about task.]

I will let you know by [time] tomorrow whether I can work my next shift.

Thank you,
[Your name]"""

docx_doc(d / "time-off-request-email-templates.docx", "Time off request email templates", [
    ("h", "1. Planned time off (vacation or personal days)"), ("letter", EMAIL_TEXT),
    ("h", "2. Short-notice request"), ("letter", SHORT_NOTICE),
    ("h", "3. Calling out sick"), ("letter", SICK),
])
pdf_doc(d / "time-off-request-email-templates.pdf", "Time off request email templates", [
    ("h", "1. Planned time off (vacation or personal days)"), ("letter", EMAIL_TEXT),
    ("h", "2. Short-notice request"), ("letter", SHORT_NOTICE),
    ("h", "3. Calling out sick"), ("letter", SICK),
])

for f in sorted(d.iterdir()):
    print(f"{f.name:55s} {f.stat().st_size:>8,d} bytes")
