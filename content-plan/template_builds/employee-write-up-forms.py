#!/usr/bin/env python3
"""Build the downloads for /templates/employee-write-up-forms/.

Run from the site root:  python content-plan/template_builds/employee-write-up-forms.py
Files go to assets/templates/employee-write-up-forms/. Forms are blank and ready to fill in;
the filled-in examples live in the page previews.
"""
import sys
import pathlib

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent.parent))
import tplkit
# tplkit's fixed .foot overlaps the last text line when a page is full; print the footer in the
# bottom page margin instead (Chrome page-margin boxes). Local to this build; tplkit is unchanged.
tplkit._CSS += (".foot { display: none; }\n@page { margin-bottom: 18mm; @bottom-left { content: "
                "\"Free template from workzenly.com\"; font-family: Arial, sans-serif; font-size: 8pt; color: #888; } }\n")
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from tplkit import out, xlsx_sheet, docx_doc, pdf_doc, pdf_table

d = out("employee-write-up-forms")
thin = Side(style="thin", color="C9C9C3")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
ACK = ("My signature confirms that this write-up was discussed with me and that I received a copy. It does not mean "
       "I agree. I may attach a written response.")


def both(name, title, blocks):
    docx_doc(d / f"{name}.docx", title, blocks)
    pdf_doc(d / f"{name}.pdf", title, blocks)


def summary_box(ws, cells, col_label, col_value, start_row=4, lw=26, vw=12):
    ws.column_dimensions[col_label].width = lw
    ws.column_dimensions[col_value].width = vw
    for i, (label, value) in enumerate(cells):
        a, b = ws[f"{col_label}{start_row + i}"], ws[f"{col_value}{start_row + i}"]
        a.value, b.value = label, value
        for c in (a, b):
            c.border = BOX
            c.alignment = Alignment(vertical="center", wrap_text=True)
        if i == 0:
            for c in (a, b):
                c.font = Font(bold=True, color="FFFFFF")
                c.fill = PatternFill("solid", fgColor="161616")
        else:
            a.font = Font(bold=True)
            b.fill = PatternFill("solid", fgColor="F2FFD9" if str(value).startswith("=") else "FFF6CC")


# 1. Simple employee write-up form -------------------------------------------------------------
both("simple-employee-write-up-form", "Employee write-up form", [
    ("fields", ["Employee name", "Job title", "Manager", "Date of write-up", "Date of incident"]),
    ("h", "Type of issue"),
    ("checks", ["Attendance", "Tardiness", "Performance", "Conduct", "Safety", "Policy violation", "Other"]),
    ("h", "Description"),
    ("fields", ["What happened (facts, dates, times)", "(continued)", "(continued)", "Policy or rule"]),
    ("h", "Action"),
    ("checks", ["Documented verbal warning", "Written warning", "Final warning", "Suspension", "Termination"]),
    ("fields", ["What must change", "Consequence if it happens again"]),
    ("p", ACK),
    ("sign", ["Employee signature", "Manager signature"]),
])

# 2. Detailed disciplinary action form ----------------------------------------------------------
both("disciplinary-action-form", "Disciplinary action form", [
    ("fields", ["Employee name and ID", "Job title and department", "Supervisor", "Date of this action"]),
    ("h", "Discipline step"),
    ("checks", ["Step 1: documented verbal", "Step 2: written", "Step 3: final written", "Suspension (dates)", "Termination"]),
    ("h", "Incident"),
    ("fields", ["Date(s), time and place", "Description (facts only)", "(continued)", "Policy, rule or standard",
                "Evidence reviewed (records, video, witnesses)"]),
    ("h", "Prior discipline on this or related issues"),
    ("table", ["Date", "Step", "Issue"], [["", "", ""]] * 2),
    ("h", "Corrective action plan"),
    ("fields", ["Expected improvement", "Deadline or review date", "Support provided"]),
    ("fields", ["Consequence of further violations"]),
    ("fields", ["Employee comments"]),
    ("p", ACK),
    ("sign", ["Employee signature", "Supervisor signature", "HR or second manager"]),
])

# 3. Attendance write-up form with occurrence log (Excel) ---------------------------------------
att_headers = ["Date", "Scheduled start", "Clock-in", "Minutes late", "Absent (Y/N)", "Called in (Y/N)",
               "Protected leave (Y/N)", "Points"]
att_path = d / "attendance-write-up-form.xlsx"
xlsx_sheet(att_path, title="Attendance write-up form",
           subtitle="Employee: ____________   Position: ____________   Manager: ____________   Date: ________",
           headers=att_headers, rows=[], blank_rows=12, widths=[12, 14, 12, 12, 12, 13, 15, 10],
           formulas={
               "Minutes late": '=IF(OR(B{r}="",C{r}=""),"",MAX(0,ROUND((C{r}-B{r})*1440,0)))',
               "Points": '=IF(A{r}="","",IF(G{r}="Y",0,IF(E{r}="Y",IF(F{r}="Y",1,2),IF(N(D{r})>5,0.5,0))))',
           },
           sheet_name="Attendance", landscape=True,
           notes="Type times as 7:00 AM. Points: late more than 5 minutes = 0.5, absent with call-in = 1, no call no show = 2, "
                 "protected leave (FMLA, legally protected sick leave, jury duty) = 0. Edit the formula to match your policy.")
wb = load_workbook(att_path)
ws = wb.active
for r in range(5, 17):
    for col in "BC":
        ws[f"{col}{r}"].number_format = "h:mm AM/PM"
summary_box(ws, [
    ("Summary", ""),
    ("Total points", "=SUM(H5:H16)"),
    ("Late arrivals", '=COUNTIF(D5:D16,">5")'),
    ("Absences", '=COUNTIF(E5:E16,"Y")'),
    ("No call, no show", '=COUNTIFS(E5:E16,"Y",F5:F16,"N")'),
    ("Protected (not counted)", '=COUNTIF(G5:G16,"Y")'),
    ("Policy limit for this step", ""),
    ("Action taken", ""),
], col_label="J", col_value="K")
ws.column_dimensions["I"].width = 3
wb.save(att_path)
pdf_doc(d / "attendance-write-up-form.pdf", "Attendance write-up form", [
    ("fields", ["Employee", "Position", "Manager", "Date"]),
    ("table", att_headers, [[""] * 8] * 8),
    ("fields", ["Total points", "Policy limit for this step", "Action taken"]),
    ("p", "Protected leave (FMLA, legally protected sick leave, jury duty and other protected absences) is not counted."),
    ("p", ACK),
    ("sign", ["Employee signature", "Manager signature"]),
])

# 4. Restaurant and retail write-up form -------------------------------------------------------
both("restaurant-retail-write-up-form", "Restaurant and retail employee write-up form", [
    ("fields", ["Employee name", "Position", "Location", "Shift (date and time)", "Manager on duty"]),
    ("h", "Issue"),
    ("checks", ["Cash handling or register variance", "Voids, comps or discounts", "Food safety or sanitation"]),
    ("checks", ["Guest or customer service", "Dress code or hygiene", "Side work or closing duties"]),
    ("checks", ["Late, left early or missed break rules", "Phone use", "Other"]),
    ("fields", ["What happened (facts, amounts, ticket or receipt numbers)", "(continued)", "Policy or standard"]),
    ("h", "Action"),
    ("checks", ["Coaching (documented)", "Written warning", "Final warning", "Removed from shift (paid / unpaid)"]),
    ("fields", ["What must change", "Retraining (date and trainer)", "Next step if it happens again"]),
    ("p", ACK),
    ("sign", ["Employee signature", "Manager signature"]),
])

# 5. Safety violation write-up form ------------------------------------------------------------
both("safety-violation-write-up-form", "Safety violation write-up form", [
    ("fields", ["Employee name", "Job title", "Site or job number", "Supervisor", "Date and time of violation"]),
    ("h", "Violation"),
    ("checks", ["PPE not worn", "Lockout / tagout", "Equipment or vehicle", "Fall protection", "Housekeeping or blocked exit"]),
    ("checks", ["Unsafe lifting", "Chemical handling", "Skipped inspection", "Horseplay", "Other"]),
    ("fields", ["What happened (facts)", "(continued)", "Safety rule or procedure", "Anyone hurt or near miss? (describe)"]),
    ("h", "Immediate action taken"),
    ("checks", ["Work stopped", "Hazard corrected", "Equipment tagged out", "Incident report filed"]),
    ("h", "Discipline and retraining"),
    ("checks", ["Documented coaching", "Written warning", "Final warning", "Suspension", "Termination"]),
    ("fields", ["Retraining required (topic, date, trainer)", "Next step if repeated"]),
    ("p", ACK),
    ("sign", ["Employee signature", "Supervisor signature"]),
])

# 6. Performance improvement plan write-up ----------------------------------------------------
both("performance-improvement-plan-write-up", "Performance improvement plan (PIP)", [
    ("fields", ["Employee name", "Job title", "Manager", "Plan start date", "Plan end date"]),
    ("h", "Why this plan"),
    ("fields", ["Performance gap (facts and numbers)", "(continued)", "Earlier feedback or warnings (dates)"]),
    ("h", "Goals"),
    ("table", ["Goal", "Measure of success", "Target", "Due", "Support"], [["", "", "", "", ""]] * 4),
    ("h", "Check-ins"),
    ("table", ["Date", "Progress", "Manager initials", "Employee initials"], [["", "", "", ""]] * 4),
    ("h", "Outcome at end of plan"),
    ("checks", ["Goals met, plan closed", "Plan extended to (date)", "Further discipline", "Termination"]),
    ("p", ACK + " This plan does not change the at-will nature of employment."),
    ("sign", ["Employee signature", "Manager signature"]),
])

# 7. Incident report and witness statement form -----------------------------------------------
both("incident-report-and-witness-statement-form", "Workplace incident report and witness statement", [
    ("h", "Part 1: Incident report (manager)"),
    ("fields", ["Reported by", "Date and time of incident", "Location", "Employees involved", "Witnesses"]),
    ("fields", ["What was reported", "(continued)", "Evidence collected (video, receipts, logs, photos)"]),
    ("h", "Part 2: Witness statement"),
    ("fields", ["Witness name and job title", "Date and time of statement"]),
    ("p", "Describe only what you saw and heard yourself, in your own words, with times and places."),
    ("fields", ["Statement", "(continued)", "(continued)", "(continued)"]),
    ("checks", ["This statement is true and complete to the best of my knowledge"]),
    ("sign", ["Witness signature", "Manager taking the statement"]),
])

# 8. Disciplinary action log (Excel) ---------------------------------------------------------
log_headers = ["Date", "Employee", "Position", "Issue type", "Action", "Written up by", "Signed (Y/N)",
               "Follow-up date", "Outcome", "Notes"]
log_path = d / "disciplinary-action-log.xlsx"
xlsx_sheet(log_path, title="Disciplinary action log",
           subtitle="Business or location: ____________   Year: ______",
           headers=log_headers, rows=[], blank_rows=50,
           widths=[12, 20, 16, 16, 16, 16, 10, 14, 18, 30],
           sheet_name="Discipline log", landscape=True,
           notes="Issue type: Attendance, Performance, Conduct, Safety, Policy. "
                 "Action: Coaching, Written, Final, Suspension, Termination.")
wb = load_workbook(log_path)
ws = wb.active
summary_box(ws, [
    ("Summary", ""),
    ("Attendance", '=COUNTIF(D5:D54,"Attendance")'),
    ("Performance", '=COUNTIF(D5:D54,"Performance")'),
    ("Conduct", '=COUNTIF(D5:D54,"Conduct")'),
    ("Safety", '=COUNTIF(D5:D54,"Safety")'),
    ("Policy", '=COUNTIF(D5:D54,"Policy")'),
    ("Final warnings", '=COUNTIF(E5:E54,"Final")'),
    ("Terminations", '=COUNTIF(E5:E54,"Termination")'),
], col_label="L", col_value="M", lw=16, vw=10)
wb.save(log_path)
pdf_table(d / "disciplinary-action-log.pdf", "Disciplinary action log", log_headers, blank_rows=16,
          subtitle="Business or location: ______________   Year: ______",
          notes="Issue type: Attendance, Performance, Conduct, Safety, Policy. Action: Coaching, Written, Final, Suspension, Termination.")

for f in sorted(d.iterdir()):
    print(f"{f.name:52s} {f.stat().st_size:>8,d} bytes")
