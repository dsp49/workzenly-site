#!/usr/bin/env python3
"""Build the downloads for /templates/employee-warning-notice-templates/.

Run from the site root:  python content-plan/template_builds/employee-warning-notice-templates.py
Files go to assets/templates/employee-warning-notice-templates/. Forms are blank; letters use [brackets].
The filled-in examples live in the page previews.
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
from tplkit import out, xlsx_sheet, docx_doc, pdf_doc, pdf_table

d = out("employee-warning-notice-templates")
thin = Side(style="thin", color="C9C9C3")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)

ACK = ("By signing, I confirm that I received this warning and that it was discussed with me. My signature does not "
       "mean I agree. I may add a written response, which will be kept with this notice.")


def both(name, title, blocks):
    docx_doc(d / f"{name}.docx", title, blocks)
    pdf_doc(d / f"{name}.pdf", title, blocks)


# 1. Employee warning notice form -------------------------------------------------------------
both("employee-warning-notice-form", "Employee warning notice", [
    ("fields", ["Employee name", "Job title and department", "Manager", "Date of notice"]),
    ("h", "Type of warning"),
    ("checks", ["Verbal warning (documented)", "First written warning", "Second written warning", "Final written warning"]),
    ("h", "Reason"),
    ("checks", ["Attendance or tardiness", "Performance", "Conduct", "Safety", "Policy violation", "Other"]),
    ("fields", ["Date(s) of the issue", "What happened (facts only)", "(continued)", "Policy or standard involved",
                "Earlier warnings on this issue (dates)"]),
    ("h", "Expected improvement"),
    ("fields", ["What must change", "By when", "Support offered (training, schedule, check-ins)",
                "Next step if the issue continues"]),
    ("p", ACK),
    ("sign", ["Employee signature", "Manager signature"]),
])

# 2. Verbal warning record -------------------------------------------------------------------
both("verbal-warning-record", "Verbal warning record", [
    ("p", "Use this record to document a verbal warning right after the conversation. Give the employee a copy."),
    ("fields", ["Employee name", "Job title", "Manager", "Date and time of conversation", "Where", "Witness (optional)"]),
    ("h", "Issue discussed"),
    ("fields", ["What happened (facts, dates, times)", "(continued)", "Rule, policy or standard", "What the employee said"]),
    ("h", "Agreed next steps"),
    ("fields", ["What will change", "Follow-up date"]),
    ("checks", ["Employee was told a written warning may follow if the issue continues"]),
    ("sign", ["Manager signature", "Employee signature (acknowledges the conversation)"]),
])

# 3. Written warning letter ------------------------------------------------------------------
both("written-warning-letter", "Written warning letter", [
    ("letter", """[Company name]

[Date]

To: [Employee name], [Job title]
From: [Manager name], [Title]
Subject: Written warning: [performance / conduct issue]

This letter is a written warning about [issue in a few words].

What happened
[Describe the facts: dates, what was expected, what actually happened, with numbers or records where possible.]

Earlier discussion
We discussed this issue on [date(s)] [in a verbal warning on date]. [What was agreed then.]

What must change
- [Specific, measurable expectation]
- [Second expectation, if any]
Starting [date], I will review progress with you [weekly on day]. We will offer [training, a schedule change, a checklist or other support].

What happens next
If the required improvement is not made by [date], or the issue happens again, the next step may be [a final written warning / further discipline up to and including termination]. This letter does not change the at-will nature of your employment.

You may add a written response, which will be kept in your file with this letter.

[Manager name]
[Title]"""),
    ("p", ACK),
    ("sign", ["Employee signature", "Manager signature"]),
])

# 4. Attendance and tardiness warning notice -------------------------------------------------
both("attendance-warning-notice", "Attendance and tardiness warning notice", [
    ("fields", ["Employee name", "Position and location", "Manager", "Date of notice"]),
    ("checks", ["First written warning", "Second written warning", "Final written warning"]),
    ("h", "Attendance record (from time clock and schedule)"),
    ("table", ["Date", "Scheduled start", "Clock-in", "Minutes late / absent", "Notice given?", "Points"],
     [["", "", "", "", "", ""]] * 6),
    ("fields", ["Total points in the rolling period", "Policy limit before this level of warning"]),
    ("p", "Absences protected by law are not counted, including FMLA leave, paid sick leave under state or local law, "
          "jury duty and other protected leave. Tell your manager or HR if any absence above should be protected."),
    ("fields", ["Expectation going forward", "Next step if the issue continues"]),
    ("p", ACK),
    ("sign", ["Employee signature", "Manager signature"]),
])

# 5. Final written warning letter -------------------------------------------------------------
both("final-written-warning-letter", "Final written warning", [
    ("letter", """[Company name]

[Date]

PERSONAL AND CONFIDENTIAL

To: [Employee name], [Job title]
From: [Manager name], [Title]
Subject: Final written warning

This is a final written warning about [issue].

History
- [Date]: verbal warning for [issue]
- [Date]: written warning for [issue]
- [Date of latest incident]: [what happened, facts only]

Despite these warnings and the support we offered ([training, check-ins, schedule change]), the issue has continued.

What must change, starting immediately
- [Specific, measurable expectation]
- [Second expectation]

This is the last step before termination. If [the issue] happens again within [__ days / months], or you do not meet the expectations above, your employment may be terminated. This warning does not change the at-will nature of your employment, and the company may skip steps for serious misconduct.

We want you to succeed. If something is making it hard to meet these expectations, including a health issue or a need for leave or an accommodation, please tell me or HR.

You may add a written response, which will be kept in your file with this letter.

[Manager name]
[Title]"""),
    ("p", ACK),
    ("sign", ["Employee signature", "Manager signature", "Witness (if employee declines to sign)"]),
])

# 6. Conduct warning letter -------------------------------------------------------------------
both("conduct-warning-letter", "Conduct warning letter", [
    ("letter", """[Company name]

[Date]

To: [Employee name], [Job title]
From: [Manager name], [Title]
Subject: Written warning: workplace conduct

This letter is a written warning about your conduct on [date].

What happened
On [date] at about [time], [describe the behavior factually, what was said or done, who was present]. [We spoke with (number) witnesses and reviewed (records).]

Why it matters
This conduct does not meet our [policy name, handbook section], which requires [the standard in plain words].

What we expect
- [Specific behavior expected, such as: speak to customers and coworkers without profanity or raised voices]
- [If you disagree with a coworker or customer, step away and bring in the shift lead]

If this conduct happens again, the next step may be [a final written warning / termination]. This letter does not change the at-will nature of your employment.

You may add a written response, which will be kept in your file with this letter.

[Manager name]
[Title]"""),
    ("p", ACK),
    ("sign", ["Employee signature", "Manager signature"]),
])

# 7. Warning acknowledgment and employee response form ----------------------------------------
both("warning-acknowledgment-and-response-form", "Warning acknowledgment and employee response", [
    ("fields", ["Employee name", "Date of warning", "Warning given by", "Type of warning"]),
    ("h", "Acknowledgment"),
    ("p", "I received a copy of the warning dated above. My signature confirms that I received it, not that I agree with it."),
    ("sign", ["Employee signature"]),
    ("h", "My response (optional)"),
    ("checks", ["I agree with the warning", "I partly agree", "I disagree"]),
    ("fields", ["My explanation", "(continued)", "(continued)", "(continued)", "Facts, documents or witnesses I want considered", "(continued)"]),
    ("h", "For the manager or HR"),
    ("fields", ["Date response received", "Reviewed by", "Action taken"]),
    ("sign", ["Manager or HR signature"]),
])

# 8. Warning notice tracker (Excel) -----------------------------------------------------------
trk_headers = ["Date", "Employee", "Department", "Type", "Reason", "Policy", "Issued by",
               "Signed (Y/N)", "Response (Y/N)", "Review date", "Expires", "Notes"]
trk_path = d / "warning-notice-tracker.xlsx"
xlsx_sheet(trk_path, title="Employee warning notice tracker",
           subtitle="Location: ____________   Year: ______",
           headers=trk_headers, rows=[], blank_rows=40,
           widths=[12, 20, 14, 16, 16, 14, 14, 10, 11, 12, 12, 28],
           sheet_name="Warnings", landscape=True,
           notes="Type: Verbal, Written, Second written or Final. Reason: Attendance, Performance, Conduct, Safety, Policy.")
wb = load_workbook(trk_path)
ws = wb.active
summary = [("Summary", ""), ("Verbal", '=COUNTIF(D5:D44,"Verbal")'), ("Written", '=COUNTIF(D5:D44,"Written")'),
           ("Second written", '=COUNTIF(D5:D44,"Second written")'), ("Final", '=COUNTIF(D5:D44,"Final")'),
           ("Not signed", '=COUNTIF(H5:H44,"N")')]
ws.column_dimensions["N"].width = 18
ws.column_dimensions["O"].width = 10
for i, (label, val) in enumerate(summary):
    a, b = ws[f"N{4 + i}"], ws[f"O{4 + i}"]
    a.value, b.value = label, val
    for c in (a, b):
        c.border = BOX
        c.alignment = Alignment(vertical="center")
    if i == 0:
        for c in (a, b):
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor="161616")
    else:
        a.font = Font(bold=True)
        b.fill = PatternFill("solid", fgColor="F2FFD9")
wb.save(trk_path)
pdf_table(d / "warning-notice-tracker.pdf", "Employee warning notice tracker", trk_headers, blank_rows=16,
          subtitle="Location: ______________   Year: ______",
          notes="Type: Verbal, Written, Second written or Final. Reason: Attendance, Performance, Conduct, Safety, Policy.")

for f in sorted(d.iterdir()):
    print(f"{f.name:50s} {f.stat().st_size:>8,d} bytes")
