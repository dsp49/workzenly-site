#!/usr/bin/env python3
"""Build the downloads for /templates/performance-review-templates/.

Run from the site root:  python content-plan/template_builds/performance-review-templates.py
Files go to assets/templates/performance-review-templates/. Files are blank and ready to fill in;
the filled-in examples live in the page previews.
"""
import sys
import pathlib

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent.parent))
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.worksheet.datavalidation import DataValidation
from tplkit import out, xlsx_sheet, docx_doc, pdf_table, pdf_doc

d = out("performance-review-templates")
thin = Side(style="thin", color="C9C9C3")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
SCALE = ("Rating scale: 5 = Outstanding, 4 = Exceeds expectations, 3 = Meets expectations, "
         "2 = Needs improvement, 1 = Unsatisfactory.")


def both(name, title, blocks, landscape=False):
    docx_doc(d / f"{name}.docx", title, blocks)
    pdf_doc(d / f"{name}.pdf", title, blocks, landscape=landscape)


def add_summary(path, cells, col_label, col_value, start_row):
    wb = load_workbook(path)
    ws = wb.active
    for i, (label, value) in enumerate(cells):
        r = start_row + i
        a, b = ws[f"{col_label}{r}"], ws[f"{col_value}{r}"]
        a.value, b.value = label, value
        a.font = Font(bold=True)
        a.alignment = Alignment(wrap_text=True, vertical="center")
        if str(value).startswith("="):
            b.fill = PatternFill("solid", fgColor="F2FFD9")
        for c in (a, b):
            c.border = BOX
    wb.save(path)
    return wb


# 1. Annual performance review ------------------------------------------------------------------
both("annual-performance-review-template", "Annual performance review", [
    ("fields", ["Employee name", "Job title", "Department or location", "Reviewer", "Review period",
                "Date of review meeting"]),
    ("h", "Part 1. Goals from last year"),
    ("table", ["Goal", "Result", "Met / Partly / Not met"], [["", "", ""]] * 4),
    ("h", "Part 2. Core responsibilities"),
    ("p", SCALE),
    ("table", ["Responsibility", "Rating (1 to 5)", "Evidence"], [["", "", ""]] * 5),
    ("h", "Part 3. Competencies"),
    ("table", ["Competency", "Rating (1 to 5)", "Evidence"],
     [["Quality and accuracy", "", ""], ["Reliability", "", ""], ["Communication", "", ""],
      ["Teamwork", "", ""], ["Problem solving", "", ""], ["Customer focus", "", ""]]),
    ("h", "Part 4. Overall"),
    ("checks", ["5 Outstanding", "4 Exceeds", "3 Meets", "2 Needs improvement", "1 Unsatisfactory"]),
    ("fields", ["Summary of the year (2 to 3 sentences)"]),
    ("h", "Part 5. Goals for next year"),
    ("table", ["Goal", "How it will be measured", "Due date"], [["", "", ""]] * 3),
    ("h", "Part 6. Development plan"),
    ("fields", ["Skill to build", "Training, mentoring or project", "Support from the manager"]),
    ("h", "Part 7. Comments and signatures"),
    ("fields", ["Employee comments", "Reviewer comments"]),
    ("p", "Signing confirms the review was discussed. It does not mean the employee agrees with it."),
    ("sign", ["Employee signature", "Reviewer signature", "Second-level approval (optional)"]),
])

# 2. One-page performance review ----------------------------------------------------------------
both("one-page-performance-review-template", "One-page performance review", [
    ("fields", ["Employee name", "Job title", "Reviewer", "Review period"]),
    ("h", "Keep doing (strengths)"),
    ("fields", ["1", "2", "3"]),
    ("h", "Start doing"),
    ("fields", ["1", "2"]),
    ("h", "Stop doing"),
    ("fields", ["1"]),
    ("h", "Overall"),
    ("checks", ["Exceeds expectations", "Meets expectations", "Below expectations"]),
    ("h", "Top goals for next period"),
    ("fields", ["1", "2", "3"]),
    ("sign", ["Employee signature", "Reviewer signature"]),
])

# 3. Goal-based (SMART goals) review, Excel ------------------------------------------------------
g_headers = ["Goal (specific and measurable)", "Weight (%)", "Target", "Actual", "% achieved",
             "Weighted %", "Notes"]
g_path = d / "goal-based-performance-review.xlsx"
xlsx_sheet(g_path, title="Goal-based performance review",
           subtitle="Employee: ____________   Job title: ____________   Reviewer: ____________   Period: ________",
           headers=g_headers, rows=[], blank_rows=8, widths=[42, 11, 12, 12, 12, 12, 36],
           formulas={"% achieved": '=IF(OR(C{r}="",D{r}=""),"",MIN(D{r}/C{r},1.5))',
                     "Weighted %": '=IF(E{r}="","",B{r}/100*E{r})'},
           sheet_name="Goals", landscape=True,
           notes="Enter numeric targets and actuals. For goals where lower is better, such as errors or wait time, "
                 "type the actual in the Target column and the target in the Actual column. % achieved is capped at 150%. "
                 "Score bands: 110%+ Exceeds, 90-109% Meets, below 90% Below expectations.")
wb = load_workbook(g_path)
ws = wb.active
for r in range(5, 13):
    ws[f"E{r}"].number_format = "0%"
    ws[f"F{r}"].number_format = "0%"
note, foot = ws["A14"].value, ws["A16"].value
ws["A14"].value = ws["A16"].value = None
wb.save(g_path)
wb = add_summary(g_path, [
    ("Total weight (must be 100)", "=SUM(B5:B12)"),
    ("Overall goal score", '=IF(COUNT(F5:F12)=0,"",SUM(F5:F12))'),
    ("Result", '=IF(F15="","",IF(F15>=1.1,"Exceeds expectations",IF(F15>=0.9,"Meets expectations",'
               '"Below expectations")))'),
], "E", "F", 14)
ws = wb.active
ws["F15"].number_format = "0%"
ws["A18"].value = note
ws["A18"].font = Font(italic=True, size=10, color="555555")
ws["A20"].value = foot
ws["A20"].font = Font(size=9, color="888888")
wb.save(g_path)
pdf_doc(d / "goal-based-performance-review.pdf", "Goal-based performance review", [
    ("fields", ["Employee name", "Job title", "Reviewer", "Review period"]),
    ("table", ["Goal", "Weight", "Target", "Actual", "% achieved", "Weighted %"], [[""] * 6] * 6),
    ("fields", ["Overall goal score (sum of weighted %)", "Result"]),
    ("p", "% achieved = actual / target, capped at 150%. 110% and up Exceeds, 90% to 109% Meets, "
          "below 90% Below expectations."),
    ("fields", ["Comments"]),
    ("sign", ["Employee signature", "Reviewer signature"]),
])

# 4. Quarterly check-in -------------------------------------------------------------------------
both("quarterly-check-in-template", "Quarterly performance check-in", [
    ("p", "A 20-minute conversation, four times a year. The employee fills in the first column before the "
          "meeting; the manager adds notes during it."),
    ("fields", ["Employee name", "Manager", "Quarter", "Date"]),
    ("table", ["Question", "Employee", "Manager"], [
        ["What went well this quarter?", "", ""],
        ["What got in the way?", "", ""],
        ["Progress on each goal (on track / at risk / done)", "", ""],
        ["What support do you need?", "", ""],
        ["Schedule, hours or time off concerns", "", ""],
        ["Focus for next quarter", "", ""],
    ]),
    ("fields", ["Agreed next steps", "Next check-in date"]),
    ("sign", ["Employee initials", "Manager initials"]),
])

# 5. Competency-based review with rubric --------------------------------------------------------
both("competency-based-performance-review", "Competency-based performance review", [
    ("fields", ["Employee name", "Job title", "Reviewer", "Review period"]),
    ("p", "Circle the description that best matches the employee's typical performance over the whole period. "
          "Write one example for each competency."),
    ("table", ["Competency", "2 Needs improvement", "3 Meets expectations", "4 Exceeds expectations", "Example"], [
        ["Reliability", "Late or absent more than policy allows; misses handoffs",
         "Works scheduled shifts, follows call-out rules", "Covers gaps, plans time off early", ""],
        ["Quality", "Frequent errors or redos", "Work meets the standard with normal checking",
         "Rarely needs checking; catches others' errors", ""],
        ["Customer focus", "Complaints or slow follow-up", "Polite, solves routine issues",
         "Turns problems into repeat customers", ""],
        ["Teamwork", "Avoids shared tasks", "Does their share, helps when asked", "Helps before being asked; trains others", ""],
        ["Communication", "Information is missing or late", "Shares what others need on time",
         "Clear in writing and in person; keeps the team informed", ""],
        ["Safety and compliance", "Skips required steps", "Follows rules and checklists", "Spots and reports risks", ""],
    ]),
    ("p", "Use 1 (Unsatisfactory) or 5 (Outstanding) only with a written example that clearly falls below or "
          "above these descriptions."),
    ("fields", ["Overall rating (1 to 5)", "Development focus"]),
    ("sign", ["Employee signature", "Reviewer signature"]),
], landscape=True)

# 6. 360-degree feedback summary, Excel ---------------------------------------------------------
s_headers = ["Competency", "Self", "Manager", "Peer 1", "Peer 2", "Peer 3", "Direct report avg",
             "Others' average", "Gap (self - others)"]
s_rows = [[c, "", "", "", "", "", "", "", ""] for c in
          ["Quality of work", "Reliability", "Communication", "Teamwork", "Problem solving", "Leadership",
           "Customer focus"]]
s_path = d / "360-degree-feedback-summary.xlsx"
xlsx_sheet(s_path, title="360-degree feedback summary",
           subtitle="Employee: ____________   Period: ________   Ratings 1 to 5, leave blank if not observed",
           headers=s_headers, rows=s_rows, blank_rows=1, widths=[22, 9, 10, 9, 9, 9, 14, 13, 14],
           formulas={"Others' average": '=IF(COUNT(C{r}:G{r})=0,"",ROUND(AVERAGE(C{r}:G{r}),2))',
                     "Gap (self - others)": '=IF(OR(B{r}="",H{r}=""),"",ROUND(B{r}-H{r},2))'},
           sheet_name="360 summary", landscape=True,
           notes="Others' average = manager, peers and direct reports. A gap of +1 or more means the employee rates "
                 "themselves higher than others do; -1 or more means they undersell themselves. Share themes, not "
                 "individual peer scores.")
wb = load_workbook(s_path)
ws = wb.active
dv = DataValidation(type="decimal", operator="between", formula1="1", formula2="5", allow_blank=True,
                    showErrorMessage=True, error="Enter a rating from 1 to 5.")
ws.add_data_validation(dv)
dv.add("B5:G12")
wb.save(s_path)
pdf_table(d / "360-degree-feedback-summary.pdf", "360-degree feedback summary", s_headers,
          [r[:1] + [""] * 8 for r in s_rows], blank_rows=1,
          subtitle="Employee: ____________   Period: ________   Ratings 1 to 5", landscape=True,
          notes="Others' average = mean of manager, peer and direct report ratings. Gap = self rating minus "
                "others' average.")

# 7. Performance improvement plan ---------------------------------------------------------------
both("performance-improvement-plan-template", "Performance improvement plan (PIP)", [
    ("fields", ["Employee name", "Job title", "Manager", "HR or second reviewer", "Plan start date",
                "Plan end date (30, 60 or 90 days)"]),
    ("h", "1. The performance gap"),
    ("p", "Describe what is happening, with dates and facts, and what the job requires instead. Refer to earlier "
          "conversations or written warnings."),
    ("table", ["Expected standard", "Current performance (with dates)", "Earlier conversations"], [["", "", ""]] * 3),
    ("h", "2. Goals for the plan"),
    ("table", ["Measurable goal", "Deadline", "Support or training provided"], [["", "", ""]] * 3),
    ("h", "3. Check-in meetings"),
    ("table", ["Date", "Progress (on track / not on track)", "Notes"], [["", "", ""]] * 4),
    ("h", "4. Outcome"),
    ("checks", ["Plan completed successfully", "Plan extended to: ________", "Further action taken"]),
    ("p", "This plan does not change the employment relationship or guarantee employment for the plan period "
          "where employment is at will. Failing to meet the goals may lead to further action, up to and "
          "including termination."),
    ("fields", ["Employee comments"]),
    ("sign", ["Employee signature (acknowledges receipt)", "Manager signature", "HR or second reviewer"]),
])

print("built", sorted(p.name for p in d.iterdir()))
