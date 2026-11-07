#!/usr/bin/env python3
"""Build the downloads for /templates/employee-evaluation-forms/.

Run from the site root:  python content-plan/template_builds/employee-evaluation-forms.py
Files go to assets/templates/employee-evaluation-forms/. Files are blank and ready to fill in;
the filled-in examples live in the page previews.
"""
import sys
import pathlib

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent.parent))
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.worksheet.datavalidation import DataValidation
from tplkit import out, xlsx_sheet, docx_doc, pdf_table, pdf_doc

d = out("employee-evaluation-forms")
SCALE = ("Rating scale: 5 = Outstanding, 4 = Exceeds expectations, 3 = Meets expectations, "
         "2 = Needs improvement, 1 = Unsatisfactory. Add a comment for any 1, 2 or 5.")
RATINGS = ["5 Outstanding", "4 Exceeds", "3 Meets", "2 Needs improvement", "1 Unsatisfactory"]
thin = Side(style="thin", color="C9C9C3")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)


def both(name, title, blocks):
    docx_doc(d / f"{name}.docx", title, blocks)
    pdf_doc(d / f"{name}.pdf", title, blocks)


def rating_rows(factors):
    return [[f, "", ""] for f in factors]


# 1. Simple employee evaluation form ------------------------------------------------------------
simple_factors = ["Quality of work", "Quantity of work", "Job knowledge", "Reliability and attendance",
                  "Communication", "Teamwork", "Initiative", "Follows policies and safety rules"]
both("simple-employee-evaluation-form", "Employee evaluation form", [
    ("fields", ["Employee name", "Job title", "Department or location", "Evaluator", "Review period (from / to)",
                "Date of evaluation"]),
    ("p", SCALE),
    ("table", ["Factor", "Rating (1 to 5)", "Comments and examples"], rating_rows(simple_factors)),
    ("fields", ["Overall rating (1 to 5)"]),
    ("h", "Strengths"),
    ("fields", ["1", "2"]),
    ("h", "Areas to improve"),
    ("fields", ["1", "2"]),
    ("h", "Goals for the next review period"),
    ("fields", ["1", "2", "3"]),
    ("p", "Employee comments (optional). Signing confirms the employee received this evaluation and discussed it "
          "with the evaluator. It does not mean the employee agrees with it."),
    ("fields", ["Employee comments"]),
    ("sign", ["Employee signature", "Evaluator signature"]),
])

# 2. Weighted scoring evaluation (Excel) ---------------------------------------------------------
w_headers = ["Factor", "What good looks like", "Weight (%)", "Rating (1-5)", "Weighted score", "Comments"]
w_rows = [
    ["Quality of work", "Accurate, complete, few redos", 25, "", "", ""],
    ["Productivity", "Meets volume and deadline targets", 20, "", "", ""],
    ["Job knowledge", "Knows procedures, tools and products", 15, "", "", ""],
    ["Reliability", "On time, follows the schedule, gives notice", 15, "", "", ""],
    ["Customer service", "Courteous, solves problems, follows up", 15, "", "", ""],
    ["Teamwork", "Helps others, shares information", 10, "", "", ""],
]
w_path = d / "weighted-employee-evaluation-form.xlsx"
xlsx_sheet(w_path, title="Employee evaluation form (weighted)",
           subtitle="Employee: ____________   Job title: ____________   Evaluator: ____________   Period: ________",
           headers=w_headers, rows=w_rows, blank_rows=2, widths=[22, 38, 12, 12, 15, 40],
           formulas={"Weighted score": '=IF(D{r}="","",C{r}/100*D{r})'},
           sheet_name="Evaluation", landscape=True,
           notes="Weights must add up to 100. Enter a rating of 1 to 5 for each factor. "
                 "Overall score = sum of weight x rating. 4.5+ Outstanding, 3.5-4.49 Exceeds, "
                 "2.5-3.49 Meets, 1.5-2.49 Needs improvement, below 1.5 Unsatisfactory.")
wb = load_workbook(w_path)
ws = wb.active
# header row 4, data rows 5-12
dv = DataValidation(type="whole", operator="between", formula1="1", formula2="5", allow_blank=True,
                    showErrorMessage=True, errorTitle="Rating", error="Enter a whole number from 1 to 5.")
ws.add_data_validation(dv)
dv.add("D5:D12")
summary = [
    ("Total weight (must be 100)", "=SUM(C5:C12)"),
    ("Overall score (out of 5)", '=IF(COUNT(D5:D12)=0,"",ROUND(SUMPRODUCT(C5:C12,D5:D12)/100,2))'),
    ("Overall rating", '=IF(E15="","",IF(E15>=4.5,"Outstanding",IF(E15>=3.5,"Exceeds expectations",'
                       'IF(E15>=2.5,"Meets expectations",IF(E15>=1.5,"Needs improvement","Unsatisfactory")))))'),
]
# move notes/footer down: they sit at row 14+; put the summary at rows 14-16 in columns D:E
for i, (label, f) in enumerate(summary):
    r = 14 + i
    a, b = ws[f"D{r}"], ws[f"E{r}"]
    a.value, b.value = label, f
    a.font = Font(bold=True)
    a.alignment = Alignment(wrap_text=True, vertical="center")
    b.fill = PatternFill("solid", fgColor="F2FFD9")
    for c in (a, b):
        c.border = BOX
# the summary formulas reference E15 for the overall score
ws["E16"].value = summary[2][1]
# notes were written at row 14 col A; shift them below the summary
note = ws["A14"].value
foot = ws["A16"].value
ws["A14"].value, ws["A16"].value = None, None
ws["A18"].value = note
ws["A18"].font = Font(italic=True, size=10, color="555555")
ws["A20"].value = foot
ws["A20"].font = Font(size=9, color="888888")
wb.save(w_path)
pdf_doc(d / "weighted-employee-evaluation-form.pdf", "Employee evaluation form (weighted)", [
    ("fields", ["Employee name", "Job title", "Evaluator", "Review period"]),
    ("p", "Multiply each rating (1 to 5) by its weight, then add the results for an overall score out of 5."),
    ("table", ["Factor", "What good looks like", "Weight", "Rating", "Weight x rating", "Comments"],
     [[r[0], r[1], f"{r[2]}%", "", "", ""] for r in w_rows] + [["", "", "", "", "", ""]] * 2),
    ("fields", ["Overall score (out of 5)", "Overall rating"]),
    ("p", "4.5 and up Outstanding. 3.5 to 4.49 Exceeds expectations. 2.5 to 3.49 Meets expectations. "
          "1.5 to 2.49 Needs improvement. Below 1.5 Unsatisfactory."),
    ("sign", ["Employee signature", "Evaluator signature"]),
])

# 3. Hourly and frontline employee evaluation ----------------------------------------------------
hourly_factors = ["Attendance (absences this period: ___)", "Punctuality (late clock-ins this period: ___)",
                  "Customer or patient service", "Speed and accuracy", "Safety and food or product handling",
                  "Cleanliness and opening/closing duties", "Teamwork and shift coverage",
                  "Follows timekeeping rules (clocks in and out, takes breaks)"]
both("hourly-employee-evaluation-form", "Hourly employee evaluation form", [
    ("p", "For hourly and shift employees in restaurants, retail, healthcare, warehouses and field services. "
          "Base ratings on time records and specific examples from the review period."),
    ("fields", ["Employee name", "Position", "Location", "Supervisor", "Review period", "Current hourly rate"]),
    ("p", SCALE),
    ("table", ["Area", "Rating (1 to 5)", "Example from this period"], rating_rows(hourly_factors)),
    ("fields", ["Overall rating (1 to 5)"]),
    ("h", "Next steps"),
    ("checks", ["Ready for more responsibility", "Cross-train on: ________", "Needs coaching on: ________",
                "Pay review recommended"]),
    ("fields", ["Goal for the next 90 days", "Support the supervisor will give"]),
    ("sign", ["Employee signature", "Supervisor signature"]),
])

# 4. 90-day probationary evaluation --------------------------------------------------------------
both("90-day-probationary-evaluation-form", "90-day employee evaluation", [
    ("p", "Complete before the end of the introductory period. Use it to confirm the job is a fit, to set "
          "expectations, and to record the decision."),
    ("fields", ["Employee name", "Job title", "Start date", "Manager", "Evaluation date"]),
    ("h", "Progress since hire"),
    ("table", ["Question", "Yes / Partly / No", "Comments"], [
        ["Has learned the core tasks of the job", "", ""],
        ["Completes work to the expected standard", "", ""],
        ["Works the scheduled shifts and follows call-out rules", "", ""],
        ["Follows safety, timekeeping and conduct policies", "", ""],
        ["Works well with coworkers and customers", "", ""],
        ["Asks for help and uses feedback", "", ""],
        ["Completed required training", "", ""],
    ]),
    ("fields", ["What is going well", "What needs to improve", "Training still needed"]),
    ("h", "Decision"),
    ("checks", ["Successfully completed introductory period", "Extend review by ___ days (give reasons)",
                "Not continuing in this role"]),
    ("p", "Completing an introductory period does not change at-will employment where at-will employment applies. "
          "In Montana, finishing a probationary period affects the rules for ending employment, so check state law."),
    ("sign", ["Employee signature", "Manager signature"]),
])

# 5. Employee self-evaluation form ---------------------------------------------------------------
both("employee-self-evaluation-form", "Employee self-evaluation form", [
    ("p", "Complete this before your review meeting and give it to your manager. Use specific examples "
          "and numbers where you can."),
    ("fields", ["Name", "Job title", "Manager", "Review period"]),
    ("h", "Your results"),
    ("fields", ["My three biggest accomplishments this period", "A goal I met (and how)",
                "A goal I did not fully meet (and why)"]),
    ("h", "Rate yourself"),
    ("p", SCALE),
    ("table", ["Area", "My rating (1 to 5)", "Example"], rating_rows(
        ["Quality of my work", "Meeting deadlines and schedules", "Job knowledge", "Communication",
         "Teamwork", "Problem solving"])),
    ("h", "Looking ahead"),
    ("fields", ["Skill I want to build", "Support or training I need", "Goals for the next period",
                "Anything that gets in the way of doing my best work"]),
    ("sign", ["Employee signature"]),
])

# 6. Manager evaluation (upward feedback) --------------------------------------------------------
both("manager-evaluation-form", "Manager evaluation form (upward feedback)", [
    ("p", "Employees use this form to give feedback on their manager. Responses are collected by HR or the "
          "owner and shared with the manager as a summary, without names."),
    ("fields", ["Manager being evaluated", "Department or location", "Period covered"]),
    ("p", "Rate each statement: 5 = Strongly agree, 4 = Agree, 3 = Neutral, 2 = Disagree, 1 = Strongly disagree."),
    ("table", ["Statement", "Rating (1 to 5)", "Comment"], rating_rows([
        "Posts schedules early enough for me to plan", "Treats everyone fairly and consistently",
        "Gives me clear expectations", "Gives useful feedback, good and bad", "Handles time off and shift "
        "swap requests promptly", "Is available when I need help", "Recognizes good work",
        "I would recommend working for this manager"])),
    ("fields", ["One thing this manager should keep doing", "One thing this manager should change"]),
    ("p", "Do not include your name. Do not use this form to report harassment or a safety issue; report those "
          "directly as described in the employee handbook."),
])

# 7. Peer evaluation form ------------------------------------------------------------------------
both("peer-evaluation-form", "Peer evaluation form", [
    ("p", "Coworkers complete this form as part of a 360-degree or team review. Comment only on work you have "
          "seen yourself."),
    ("fields", ["Employee being reviewed", "Your relationship (same team, other team, shift lead)",
                "How often you work together"]),
    ("p", "Rate each area: 5 = Always, 4 = Usually, 3 = Sometimes, 2 = Rarely, 1 = Never. Leave blank if you "
          "have not seen it."),
    ("table", ["This coworker...", "Rating (1 to 5)", "Example"], rating_rows([
        "Does their share of the work", "Shares information and helps when busy", "Communicates clearly",
        "Is dependable for shifts and handoffs", "Treats coworkers with respect", "Is open to feedback"])),
    ("fields", ["What this coworker does best", "One suggestion for this coworker"]),
    ("checks", ["I am comfortable having my name shared", "Keep my name confidential"]),
    ("fields", ["Your name (optional)"]),
])

# 8. Team evaluation summary (Excel) -------------------------------------------------------------
t_headers = ["Employee", "Position", "Review date", "Quality", "Reliability", "Teamwork", "Customer service",
             "Job knowledge", "Average", "Overall rating", "Next review"]
t_path = d / "team-evaluation-summary.xlsx"
xlsx_sheet(t_path, title="Team evaluation summary",
           subtitle="Location: ____________   Manager: ____________   Ratings 1 to 5",
           headers=t_headers, rows=[], blank_rows=25,
           widths=[22, 16, 13, 10, 11, 10, 12, 12, 10, 20, 13],
           formulas={"Average": '=IF(COUNT(D{r}:H{r})=0,"",ROUND(AVERAGE(D{r}:H{r}),2))',
                     "Overall rating": '=IF(I{r}="","",IF(I{r}>=4.5,"Outstanding",IF(I{r}>=3.5,"Exceeds",'
                                       'IF(I{r}>=2.5,"Meets",IF(I{r}>=1.5,"Needs improvement","Unsatisfactory")))))',
                     "Next review": '=IF(C{r}="","",EDATE(C{r},12))'},
           sheet_name="Team summary", landscape=True,
           notes="Enter the review date and five ratings. Average, overall rating and the next review date "
                 "(12 months later) fill in automatically. Change 12 to 6 for twice-a-year reviews.")
wb = load_workbook(t_path)
ws = wb.active
for r in range(5, 30):
    ws[f"C{r}"].number_format = "mm/dd/yyyy"
    ws[f"K{r}"].number_format = "mm/dd/yyyy"
dv = DataValidation(type="whole", operator="between", formula1="1", formula2="5", allow_blank=True,
                    showErrorMessage=True, error="Enter a whole number from 1 to 5.")
ws.add_data_validation(dv)
dv.add("D5:H29")
wb.save(t_path)
pdf_table(d / "team-evaluation-summary.pdf", "Team evaluation summary", t_headers, blank_rows=18,
          subtitle="Location: ____________   Manager: ____________   Ratings 1 to 5", landscape=True,
          notes="Average = mean of the five ratings. 4.5+ Outstanding, 3.5+ Exceeds, 2.5+ Meets, "
                "1.5+ Needs improvement, below 1.5 Unsatisfactory.")

print("built", sorted(p.name for p in d.iterdir()))
