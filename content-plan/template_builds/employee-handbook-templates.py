#!/usr/bin/env python3
"""Build the downloads for /templates/employee-handbook-templates/.

Run from the site root:  python content-plan/template_builds/employee-handbook-templates.py
Files go to assets/templates/employee-handbook-templates/. Handbooks contain sample policy wording with
[bracketed] blanks to replace; the filled-in examples live in the page previews.
"""
import sys
import pathlib

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent.parent))
from openpyxl import load_workbook
from openpyxl.worksheet.datavalidation import DataValidation
from tplkit import out, xlsx_sheet, docx_doc, pdf_table, pdf_doc

d = out("employee-handbook-templates")


def both(name, title, blocks):
    docx_doc(d / f"{name}.docx", title, blocks)
    pdf_doc(d / f"{name}.pdf", title, blocks)


NOTE = ("How to use this template: replace every [bracketed] item, delete policies that do not apply, and add "
        "the policies your state and city require. This is sample wording, not legal advice. Have an employment "
        "attorney or HR professional in your state review the final handbook before you give it to employees.")

# ---- shared policy sections ------------------------------------------------------------------
WELCOME = [
    ("h", "Welcome"),
    ("p", "Welcome to [Company name]. This handbook explains how we work, what you can expect from us, and what we "
          "expect from you. Please read it, keep it, and ask [HR contact or owner] if anything is unclear."),
]
ABOUT = [
    ("h", "About this handbook"),
    ("p", "This handbook is a summary of our policies. It is not a contract of employment and does not promise "
          "employment for any length of time. We may change, add or remove policies at any time, and we will tell "
          "you about changes in writing. If a federal, state or local law gives you more rights than a policy in "
          "this handbook, we follow the law."),
    ("p", "At-will employment: Your employment with [Company name] is at will. You or the company may end it at any "
          "time, with or without cause or notice, as long as the reason is not unlawful. Only [owner or title] can "
          "change this, and only in a signed written agreement. [Montana employers: replace this paragraph; Montana "
          "law requires good cause to discharge after the probationary period.]"),
]
EEO = [
    ("h", "Equal employment opportunity"),
    ("p", "[Company name] makes employment decisions based on merit and business needs. We do not discriminate "
          "on the basis of race, color, religion, sex (including pregnancy, sexual orientation and gender identity), "
          "national origin, age (40 and over), disability, genetic information, military or veteran status, or any "
          "other status protected by federal, state or local law. [Add the protected categories under your state "
          "and city law.]"),
    ("h", "Reasonable accommodation"),
    ("p", "We provide reasonable accommodations for qualified employees with disabilities, for known limitations "
          "related to pregnancy, childbirth or related medical conditions, and for sincerely held religious beliefs, "
          "unless the accommodation would cause undue hardship. To ask for one, talk to [HR contact]. We will work "
          "with you to find an option that fits."),
]
HARASSMENT = [
    ("h", "Harassment, discrimination and retaliation"),
    ("p", "We do not tolerate harassment, including sexual harassment, by anyone: managers, coworkers, customers "
          "or vendors. Harassment includes unwelcome conduct based on a protected characteristic, such as slurs, "
          "jokes, unwanted touching, threats, or displaying offensive images."),
    ("p", "How to report: Tell your manager, [HR contact], or [owner or second contact]. You do not have to report "
          "to your direct manager, and you can report in person, by phone at [number] or by email at [email]. We "
          "will investigate promptly and as confidentially as we can, and we will take corrective action where "
          "needed."),
    ("p", "No retaliation: We will not retaliate against anyone who reports a concern in good faith or takes "
          "part in an investigation. Report any retaliation the same way."),
]
CLASSES = [
    ("h", "Employment categories"),
    ("bullets", ["Full-time: regularly scheduled [30 / 32 / 40] or more hours a week.",
                 "Part-time: regularly scheduled fewer than [30 / 32 / 40] hours a week.",
                 "Temporary or seasonal: hired for a set period or project.",
                 "Non-exempt: paid for every hour worked and eligible for overtime.",
                 "Exempt: paid a salary and not eligible for overtime, because the job meets the legal salary "
                 "and duties tests. Your offer letter states your category."]),
]
TIME = [
    ("h", "Work hours, schedules and timekeeping"),
    ("p", "Our workweek runs from [Sunday 12:00 a.m. to Saturday 11:59 p.m.]. Schedules are posted [day and how far "
          "ahead] in [the scheduling app / on the board]. Non-exempt employees must clock in at the start of each "
          "shift, out and in for unpaid meal periods, and out at the end of the shift, using [the time clock app or "
          "kiosk]."),
    ("p", "Never work off the clock. Do not answer work calls, texts or emails, or do any work, without recording "
          "the time. If you miss a punch or your time record is wrong, tell your manager within [24 hours] so it "
          "can be corrected. No one may change your time record without telling you."),
    ("h", "Overtime"),
    ("p", "Non-exempt employees are paid 1.5 times their regular rate for all hours worked over 40 in a workweek, "
          "and any additional overtime required by state law. [California, for example, also requires daily "
          "overtime.] Overtime must be approved in advance by your manager, but all overtime worked is paid."),
    ("h", "Meal and rest breaks"),
    ("p", "Short rest breaks of [10] minutes are paid and count as hours worked. Meal periods of [30] minutes or "
          "more are unpaid only if you are fully relieved of duty; if you are asked to work during a meal period, "
          "it is paid. [Add your state's meal and rest break rules, and any lactation break policy.]"),
]
PAY = [
    ("h", "Pay"),
    ("p", "Payday is every [other Friday / weekday], for the pay period ending [date or day]. Pay is by [direct "
          "deposit or paycheck]. Your pay stub shows hours, rate, gross pay, deductions and net pay, and is "
          "available [in the app / with your check]. We only make deductions required by law or that you authorize "
          "in writing. If your pay looks wrong, tell [payroll contact] right away and we will correct any error "
          "promptly."),
    ("p", "You may talk about your pay, benefits and working conditions with coworkers and others. Nothing in this "
          "handbook limits that right."),
]
PTO = [
    ("h", "Paid time off"),
    ("p", "Eligible employees earn [number] hours of paid time off per [pay period / hour worked], up to [number] "
          "hours per year. PTO can be used for vacation, personal needs or illness. Request planned time off at "
          "least [two weeks] ahead through [the app / the time off request form]. Requests are approved based on "
          "staffing and the order they were received. Unused PTO [carries over up to (number) hours / see state "
          "law] and is [paid / not paid] at separation, as required by state law."),
    ("h", "Sick leave"),
    ("p", "[Company name] provides paid sick leave as required by [state or city law], or [number] hours per year, "
          "whichever is greater. You may use it for your own or a family member's illness, care or medical "
          "appointments, and for other reasons the law allows. We will not retaliate against you for using sick "
          "leave."),
    ("h", "Holidays"),
    ("p", "We observe these paid holidays: [New Year's Day, Memorial Day, Independence Day, Labor Day, Thanksgiving "
          "Day, Christmas Day]. When a holiday falls on a Saturday it is observed on [Friday]; on a Sunday, on "
          "[Monday]."),
    ("h", "Leaves of absence"),
    ("bullets", ["Family and medical leave: [If you have 50 or more employees, insert your FMLA policy here. Covered "
                 "employers with eligible employees must include the FMLA general notice in the handbook.] Eligible "
                 "employees may take up to 12 workweeks of unpaid, job-protected leave in a 12-month period for "
                 "qualifying reasons.",
                 "Military leave: we follow USERRA and state law for employees who serve in the uniformed services.",
                 "Jury duty and voting: we give time off for jury service and, where state law requires, for voting. "
                 "[State whether this time is paid.]",
                 "Other leave required by state or local law: [list, for example paid family leave or crime victim "
                 "leave]."]),
]
CONDUCT = [
    ("h", "Attendance"),
    ("p", "Come to work on time for every scheduled shift. If you will be late or absent, contact [manager] at "
          "least [2 hours] before your shift by [phone / app]. Missing a shift without notice is a no call, no "
          "show. Absences protected by law, such as approved FMLA leave or legally protected sick leave, are never "
          "counted against you."),
    ("h", "Conduct and discipline"),
    ("p", "We expect honesty, respect and safe work. Examples of unacceptable conduct include theft, violence or "
          "threats, harassment, working under the influence of alcohol or drugs, falsifying time records, and "
          "repeated attendance problems. We may use coaching, written warnings or other steps, but we are not "
          "required to follow any set order, and employment remains at will where at-will employment applies."),
    ("h", "Safety"),
    ("p", "Follow all safety rules and training. Report every injury, near miss or unsafe condition to [manager] "
          "right away. We carry workers' compensation insurance. We will not retaliate against anyone who reports "
          "an injury or a safety concern."),
    ("h", "Phones, technology and social media"),
    ("p", "Use company systems for work. Do not share customer or confidential business information. Nothing in "
          "this policy prevents you from discussing wages, hours or working conditions, or from talking with a "
          "government agency."),
]
SEPARATION = [
    ("h", "Leaving the company"),
    ("p", "We ask for [two weeks'] written notice if you resign. Please return keys, uniforms, equipment and "
          "badges by your last day. Your final paycheck will be paid by the deadline in [state] law, including any "
          "accrued PTO the law or this policy requires us to pay."),
]


def ack(company="[Company name]"):
    return [
        ("h", "Employee handbook acknowledgment"),
        ("p", f"I have received the {company} employee handbook dated [date]. I understand it is my responsibility "
              "to read it and to ask questions about anything I do not understand. I understand the handbook is "
              "not a contract of employment, that the company may change its policies, and that my employment is "
              "at will where state law allows. Nothing in the handbook limits my rights under the National Labor "
              "Relations Act or other laws."),
        ("fields", ["Employee name (print)"]),
        ("sign", ["Employee signature"]),
    ]


# 1. Small business employee handbook -----------------------------------------------------------
full = ([("p", NOTE)] + WELCOME + ABOUT + EEO + HARASSMENT + CLASSES + TIME + PAY + PTO + CONDUCT + SEPARATION
        + ack())
both("small-business-employee-handbook-template", "[Company name] Employee Handbook", full)

# 2. Employee handbook outline ------------------------------------------------------------------
outline = [
    ["1. Welcome and company story", "Recommended", "Who you are and what you value"],
    ["2. About this handbook and at-will statement", "Recommended", "Not a contract; policies can change; Montana differs"],
    ["3. Equal employment opportunity", "Recommended", "List federal, state and local protected categories"],
    ["4. Harassment, discrimination and retaliation", "Required in writing in some states", "e.g. California, New York; give two ways to report"],
    ["5. Reasonable accommodation", "Recommended", "Disability, pregnancy, religion"],
    ["6. Employment categories", "Recommended", "Full/part-time, exempt/non-exempt"],
    ["7. Work hours, timekeeping, overtime", "Recommended", "Workweek, no off-the-clock work, overtime rule"],
    ["8. Meal and rest breaks", "State law decides", "Federal law does not require breaks; many states do"],
    ["9. Pay and payday", "Notice required in some states", "Payday, deductions, pay stubs, fixing errors"],
    ["10. PTO, sick leave and holidays", "Sick leave required in many states and cities", "Accrual, requests, carryover, payout"],
    ["11. Leaves of absence (FMLA, military, jury)", "FMLA notice required if covered", "Covered employers with handbooks must include it"],
    ["12. Attendance and call-out rules", "Recommended", "How and when to call out"],
    ["13. Conduct and discipline", "Recommended", "Examples, not an exhaustive list"],
    ["14. Safety and injury reporting", "Recommended", "Report injuries; no retaliation"],
    ["15. Technology, privacy and social media", "Recommended", "Protect the right to discuss pay and conditions"],
    ["16. Benefits", "If offered", "Health, retirement, discounts"],
    ["17. Leaving the company", "Recommended", "Notice, property, final pay per state law"],
    ["18. Acknowledgment form", "Recommended", "Signed receipt kept in the personnel file"],
]
both("employee-handbook-outline", "Employee handbook outline", [
    ("p", "Use this outline to plan your handbook or check an existing one. 'Required' items depend on your size and "
          "state; confirm with your state labor department."),
    ("table", ["Section", "Status", "Notes"], outline),
    ("fields", ["Our state(s)", "Number of employees", "Handbook owner", "Target date"]),
])

# 3. Restaurant employee handbook ---------------------------------------------------------------
RESTAURANT = [
    ("h", "Tips and tip pools"),
    ("p", "Tips belong to the employees who receive them. Managers, supervisors and owners do not keep any part of "
          "employees' tips, including from a tip pool. [If you take a tip credit:] Tipped employees are paid a cash "
          "wage of $[amount] an hour and we take a tip credit of $[amount] an hour, as allowed by [federal / state] "
          "law. If your cash wage plus tips does not reach the full minimum wage in a workweek, we pay the "
          "difference. [Describe your tip pool: who is in it, how it is shared, and how often.]"),
    ("p", "Report all tips on [the POS / tip report] at the end of each shift. Credit card tips are paid "
          "[on the next payday / daily]."),
    ("h", "Food safety"),
    ("bullets", ["Wash hands before starting work, after breaks, after handling raw food and after using the restroom.",
                 "Get a food handler card within [number] days of hire where [state or county] requires one.",
                 "Do not work while you have vomiting, diarrhea, jaundice, a sore throat with fever, or an infected "
                 "cut; tell the manager on duty.",
                 "Follow temperature logs, labeling and allergy procedures every shift."]),
    ("h", "Alcohol service"),
    ("p", "Only employees who meet [state] age rules and have completed [responsible service training] may serve "
          "alcohol. Check ID for anyone who appears under [age]. Refuse service to anyone visibly intoxicated."),
    ("h", "Schedules, shift swaps and staff meals"),
    ("p", "Schedules are posted [day] for the following week. You may swap a shift with a qualified coworker "
          "through [the app] if a manager approves it. Employees on shift may have [one staff meal at (price or "
          "free)]."),
    ("h", "Uniforms"),
    ("p", "Wear [uniform details] and non-slip shoes. [Company name] provides [items]. Any uniform cost you pay will "
          "never reduce your pay below the minimum wage or reduce overtime pay."),
    ("h", "Employees under 18"),
    ("p", "We follow federal and state child labor laws, including limits on hours and on tasks such as using slicers "
          "and fryers, for employees under 18."),
]
both("restaurant-employee-handbook-template", "[Restaurant name] Employee Handbook",
     [("p", NOTE)] + WELCOME + ABOUT + EEO[:2] + HARASSMENT + TIME + PAY + RESTAURANT + PTO[:4] + CONDUCT
     + SEPARATION + ack("[Restaurant name]"))

# 4. Retail employee handbook -------------------------------------------------------------------
RETAIL = [
    ("h", "Cash handling"),
    ("bullets", ["Count your drawer with a manager at the start and end of each shift.",
                 "Never share register logins or leave a drawer open and unattended.",
                 "Report any shortage or overage over $[amount] to the manager on duty.",
                 "[Company name] does not deduct register shortages from employees' pay."]),
    ("h", "Loss prevention"),
    ("p", "Report suspected shoplifting to the manager on duty. Do not chase, touch or detain a customer. Employee "
          "purchases must be rung up by another employee, and personal bags [are stored in the back room]."),
    ("h", "Customer service"),
    ("p", "Greet every customer within [time], know our return policy, and call a manager for any refund over "
          "$[amount] or any upset customer you cannot help."),
    ("h", "Schedules and the holiday season"),
    ("p", "Schedules are posted at least [number] days ahead. [If you are covered by a predictive scheduling law, "
          "such as Oregon's for retail, hospitality and food service employers with 500 or more employees, or a city "
          "law, state the notice period and premium pay rules here.] From [date] to [date], time off requests are "
          "limited to [rule]."),
    ("h", "Employee discount and dress code"),
    ("p", "Employees receive [percent]% off regular-price merchandise for personal use. Wear [dress code] and your "
          "name badge on every shift."),
]
both("retail-employee-handbook-template", "[Store name] Employee Handbook",
     [("p", NOTE)] + WELCOME + ABOUT + EEO[:2] + HARASSMENT + TIME + PAY + RETAIL + PTO[:4] + CONDUCT
     + SEPARATION + ack("[Store name]"))

# 5. Time off and leave policy section ----------------------------------------------------------
both("time-off-and-leave-policy-template", "Time off and leave policy (handbook section)", [
    ("p", NOTE),
    ("h", "Paid time off (PTO)"),
    ("table", ["Employee type", "PTO earned", "Yearly maximum", "Carryover limit"], [
        ["Full-time", "[hours] per pay period", "[hours]", "[hours]"],
        ["Part-time", "[1 hour per (number) hours worked]", "[hours]", "[hours]"],
        ["Temporary or seasonal", "[sick leave only, as required by law]", "", ""],
    ]),
    ("p", "PTO starts accruing on your first day and may be used after [number] days, unless state law allows "
          "earlier use. Request planned PTO at least [two weeks] ahead. PTO is paid at your regular rate and does not "
          "count as hours worked for overtime. Unused PTO is [paid / not paid] at separation, as state law requires."),
    ("h", "Sick leave"),
    ("p", "Employees earn paid sick leave at the rate required by [state or city], or [rate], whichever is greater. "
          "Sick leave may be used for your own or a family member's illness, preventive care and other reasons the "
          "law allows. Give notice as soon as you can. We only ask for documentation where the law permits it."),
    ("h", "Holidays"),
    ("table", ["Holiday", "Date observed this year"], [["New Year's Day", ""], ["Memorial Day", ""],
                                                     ["Independence Day", ""], ["Labor Day", ""],
                                                     ["Thanksgiving Day", ""], ["Christmas Day", ""]]),
    ("p", "Non-exempt employees who work on a company holiday are paid [rate]. Holidays that fall during PTO are "
          "paid as holiday pay, not PTO."),
    ("h", "Family and medical leave (FMLA)"),
    ("p", "[For employers with 50 or more employees.] Eligible employees may take up to 12 workweeks of unpaid, "
          "job-protected leave in a 12-month period for their own serious health condition, to care for a spouse, "
          "child or parent with a serious health condition, for the birth or placement of a child, or for certain "
          "military family needs. [Insert the full FMLA general notice from the Department of Labor here.]"),
    ("h", "Other leave"),
    ("bullets", ["Bereavement: [number] paid days for [family members].", "Jury duty: [paid / unpaid] for [days].",
                 "Military leave: as required by USERRA and state law.", "Voting: as required by state law.",
                 "Other leave required by state or local law: [list]."]),
])

# 6. Remote and hybrid work policy --------------------------------------------------------------
both("remote-work-policy-template", "Remote and hybrid work policy (handbook section)", [
    ("p", NOTE),
    ("h", "Who may work remotely"),
    ("p", "Remote or hybrid work is available for [roles] with manager approval. Arrangements are reviewed [every six "
          "months] and may change based on business needs."),
    ("h", "Schedule and availability"),
    ("p", "Core hours are [time] to [time] [time zone]. Be reachable by [chat / phone] during core hours and respond "
          "within [time]. Hybrid employees work on site on [days]."),
    ("h", "Timekeeping for non-exempt employees"),
    ("p", "Record all time worked, including time spent on email or calls outside your schedule. Do not work off the "
          "clock. Overtime must be approved in advance, but all overtime worked is paid."),
    ("h", "Equipment and expenses"),
    ("p", "[Company name] provides [laptop, monitor, headset]. We reimburse approved business expenses, including "
          "[internet / phone stipend of $amount], as required by state law. Return all equipment within [number] "
          "days of leaving."),
    ("h", "Work location and taxes"),
    ("p", "Tell [HR] at least [30 days] before you change where you work, including moving to another state. Your work "
          "location affects payroll taxes and which state's employment laws apply."),
    ("h", "Security and safety"),
    ("bullets", ["Use a secure, password-protected network; no public Wi-Fi without the VPN.",
                 "Lock your screen when you step away and keep customer data off personal devices.",
                 "Keep your workspace safe and report any work-related injury the same day."]),
    ("sign", ["Employee signature", "Manager approval"]),
])

# 7. Handbook acknowledgment form ---------------------------------------------------------------
both("employee-handbook-acknowledgment-form", "Employee handbook acknowledgment form", [
    ("fields", ["Company", "Handbook version or date"]),
    ("p", "I acknowledge that I have received a copy of the employee handbook, or been told how to access it "
          "electronically. I understand that it is my responsibility to read it, follow its policies, and ask "
          "[HR contact] about anything I do not understand."),
    ("p", "I understand that the handbook is not a contract of employment and that the company may change, add or "
          "remove policies, with written notice to employees. I understand that my employment is at will, meaning "
          "that either I or the company may end it at any time, except where state law or a signed written agreement "
          "says otherwise."),
    ("p", "I understand that nothing in the handbook prevents me from discussing my wages, hours or working "
          "conditions, or from filing a complaint with a government agency."),
    ("fields", ["Employee name (print)", "Job title"]),
    ("sign", ["Employee signature", "Received by (company representative)"]),
])

# 8. Handbook policy review tracker (Excel) -----------------------------------------------------
r_headers = ["Section", "Policy owner", "Applies to (states or locations)", "Last reviewed", "Next review due",
             "Status", "What changed / notes"]
r_rows = [[row[0], "", "", "", "", "", ""] for row in outline]
r_path = d / "handbook-policy-review-tracker.xlsx"
xlsx_sheet(r_path, title="Employee handbook policy review tracker",
           subtitle="Review every section at least once a year and whenever a federal, state or local law changes.",
           headers=r_headers, rows=r_rows, blank_rows=6, widths=[42, 16, 24, 13, 15, 14, 44],
           formulas={"Next review due": '=IF(D{r}="","",EDATE(D{r},12))',
                     "Status": '=IF(D{r}="",IF(A{r}="","","Not reviewed"),IF(E{r}<TODAY(),"Overdue",'
                               'IF(E{r}-TODAY()<=60,"Due soon","Current")))'},
           sheet_name="Review tracker", landscape=True,
           notes="Next review = last reviewed + 12 months. Status shows Due soon within 60 days and Overdue after the "
                 "due date. Common triggers for an early review: January 1 and July 1 minimum wage and leave law "
                 "changes, opening in a new state, and passing 15, 20 or 50 employees.")
wb = load_workbook(r_path)
ws = wb.active
last = 4 + len(r_rows) + 6
for r in range(5, last + 1):
    ws[f"D{r}"].number_format = "mm/dd/yyyy"
    ws[f"E{r}"].number_format = "mm/dd/yyyy"
wb.save(r_path)
pdf_table(d / "handbook-policy-review-tracker.pdf", "Employee handbook policy review tracker",
          ["Section", "Policy owner", "Applies to", "Last reviewed", "Next review due", "Notes"],
          [[row[0], "", "", "", "", ""] for row in outline], blank_rows=2,
          subtitle="Review every section at least once a year and whenever a law changes.", landscape=True)

print("built", sorted(p.name for p in d.iterdir()))
