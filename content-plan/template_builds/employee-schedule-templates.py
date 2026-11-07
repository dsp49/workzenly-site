#!/usr/bin/env python3
"""Build the downloads for /templates/employee-schedule-templates/.

Run from the site root:  python content-plan/template_builds/employee-schedule-templates.py
Files go to assets/templates/employee-schedule-templates/. Files are blank and ready to fill in, with working
formulas; the filled-in examples live in the page previews.
"""
import sys
import pathlib
import datetime

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent.parent))
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.datavalidation import DataValidation
from tplkit import out, xlsx_sheet, docx_doc, pdf_table, pdf_doc, _pdf, _esc, FOOT

THIN = Side(style="thin", color="C9C9C3")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
DARK = PatternFill("solid", fgColor="161616")
GREY = PatternFill("solid", fgColor="E9E9E3")
ZEBRA = PatternFill("solid", fgColor="F6F6F3")
CALC = PatternFill("solid", fgColor="F2FFD9")   # cells with formulas
INPUT = PatternFill("solid", fgColor="FFF8D6")  # cells the user types into at the top


def new_sheet(title, subtitle=None, sheet_name=None, wb=None):
    if wb is None:
        wb = Workbook()
        ws = wb.active
    else:
        ws = wb.create_sheet()
    ws.title = (sheet_name or title)[:31]
    ws["A1"] = title
    ws["A1"].font = Font(bold=True, size=16)
    if subtitle:
        ws["A2"] = subtitle
        ws["A2"].font = Font(size=11, color="444444")
    return wb, ws


def hdr(ws, r, c, text, merge_to=None):
    cell = ws.cell(row=r, column=c, value=text)
    cell.font = Font(bold=True, color="FFFFFF")
    cell.fill = DARK
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = BOX
    if merge_to:
        for cc in range(c + 1, merge_to + 1):
            x = ws.cell(row=r, column=cc)
            x.fill, x.border = DARK, BOX
        ws.merge_cells(start_row=r, start_column=c, end_row=r, end_column=merge_to)
    return cell


def put(ws, r, c, value=None, fmt=None, fill=None, bold=False, center=False, zebra=False):
    cell = ws.cell(row=r, column=c, value=value)
    cell.border = BOX
    cell.alignment = Alignment(vertical="center", wrap_text=True, horizontal="center" if center else None)
    if fmt:
        cell.number_format = fmt
    if fill is not None:
        cell.fill = fill
    elif zebra:
        cell.fill = ZEBRA
    if bold:
        cell.font = Font(bold=True)
    return cell


def label(ws, r, c, text, value=None, vcol=None, fmt=None):
    """A bold label with an input cell next to it (used for start dates, rates and budgets)."""
    a = ws.cell(row=r, column=c, value=text)
    a.font = Font(bold=True)
    a.alignment = Alignment(horizontal="right", vertical="center")
    v = ws.cell(row=r, column=vcol or c + 1, value=value)
    v.fill, v.border = INPUT, BOX
    if fmt:
        v.number_format = fmt
    return v


def note(ws, r, text):
    ws.cell(row=r, column=1, value=text).font = Font(italic=True, size=10, color="555555")


def finish(ws, widths, header_row, foot_row, landscape=True, freeze=None):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[L(i)].width = w
    ws.cell(row=foot_row, column=1, value=FOOT).font = Font(size=9, color="888888")
    ws.freeze_panes = freeze or ws.cell(row=header_row + 1, column=1)
    ws.page_setup.orientation = "landscape" if landscape else "portrait"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = f"{header_row}:{header_row}"


def choices(ws, rng, options):
    dv = DataValidation(type="list", formula1='"' + ",".join(options) + '"', allow_blank=True)
    dv.error = "Pick a value from the list or clear the cell."
    ws.add_data_validation(dv)
    dv.add(rng)


def hours_term(a, b):
    """Excel: hours between two time cells, overnight-safe, 0 when either is blank."""
    return f"IF(COUNT({a},{b})=2,MOD({b}-{a},1),0)"


def html_pdf(path, title, body_html, subtitle=None, landscape=True, extra_css=""):
    body = f"<style>{extra_css}</style><h1>{_esc(title)}</h1>"
    if subtitle:
        body += f"<p class='sub'>{_esc(subtitle)}</p>"
    return _pdf(path, body + body_html, landscape)


def grid_html(headers, rows, cls=""):
    h = "<table class='%s'><thead><tr>" % cls + "".join(f"<th>{_esc(x)}</th>" for x in headers) + "</tr></thead><tbody>"
    h += "".join("<tr>" + "".join(f"<td>{_esc(v)}</td>" for v in r) + "</tr>" for r in rows)
    return h + "</tbody></table>"


def report(d):
    for f in sorted(d.iterdir()):
        print(f"{f.name:55s} {f.stat().st_size:>8,d} bytes")


# ============================================================== /templates/employee-schedule-templates/
from openpyxl.formatting.rule import CellIsRule, FormulaRule

d = out("employee-schedule-templates")
DAYS = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
WEEK = datetime.date(2026, 11, 16)


def day_headers(ws, r, c0, n=7, start_cell="$B$3"):
    for k in range(n):
        hdr(ws, r, c0 + k, f'=TEXT({start_cell}+{k},"ddd m/d")')


def grey_row(ws, r, c_from, c_to, text=None):
    for c in range(c_from, c_to + 1):
        put(ws, r, c, text if c == c_from else "", fill=GREY, bold=c == c_from)


# 1. Two-week employee schedule (biweekly pay period) -------------------------------------
wb, ws = new_sheet("Two-week employee schedule", "Enter scheduled hours per day as numbers. Overtime is checked for each workweek separately.",
                   sheet_name="Two weeks")
label(ws, 3, 1, "First day of the pay period", WEEK, vcol=2, fmt="m/d/yyyy")
hr = 5
hdr(ws, hr, 1, "Employee")
hdr(ws, hr, 2, "Role")
day_headers(ws, hr, 3, 14)
for j, t in enumerate(["Week 1 hours", "Week 2 hours", "Overtime week 1", "Overtime week 2", "Pay period hours"]):
    hdr(ws, hr, 17 + j, t)
r0, n = hr + 1, 20
for i in range(n):
    r = r0 + i
    put(ws, r, 1, "", zebra=i % 2)
    put(ws, r, 2, "", zebra=i % 2)
    for k in range(14):
        put(ws, r, 3 + k, None, center=True, zebra=i % 2)
    put(ws, r, 17, f"=SUM(C{r}:I{r})", fill=CALC, center=True)
    put(ws, r, 18, f"=SUM(J{r}:P{r})", fill=CALC, center=True)
    put(ws, r, 19, f"=MAX(0,Q{r}-40)", fill=CALC, center=True)
    put(ws, r, 20, f"=MAX(0,R{r}-40)", fill=CALC, center=True)
    put(ws, r, 21, f"=Q{r}+R{r}", fill=CALC, center=True)
rl = r0 + n - 1
grey_row(ws, rl + 1, 1, 2, "Hours per day")
for c in range(3, 22):
    put(ws, rl + 1, c, f"=SUM({L(c)}{r0}:{L(c)}{rl})", fill=CALC, center=True, bold=True)
ws.conditional_formatting.add(f"S{r0}:T{rl}", CellIsRule(operator="greaterThan", formula=["0"],
                                                       fill=PatternFill("solid", fgColor="F8C9C0")))
note(ws, rl + 3, "The FLSA looks at each workweek on its own: 45 hours in week 1 and 35 in week 2 is 5 hours of overtime, "
                 "even though the pay period totals 80. Set the first day to the first day of your workweek.")
finish(ws, [22, 13] + [7.5] * 14 + [9, 9, 10, 10, 10], hr, rl + 5, freeze="C6")
wb.save(d / "two-week-employee-schedule.xlsx")
pdf_table(d / "two-week-employee-schedule.pdf", "Two-week employee schedule",
          ["Employee"] + [f"{DAYS[k % 7]}" for k in range(14)] + ["Wk 1", "Wk 2", "OT"], blank_rows=16,
          subtitle="Pay period: ____________ to ____________   Write hours per day; overtime is over 40 in each week.")

# 2. Restaurant employee schedule (FOH / BOH) ----------------------------------------------
wb, ws = new_sheet("Restaurant employee schedule", "Restaurant: ____________   Manager: ____________", sheet_name="Restaurant")
label(ws, 3, 1, "Week starting", WEEK, vcol=2, fmt="m/d/yyyy")
hr = 5
for c, t in enumerate(["Employee", "Position", "Under 18?"], 1):
    hdr(ws, hr, c, t)
day_headers(ws, hr, 4)
hdr(ws, hr, 11, "Scheduled hours")
r = hr + 1
section_rows = {}
for sec, n in [("Front of house (servers, hosts, bartenders, bussers)", 12), ("Back of house (cooks, prep, dish)", 10)]:
    grey_row(ws, r, 1, 11, sec)
    r += 1
    start = r
    for i in range(n):
        put(ws, r, 1, "", zebra=i % 2)
        put(ws, r, 2, "", zebra=i % 2)
        put(ws, r, 3, "", center=True, zebra=i % 2)
        for k in range(7):
            put(ws, r, 4 + k, "", center=True, zebra=i % 2)
        put(ws, r, 11, None, center=True, zebra=i % 2)
        r += 1
    section_rows[sec] = (start, r - 1)
    put(ws, r, 1, "Section hours", bold=True, fill=GREY)
    for c in range(2, 11):
        put(ws, r, c, "", fill=GREY)
    put(ws, r, 11, f"=SUM(K{start}:K{r - 1})", fill=CALC, center=True, bold=True)
    r += 1
(fa, fb), (ba, bb) = section_rows.values()
put(ws, r, 1, "Total scheduled labor hours", bold=True, fill=GREY)
for c in range(2, 11):
    put(ws, r, c, "", fill=GREY)
put(ws, r, 11, f"=SUM(K{fa}:K{fb})+SUM(K{ba}:K{bb})", fill=CALC, center=True, bold=True)
choices(ws, f"C{fa}:C{bb}", ["Y", "N"])
ws.conditional_formatting.add(f"A{fa}:K{bb}", FormulaRule(formula=[f'$C{fa}="Y"'], fill=PatternFill("solid", fgColor="FFF1B8")))
note(ws, r + 2, "Write shifts as 10a-3p or 4p-close. Mark Under 18 = Y to highlight minors: federal law limits 14- and 15-year-olds "
                "to 3 hours on a school day, 18 hours in a school week, and work between 7 a.m. and 7 p.m. (9 p.m. from June 1 "
                "through Labor Day). States can be stricter.")
finish(ws, [22, 14, 9] + [12] * 7 + [11], hr, r + 4)
wb.save(d / "restaurant-employee-schedule.xlsx")
rows = [["FRONT OF HOUSE"] + [""] * 9] + [[""] * 10] * 9 + [["BACK OF HOUSE"] + [""] * 9] + [[""] * 10] * 7
pdf_table(d / "restaurant-employee-schedule.pdf", "Restaurant employee schedule", ["Employee", "Position"] + DAYS + ["Hours"],
          rows=rows, blank_rows=0, subtitle="Week of: ____________   Restaurant: ____________   Manager: ____________",
          notes="Mark minors with *: 14- and 15-year-olds may work 3 hours on a school day, 18 hours in a school week, "
                "7 a.m. to 7 p.m. (9 p.m. June 1 through Labor Day). Check your state's rules too.")

# 3. Retail staff schedule with open / mid / close codes -----------------------------------
wb, ws = new_sheet("Retail staff schedule", "Store: ____________   Store manager: ____________", sheet_name="Retail")
label(ws, 3, 1, "Week starting", WEEK, vcol=2, fmt="m/d/yyyy")
# shift definitions (editable)
hdr(ws, 3, 5, "Code")
hdr(ws, 3, 6, "Shift")
hdr(ws, 3, 7, "Start")
hdr(ws, 3, 8, "End")
hdr(ws, 3, 9, "Paid hours")
defs = [("O", "Open", datetime.time(8, 30), datetime.time(17, 0), 8), ("M", "Mid", datetime.time(11, 0), datetime.time(19, 30), 8),
        ("C", "Close", datetime.time(13, 30), datetime.time(22, 0), 8)]
for i, (code, name, a, b, h) in enumerate(defs):
    rr = 4 + i
    put(ws, rr, 5, code, bold=True, center=True)
    put(ws, rr, 6, name)
    put(ws, rr, 7, a, fmt="h:mm AM/PM", fill=INPUT)
    put(ws, rr, 8, b, fmt="h:mm AM/PM", fill=INPUT)
    put(ws, rr, 9, h, fill=INPUT, center=True)
hr = 8
for c, t in enumerate(["Employee", "Role", "Keyholder?"], 1):
    hdr(ws, hr, c, t)
day_headers(ws, hr, 4)
hdr(ws, hr, 11, "Shifts")
hdr(ws, hr, 12, "Hours")
r0, n = hr + 1, 18
for i in range(n):
    r = r0 + i
    put(ws, r, 1, "", zebra=i % 2)
    put(ws, r, 2, "", zebra=i % 2)
    put(ws, r, 3, "", center=True, zebra=i % 2)
    for k in range(7):
        put(ws, r, 4 + k, "", center=True, zebra=i % 2)
    rng = f"D{r}:J{r}"
    put(ws, r, 11, f'=COUNTIF({rng},"O")+COUNTIF({rng},"M")+COUNTIF({rng},"C")', fill=CALC, center=True)
    put(ws, r, 12, f'=COUNTIF({rng},"O")*$I$4+COUNTIF({rng},"M")*$I$5+COUNTIF({rng},"C")*$I$6', fill=CALC, center=True)
rl = r0 + n - 1
choices(ws, f"D{r0}:J{rl}", ["O", "M", "C", "OFF", "PTO"])
choices(ws, f"C{r0}:C{rl}", ["Y", "N"])
for j, (lab, code) in enumerate([("Openers", "O"), ("Mids", "M"), ("Closers", "C")]):
    rr = rl + 1 + j
    grey_row(ws, rr, 1, 3, lab)
    for k in range(7):
        c = L(4 + k)
        put(ws, rr, 4 + k, f'=COUNTIF({c}{r0}:{c}{rl},"{code}")', fill=CALC, center=True, bold=True)
for j, (lab, code) in enumerate([("Keyholder opening?", "O"), ("Keyholder closing?", "C")]):
    rr = rl + 4 + j
    grey_row(ws, rr, 1, 3, lab)
    for k in range(7):
        c = L(4 + k)
        put(ws, rr, 4 + k, f'=IF(COUNTIFS($C${r0}:$C${rl},"Y",{c}{r0}:{c}{rl},"{code}")>0,"Yes","NO")',
            fill=CALC, center=True, bold=True)
    ws.conditional_formatting.add(f"D{rr}:J{rr}", CellIsRule(operator="equal", formula=['"NO"'],
                                                            fill=PatternFill("solid", fgColor="F8C9C0")))
grey_row(ws, rl + 6, 1, 11, "Total scheduled hours")
put(ws, rl + 6, 12, f"=SUM(L{r0}:L{rl})", fill=CALC, center=True, bold=True)
note(ws, rl + 8, "Use O, M or C for each day (or OFF, PTO). Edit the shift times and paid hours in the yellow table at the top; "
                 "the hours column uses them. The keyholder rows turn red on any day with no keyholder opening or closing.")
finish(ws, [22, 14, 11, 10, 10, 10, 10, 10, 10, 10, 9, 9], hr, rl + 10)
wb.save(d / "retail-staff-schedule.xlsx")
pdf_table(d / "retail-staff-schedule.pdf", "Retail staff schedule", ["Employee", "Role", "Key"] + DAYS + ["Hours"],
          blank_rows=14, subtitle="Week of: ________   O = open 8:30a-5p   M = mid 11a-7:30p   C = close 1:30p-10p",
          notes="Openers per day: ___ ___ ___ ___ ___ ___ ___   Closers per day: ___ ___ ___ ___ ___ ___ ___   "
                "Check a keyholder opens and closes every day.")

# 4. Clinic staff schedule with coverage by role -------------------------------------------
wb, ws = new_sheet("Clinic staff schedule", "Clinic: ____________   Practice manager: ____________", sheet_name="Clinic")
label(ws, 3, 1, "Week starting", WEEK, vcol=2, fmt="m/d/yyyy")
hr = 5
hdr(ws, hr, 1, "Staff member")
hdr(ws, hr, 2, "Role")
day_headers(ws, hr, 3, 6)
hdr(ws, hr, 9, "Days scheduled")
ROLES = ["Provider", "RN", "MA", "Front desk"]
r0, n = hr + 1, 20
for i in range(n):
    r = r0 + i
    put(ws, r, 1, "", zebra=i % 2)
    put(ws, r, 2, "", zebra=i % 2)
    for k in range(6):
        put(ws, r, 3 + k, "", center=True, zebra=i % 2)
    rng = f"C{r}:H{r}"
    put(ws, r, 9, f'=COUNTA({rng})-COUNTIF({rng},"OFF")-COUNTIF({rng},"PTO")', fill=CALC, center=True)
rl = r0 + n - 1
choices(ws, f"B{r0}:B{rl}", ROLES)
rr = rl + 1
for role, need in zip(ROLES, [2, 1, 3, 2]):
    grey_row(ws, rr, 1, 2, f"{role}: needed")
    for k in range(6):
        put(ws, rr, 3 + k, need, fill=INPUT, center=True)
    grey_row(ws, rr + 1, 1, 2, f"{role}: scheduled")
    for k in range(6):
        c = L(3 + k)
        put(ws, rr + 1, 3 + k,
            f'=COUNTIFS($B${r0}:$B${rl},"{role}",{c}{r0}:{c}{rl},"<>")-COUNTIFS($B${r0}:$B${rl},"{role}",{c}{r0}:{c}{rl},"OFF")'
            f'-COUNTIFS($B${r0}:$B${rl},"{role}",{c}{r0}:{c}{rl},"PTO")', fill=CALC, center=True, bold=True)
    ws.conditional_formatting.add(f"C{rr + 1}:H{rr + 1}", CellIsRule(operator="lessThan", formula=[f"C{rr}"],
                                                                    fill=PatternFill("solid", fgColor="F8C9C0")))
    rr += 2
note(ws, rr + 1, "Pick each person's role from the list, then type their shift (7:30a-4p) or OFF / PTO for each day. "
                 "Change the yellow 'needed' numbers to your own minimums; a scheduled count turns red when a role is short.")
finish(ws, [22, 13] + [13] * 6 + [10], hr, rr + 3)
wb.save(d / "clinic-staff-schedule.xlsx")
pdf_table(d / "clinic-staff-schedule.pdf", "Clinic staff schedule", ["Staff member", "Role"] + DAYS[:6],
          rows=[[""] * 8] * 14 + [[f"{x}: needed / scheduled", ""] + [""] * 6 for x in ROLES], blank_rows=0,
          subtitle="Week of: ____________   Clinic: ____________")

# 5. Three-shift staff schedule (warehouse / manufacturing) --------------------------------
wb, ws = new_sheet("Three-shift staff schedule", "Facility: ____________   Supervisor: ____________", sheet_name="Three shifts")
label(ws, 3, 1, "Week starting", WEEK, vcol=2, fmt="m/d/yyyy")
label(ws, 3, 4, "Paid hours per shift", 8, vcol=6)
hr = 5
for c, t in enumerate(["Employee", "Area or line", "Shift"], 1):
    hdr(ws, hr, c, t)
day_headers(ws, hr, 4)
hdr(ws, hr, 11, "Days")
hdr(ws, hr, 12, "Hours")
r = hr + 1
blocks = []
for name, n in [("1st shift, 6:00 a.m. to 2:30 p.m.", 8), ("2nd shift, 2:00 p.m. to 10:30 p.m.", 8), ("3rd shift, 10:00 p.m. to 6:30 a.m.", 6)]:
    grey_row(ws, r, 1, 12, name)
    r += 1
    s = r
    code = name.split(",")[0]
    for i in range(n):
        put(ws, r, 1, "", zebra=i % 2)
        put(ws, r, 2, "", zebra=i % 2)
        put(ws, r, 3, code, center=True, zebra=i % 2)
        for k in range(7):
            put(ws, r, 4 + k, "", center=True, zebra=i % 2)
        rng = f"D{r}:J{r}"
        put(ws, r, 11, f'=COUNTA({rng})-COUNTIF({rng},"OFF")-COUNTIF({rng},"PTO")', fill=CALC, center=True)
        put(ws, r, 12, f"=K{r}*$F$3", fill=CALC, center=True)
        r += 1
    blocks.append((code, s, r - 1))
first, last = blocks[0][1], blocks[-1][2]
for code, s, e in blocks:
    grey_row(ws, r, 1, 3, f"{code} headcount")
    for k in range(7):
        c = L(4 + k)
        put(ws, r, 4 + k, f'=COUNTA({c}{s}:{c}{e})-COUNTIF({c}{s}:{c}{e},"OFF")-COUNTIF({c}{s}:{c}{e},"PTO")',
            fill=CALC, center=True, bold=True)
    r += 1
grey_row(ws, r, 1, 11, "Total scheduled hours")
put(ws, r, 12, "=" + "+".join(f"SUM(L{s}:L{e})" for _, s, e in blocks), fill=CALC, center=True, bold=True)
note(ws, r + 2, "In each day, type the area the person works (Receiving, Pick, Line 2) or OFF / PTO. Paid hours assume an "
                "8.5-hour shift with a 30-minute unpaid meal break; change the yellow cell if yours differ. Third shift is "
                "listed on the day it starts.")
finish(ws, [22, 16, 10] + [12] * 7 + [7, 8], hr, r + 4)
wb.save(d / "three-shift-staff-schedule.xlsx")
rows = []
for name, n in [("1ST SHIFT 6:00a-2:30p", 6), ("2ND SHIFT 2:00p-10:30p", 6), ("3RD SHIFT 10:00p-6:30a", 5)]:
    rows += [[name] + [""] * 8] + [[""] * 9] * n
pdf_table(d / "three-shift-staff-schedule.pdf", "Three-shift staff schedule", ["Employee", "Area"] + DAYS,
          rows=rows, blank_rows=0, subtitle="Week of: ____________   Facility: ____________   Write the area or OFF / PTO.")

# 6. Employee availability form ----------------------------------------------------------
AV_BLOCKS = [
    ("p", "Tell us when you can work so we can build schedules that fit. Your manager will try to schedule within these "
          "hours but cannot guarantee it. Give a new form at least two weeks before your availability changes."),
    ("fields", ["Employee name", "Position", "Phone", "Date", "Availability starts on"]),
    ("h", "Weekly availability"),
    ("table", ["Day", "Available from", "Available until", "Not available"],
     [[x, "", "", "☐"] for x in ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]]),
    ("h", "Hours"),
    ("fields", ["Most hours I want per week", "Fewest hours I need per week"]),
    ("checks", ["Open to extra shifts", "Open to closing shifts", "Open to opening shifts"]),
    ("fields", ["Other jobs or school hours to work around (optional)", "Dates I already know I cannot work"]),
    ("sign", ["Employee signature", "Manager received"]),
]
docx_doc(d / "employee-availability-form.docx", "Employee availability form", AV_BLOCKS)
pdf_doc(d / "employee-availability-form.pdf", "Employee availability form", AV_BLOCKS)

# 7. Schedule change log ----------------------------------------------------------------
hl = ["Date of change", "Employee", "Shift date", "Original shift", "New shift", "Change type", "Requested by",
      "Employee agreed in writing?", "Days of notice", "Premium pay owed?", "Recorded by"]
wb, ws = new_sheet("Schedule change log", "Location: ____________   Log every change made after the schedule is posted.",
                   sheet_name="Change log")
hr = 4
for c, t in enumerate(hl, 1):
    hdr(ws, hr, c, t)
r0, n = hr + 1, 60
for i in range(n):
    r = r0 + i
    for c in range(1, 12):
        if c in (1, 3):
            put(ws, r, c, None, fmt="m/d/yyyy", zebra=i % 2)
        elif c == 9:
            put(ws, r, c, f'=IF(COUNT(A{r},C{r})=2,C{r}-A{r},"")', fill=CALC, center=True)
        else:
            put(ws, r, c, "", zebra=i % 2)
rl = r0 + n - 1
choices(ws, f"F{r0}:F{rl}", ["Added hours", "Cut hours", "Moved shift", "Cancelled shift", "Swap", "Added shift"])
choices(ws, f"G{r0}:G{rl}", ["Employee", "Employer"])
choices(ws, f"H{r0}:H{rl}", ["Yes", "No"])
choices(ws, f"J{r0}:J{rl}", ["Yes", "No", "Check"])
note(ws, rl + 2, "Days of notice = shift date minus the date of the change. In cities and states with predictive scheduling "
                 "laws, changes made on short notice by the employer can require premium pay, while changes the employee "
                 "asks for in writing usually do not. Check your local rule.")
finish(ws, [13, 20, 12, 16, 16, 15, 13, 13, 9, 11, 13], hr, rl + 4)
wb.save(d / "schedule-change-log.xlsx")
pdf_table(d / "schedule-change-log.pdf", "Schedule change log", hl, blank_rows=18,
          subtitle="Location: ____________   Log every change made after the schedule is posted.")

# 8. Hybrid office schedule ---------------------------------------------------------------
wb, ws = new_sheet("Hybrid office schedule", "Office: ____________   Team: ____________", sheet_name="Hybrid")
label(ws, 3, 1, "Week starting", WEEK, vcol=2, fmt="m/d/yyyy")
label(ws, 3, 4, "Desks available", 12, vcol=6)
hr = 5
hdr(ws, hr, 1, "Employee")
hdr(ws, hr, 2, "Team")
day_headers(ws, hr, 3, 5)
hdr(ws, hr, 8, "Office days")
hdr(ws, hr, 9, "Remote days")
r0, n = hr + 1, 24
for i in range(n):
    r = r0 + i
    put(ws, r, 1, "", zebra=i % 2)
    put(ws, r, 2, "", zebra=i % 2)
    for k in range(5):
        put(ws, r, 3 + k, "", center=True, zebra=i % 2)
    put(ws, r, 8, f'=COUNTIF(C{r}:G{r},"Office")', fill=CALC, center=True)
    put(ws, r, 9, f'=COUNTIF(C{r}:G{r},"Remote")', fill=CALC, center=True)
rl = r0 + n - 1
choices(ws, f"C{r0}:G{rl}", ["Office", "Remote", "Off", "PTO"])
for code, color in [("Office", "FFF1B8"), ("Remote", "C9D7F2"), ("PTO", "E2E2DC")]:
    ws.conditional_formatting.add(f"C{r0}:G{rl}", CellIsRule(operator="equal", formula=[f'"{code}"'],
                                                            fill=PatternFill("solid", fgColor=color)))
for j, (lab, f) in enumerate([("In the office", "Office"), ("Remote", "Remote"), ("Desks left", None)]):
    rr = rl + 1 + j
    grey_row(ws, rr, 1, 2, lab)
    for k in range(5):
        c = L(3 + k)
        v = f'=COUNTIF({c}{r0}:{c}{rl},"{f}")' if f else f"=$F$3-{c}{rl + 1}"
        put(ws, rr, 3 + k, v, fill=CALC, center=True, bold=True)
ws.conditional_formatting.add(f"C{rl + 3}:G{rl + 3}", CellIsRule(operator="lessThan", formula=["0"],
                                                                fill=PatternFill("solid", fgColor="F8C9C0")))
note(ws, rl + 5, "Pick Office, Remote, Off or PTO for each day. 'Desks left' turns red when more people are in than you have desks.")
finish(ws, [22, 14] + [12] * 5 + [10, 10], hr, rl + 7)
wb.save(d / "hybrid-office-schedule.xlsx")
pdf_table(d / "hybrid-office-schedule.pdf", "Hybrid office schedule", ["Employee", "Team"] + DAYS[:5] + ["Office days"],
          blank_rows=18, subtitle="Week of: ____________   Write Office, Remote, Off or PTO for each day.", landscape=False)

report(d)
