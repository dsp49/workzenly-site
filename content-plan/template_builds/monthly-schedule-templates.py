#!/usr/bin/env python3
"""Build the downloads for /templates/monthly-schedule-templates/.

Run from the site root:  python content-plan/template_builds/monthly-schedule-templates.py
Files go to assets/templates/monthly-schedule-templates/. Dates fill in from the month you type; files are
otherwise blank. The filled-in examples live in the page previews.
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


# ============================================================== /templates/monthly-schedule-templates/
from openpyxl.formatting.rule import CellIsRule, FormulaRule

d = out("monthly-schedule-templates")
MONTH = datetime.date(2026, 12, 1)
RED = PatternFill("solid", fgColor="F8C9C0")
SHADES = {"D": "FFF1B8", "E": "FFD9B3", "N": "C9D7F2", "PTO": "E2E2DC", "REQ": "FFE3A3"}


def in_month(k, start="$B$3"):
    return f"MONTH({start}+{k})=MONTH({start})"


def month_headers(ws, r, c0, start="$B$3"):
    """Two header rows: weekday and day number; blank past the end of the month. Row r+1 holds the real dates."""
    for k in range(31):
        c = c0 + k
        hdr(ws, r, c, f'=IF({in_month(k, start)},TEXT({start}+{k},"ddd"),"")')
        cell = hdr(ws, r + 1, c, f'=IF({in_month(k, start)},{start}+{k},"")')
        cell.number_format = "d"


def shade(ws, rng, codes):
    for code in codes:
        ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=[f'"{code}"'],
                                                      fill=PatternFill("solid", fgColor=SHADES[code])))


def grey_row(ws, r, c_from, c_to, text=None):
    for c in range(c_from, c_to + 1):
        put(ws, r, c, text if c == c_from else "", fill=GREY, bold=c == c_from)


def roster_pdf(path, title, subtitle, n_rows, extra_cols, notes, first_col="Employee"):
    heads = [first_col] + [str(i) for i in range(1, 32)] + extra_cols
    rows = [[""] * len(heads) for _ in range(n_rows)]
    html_pdf(path, title, grid_html(heads, rows, "ro") + f"<p class='sub' style='margin-top:10px'>{_esc(notes)}</p>",
             subtitle=subtitle, landscape=True,
             extra_css=".ro th,.ro td{padding:5px 1px;font-size:7.5pt;text-align:center;height:17px}.ro td:first-child,.ro th:first-child{width:110px}")


# 1. Monthly employee schedule (roster with shift codes) ----------------------------------
wb, ws = new_sheet("Monthly employee schedule", "Codes: D = day, E = evening, N = night, OFF, PTO. Location: ____________",
                   sheet_name="Monthly roster")
label(ws, 3, 1, "First day of the month", MONTH, vcol=2, fmt="m/d/yyyy")
label(ws, 4, 1, "Paid hours per shift", 8, vcol=2)
hr = 6
hdr(ws, hr, 1, "Employee")
hdr(ws, hr, 2, "Role")
hdr(ws, hr + 1, 1, "")
hdr(ws, hr + 1, 2, "")
month_headers(ws, hr, 3)
for j, t in enumerate(["Shifts", "Hours", "PTO days"]):
    hdr(ws, hr, 34 + j, t)
    hdr(ws, hr + 1, 34 + j, "")
r0, n = hr + 2, 25
for i in range(n):
    r = r0 + i
    put(ws, r, 1, "", zebra=i % 2)
    put(ws, r, 2, "", zebra=i % 2)
    for k in range(31):
        put(ws, r, 3 + k, "", center=True)
    rng = f"C{r}:AG{r}"
    put(ws, r, 34, f'=COUNTIF({rng},"D")+COUNTIF({rng},"E")+COUNTIF({rng},"N")', fill=CALC, center=True)
    put(ws, r, 35, f"=AH{r}*$B$4", fill=CALC, center=True)
    put(ws, r, 36, f'=COUNTIF({rng},"PTO")', fill=CALC, center=True)
rl = r0 + n - 1
choices(ws, f"C{r0}:AG{rl}", ["D", "E", "N", "OFF", "PTO"])
shade(ws, f"C{r0}:AG{rl}", ["D", "E", "N", "PTO"])
for j, (lab, code) in enumerate([("Day shifts", "D"), ("Evening shifts", "E"), ("Night shifts", "N"), ("On PTO", "PTO")]):
    rr = rl + 1 + j
    grey_row(ws, rr, 1, 2, lab)
    for k in range(31):
        c = L(3 + k)
        put(ws, rr, 3 + k, f'=IF({c}${hr + 1}="","",COUNTIF({c}{r0}:{c}{rl},"{code}"))', fill=CALC, center=True)
note(ws, rl + 6, "Change the first day of the month and the dates and weekdays update; days 29 to 31 go blank in shorter "
                 "months. Monthly hours are for planning only: overtime is owed for hours over 40 in each workweek.")
finish(ws, [20, 12] + [4.6] * 31 + [7, 7, 7], hr + 1, rl + 8, freeze="C8")
wb.save(d / "monthly-employee-schedule.xlsx")
roster_pdf(d / "monthly-employee-schedule.pdf", "Monthly employee schedule",
           "Month: ____________   Location: ____________   Codes: D = day, E = evening, N = night, OFF, PTO", 18,
           ["Shifts", "Hours"], "Write a code in each day, then count each person's shifts and hours at the right. Overtime is counted per workweek, not per month.")

# 2. Monthly calendar work schedule ------------------------------------------------------
wb, ws = new_sheet("Monthly work schedule calendar", "Team or location: ____________", sheet_name="Calendar")
label(ws, 3, 1, "First day of the month", MONTH, vcol=3, fmt="mmmm yyyy")
ws.merge_cells("A3:B3")
hr = 5
for k, name in enumerate(["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]):
    hdr(ws, hr, 1 + k, name)
grid0 = "($C$3-WEEKDAY($C$3)+1)"
r = hr + 1
for w in range(6):
    for k in range(7):
        off = w * 7 + k
        dc = put(ws, r, 1 + k, f"=IF(MONTH({grid0}+{off})=MONTH($C$3),{grid0}+{off},\"\")", fmt="d", fill=GREY, bold=True)
        dc.alignment = Alignment(horizontal="left", vertical="top")
        body = put(ws, r + 1, 1 + k, "")
        body.alignment = Alignment(vertical="top", wrap_text=True)
    ws.row_dimensions[r + 1].height = 96
    r += 2
note(ws, r + 1, "Type who works each day inside the box, for example 'Open: Ana, Luis / Close: Tom, Kara'. Press Alt+Enter for a new "
                "line in a cell. Change the month in the yellow cell and the dates move.")
finish(ws, [21] * 7, hr, r + 3, landscape=True)
wb.save(d / "monthly-work-schedule-calendar.xlsx")
cal = "<table class='cal'><thead><tr>" + "".join(f"<th>{x}</th>" for x in ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]) + "</tr></thead><tbody>"
cal += "".join("<tr>" + "<td><span class='dn'>&nbsp;</span></td>" * 7 + "</tr>" for _ in range(6)) + "</tbody></table>"
html_pdf(d / "monthly-work-schedule-calendar.pdf", "Monthly work schedule", cal,
         subtitle="Month: ____________   Team or location: ____________   Write the date in each corner, then who works.",
         landscape=True, extra_css=".cal td{height:78px;width:14.28%;vertical-align:top}.dn{display:inline-block;width:22px;height:16px;border:1px solid #999}")

# 3. Monthly schedule with weekly overtime check -----------------------------------------
wb, ws = new_sheet("Monthly schedule with overtime check", "Hours per day as numbers. Five workweeks shown so every day of the month fits.",
                   sheet_name="Hours and overtime")
label(ws, 3, 1, "First day of the first workweek", datetime.date(2026, 11, 30), vcol=2, fmt="m/d/yyyy")
hr = 5
hdr(ws, hr, 1, "Employee")
hdr(ws, hr, 2, "Role")
hdr(ws, hr + 1, 1, "")
hdr(ws, hr + 1, 2, "")
for k in range(35):
    c = 3 + k
    hdr(ws, hr, c, f'=TEXT($B$3+{k},"ddd")')
    hdr(ws, hr + 1, c, f"=$B$3+{k}").number_format = "d"
wk = 38
for w in range(5):
    hdr(ws, hr, wk + w, f"Week {w + 1}")
    hdr(ws, hr + 1, wk + w, "hours")
for j, t in enumerate(["Total hours", "Overtime hours"]):
    hdr(ws, hr, 43 + j, t)
    hdr(ws, hr + 1, 43 + j, "")
r0, n = hr + 2, 20
for i in range(n):
    r = r0 + i
    put(ws, r, 1, "", zebra=i % 2)
    put(ws, r, 2, "", zebra=i % 2)
    for k in range(35):
        put(ws, r, 3 + k, None, center=True, zebra=(k // 7) % 2 == 1)
    for w in range(5):
        a, b = L(3 + 7 * w), L(9 + 7 * w)
        put(ws, r, wk + w, f"=SUM({a}{r}:{b}{r})", fill=CALC, center=True)
    put(ws, r, 43, f"=SUM(AL{r}:AP{r})", fill=CALC, center=True)
    put(ws, r, 44, "=" + "+".join(f"MAX(0,{L(wk + w)}{r}-40)" for w in range(5)), fill=CALC, center=True, bold=True)
rl = r0 + n - 1
ws.conditional_formatting.add(f"AL{r0}:AP{rl}", CellIsRule(operator="greaterThan", formula=["40"], fill=RED))
grey_row(ws, rl + 1, 1, 2, "Hours per day")
for c in range(3, 45):
    put(ws, rl + 1, c, f"=SUM({L(c)}{r0}:{L(c)}{rl})", fill=CALC, center=True, bold=True)
note(ws, rl + 3, "Set the yellow date to the first day of the workweek that contains the 1st of the month (for December 2026 with "
                 "a Monday workweek, that is November 30). Weeks over 40 hours turn red; overtime is the sum of each week's "
                 "hours over 40, which is how the FLSA counts it.")
finish(ws, [20, 12] + [4.4] * 35 + [7] * 5 + [8, 9], hr + 1, rl + 5, freeze="C7")
wb.save(d / "monthly-schedule-with-overtime-check.xlsx")
heads = ["Employee"] + [f"{['M','T','W','T','F','S','S'][k % 7]}" for k in range(35)] + ["W1", "W2", "W3", "W4", "W5", "OT"]
html_pdf(d / "monthly-schedule-with-overtime-check.pdf", "Monthly schedule with overtime check",
         grid_html(heads, [[""] * len(heads)] * 16, "ro") +
         "<p class='sub' style='margin-top:10px'>Write hours per day. Add each week's hours; overtime is hours over 40 in each "
         "workweek, not over 160 in the month.</p>",
         subtitle="Month: ____________   First day of week 1: ____________", landscape=True,
         extra_css=".ro th,.ro td{padding:5px 1px;font-size:7pt;text-align:center;height:17px}.ro td:first-child,.ro th:first-child{width:100px}")

# 4. Monthly on-call schedule ------------------------------------------------------------
hc = ["Date", "Day", "Primary on call", "Primary phone", "Backup", "On-call window", "Called in?", "Hours worked on calls", "Notes"]
wb, ws = new_sheet("Monthly on-call schedule", "Team: ____________   Escalation contact: ____________", sheet_name="On call")
label(ws, 3, 1, "First day of the month", MONTH, vcol=3, fmt="mmmm yyyy")
ws.merge_cells("A3:B3")
hr = 5
for c, t in enumerate(hc, 1):
    hdr(ws, hr, c, t)
r0 = hr + 1
for k in range(31):
    r = r0 + k
    put(ws, r, 1, f'=IF({in_month(k, "$C$3")},$C$3+{k},"")', fmt="mmm d", fill=GREY, bold=True)
    put(ws, r, 2, f'=IF(A{r}="","",TEXT(A{r},"ddd"))', fill=GREY)
    for c in range(3, 10):
        put(ws, r, c, "", zebra=k % 2, center=c in (7, 8))
rl = r0 + 30
choices(ws, f"G{r0}:G{rl}", ["Yes", "No"])
grey_row(ws, rl + 1, 1, 7, "Total hours worked on calls")
put(ws, rl + 1, 8, f"=SUM(H{r0}:H{rl})", fill=CALC, bold=True, center=True)
sr = rl + 3
hdr(ws, sr, 3, "Person")
hdr(ws, sr, 4, "Primary days")
hdr(ws, sr, 5, "Backup days")
hdr(ws, sr, 6, "Hours worked on calls")
for i in range(8):
    r = sr + 1 + i
    put(ws, r, 3, "", fill=INPUT)
    put(ws, r, 4, f'=IF(C{r}="","",COUNTIF($C${r0}:$C${rl},C{r}))', fill=CALC, center=True)
    put(ws, r, 5, f'=IF(C{r}="","",COUNTIF($E${r0}:$E${rl},C{r}))', fill=CALC, center=True)
    put(ws, r, 6, f'=IF(C{r}="","",SUMIF($C${r0}:$C${rl},C{r},$H${r0}:$H${rl}))', fill=CALC, center=True)
note(ws, sr + 10, "Under the FLSA, on-call time at the workplace counts as hours worked; on call at home generally does not, unless "
                  "the limits on the employee's time are heavy. Time spent answering calls or going in is always work time, so "
                  "record it in the hours column and on the time card.")
finish(ws, [10, 7, 20, 15, 18, 16, 10, 12, 30], hr, sr + 12, landscape=False)
wb.save(d / "monthly-on-call-schedule.xlsx")
pdf_table(d / "monthly-on-call-schedule.pdf", "Monthly on-call schedule", ["Date", "Primary on call", "Phone", "Backup", "Window", "Hours worked"],
          rows=[[str(i), "", "", "", "", ""] for i in range(1, 32)], blank_rows=0,
          subtitle="Month: ____________   Team: ____________   Escalation contact: ____________", landscape=False)

# 5. Monthly rotating shift calendar (pattern-driven) -------------------------------------
wb, ws = new_sheet("Monthly rotating shift calendar", "Crews follow a repeating pattern. D = day, E = evening, N = night, - = off.",
                   sheet_name="Rotation")
label(ws, 3, 1, "First day of the month", MONTH, vcol=2, fmt="m/d/yyyy")
label(ws, 4, 1, "Date the cycle started (day 1)", datetime.date(2026, 11, 2), vcol=2, fmt="m/d/yyyy")
label(ws, 5, 1, "Hours per shift", 12, vcol=2)
hr = 7
hdr(ws, hr, 1, "Crew")
hdr(ws, hr, 2, "Pattern")
hdr(ws, hr, 3, "Offset")
for c in (1, 2, 3):
    hdr(ws, hr + 1, c, "")
month_headers(ws, hr, 4)
hdr(ws, hr, 35, "Shifts")
hdr(ws, hr, 36, "Hours")
hdr(ws, hr + 1, 35, "")
hdr(ws, hr + 1, 36, "")
r0 = hr + 2
crews = [("Crew A", "DDDD----", 0), ("Crew B", "DDDD----", 4), ("Crew C", "NNNN----", 0), ("Crew D", "NNNN----", 4), ("", "", 0), ("", "", 0)]
for i, (cn, p, o) in enumerate(crews):
    r = r0 + i
    put(ws, r, 1, cn, bold=True)
    put(ws, r, 2, p).font = Font(name="Consolas", size=9)
    put(ws, r, 3, o, center=True)
    for k in range(31):
        c = L(4 + k)
        put(ws, r, 4 + k, f'=IF(OR($B{r}="",{c}${hr + 1}=""),"",MID($B{r},MOD({c}${hr + 1}-$B$4+$C{r},LEN($B{r}))+1,1))', center=True)
    rng = f"D{r}:AH{r}"
    put(ws, r, 35, f'=COUNTIF({rng},"D")+COUNTIF({rng},"E")+COUNTIF({rng},"N")', fill=CALC, center=True)
    put(ws, r, 36, f"=AI{r}*$B$5", fill=CALC, center=True)
rl = r0 + len(crews) - 1
shade(ws, f"D{r0}:AH{rl}", ["D", "E", "N"])
for j, (lab, code) in enumerate([("Day crews working", "D"), ("Evening crews working", "E"), ("Night crews working", "N")]):
    rr = rl + 1 + j
    grey_row(ws, rr, 1, 3, lab)
    for k in range(31):
        c = L(4 + k)
        put(ws, rr, 4 + k, f'=IF({c}${hr + 1}="","",COUNTIF({c}{r0}:{c}{rl},"{code}"))', fill=CALC, center=True)
note(ws, rl + 5, "Example: 4 on, 4 off with 12-hour shifts. Keep the cycle start date the same every month and the rotation "
                 "carries on correctly; just change the first day of the month. Replace the patterns to use your own rotation.")
finish(ws, [28, 16, 7] + [4.6] * 31 + [7, 7], hr + 1, rl + 7, freeze="D9")
wb.save(d / "monthly-rotating-shift-calendar.xlsx")
roster_pdf(d / "monthly-rotating-shift-calendar.pdf", "Monthly rotating shift calendar",
           "Month: ____________   Pattern: ____________   D = day, E = evening, N = night, blank = off", 8, ["Hours"],
           "Write the crew's code under each date. Each crew follows the same pattern, offset so every shift is covered.",
           first_col="Crew")

# 6. Monthly staffing plan (needed vs scheduled) ------------------------------------------
wb, ws = new_sheet("Monthly staffing plan", "Location: ____________   Plan how many people each shift needs, then check the schedule against it.",
                   sheet_name="Staffing plan")
label(ws, 3, 1, "First day of the month", MONTH, vcol=3, fmt="mmmm yyyy")
ws.merge_cells("A3:B3")
hr = 5
heads = ["Date", "Day", "Events or expected volume", "AM needed", "AM scheduled", "AM gap",
         "PM needed", "PM scheduled", "PM gap", "Night needed", "Night scheduled", "Night gap"]
for c, t in enumerate(heads, 1):
    hdr(ws, hr, c, t)
r0 = hr + 1
for k in range(31):
    r = r0 + k
    put(ws, r, 1, f'=IF({in_month(k, "$C$3")},$C$3+{k},"")', fmt="mmm d", fill=GREY, bold=True)
    put(ws, r, 2, f'=IF(A{r}="","",TEXT(A{r},"ddd"))', fill=GREY)
    put(ws, r, 3, "", zebra=k % 2)
    for base in (4, 7, 10):
        put(ws, r, base, None, center=True, fill=INPUT)
        put(ws, r, base + 1, None, center=True, zebra=k % 2)
        put(ws, r, base + 2, f'=IF(COUNT({L(base)}{r}:{L(base + 1)}{r})=0,"",N({L(base + 1)}{r})-N({L(base)}{r}))',
            fill=CALC, center=True, bold=True)
rl = r0 + 30
for base in (4, 7, 10):
    ws.conditional_formatting.add(f"{L(base + 2)}{r0}:{L(base + 2)}{rl}", CellIsRule(operator="lessThan", formula=["0"], fill=RED))
grey_row(ws, rl + 1, 1, 3, "Short shifts this month")
for base in (4, 7, 10):
    put(ws, rl + 1, base, "", fill=GREY)
    put(ws, rl + 1, base + 1, "", fill=GREY)
    put(ws, rl + 1, base + 2, f'=COUNTIF({L(base + 2)}{r0}:{L(base + 2)}{rl},"<0")', fill=CALC, center=True, bold=True)
note(ws, rl + 3, "Gap = scheduled minus needed. Negative gaps (red) are shifts you still need to fill; positive gaps may be "
                 "hours you can cut. Fill in 'needed' from last year's sales, census or appointment volume.")
finish(ws, [10, 7, 30] + [9] * 9, hr, rl + 5, landscape=True)
wb.save(d / "monthly-staffing-plan.xlsx")
pdf_table(d / "monthly-staffing-plan.pdf", "Monthly staffing plan",
          ["Date", "Events / volume", "AM need", "AM sched", "PM need", "PM sched", "Night need", "Night sched"],
          rows=[[str(i)] + [""] * 7 for i in range(1, 32)], blank_rows=0,
          subtitle="Month: ____________   Location: ____________", landscape=False)

# 7. Monthly time off and availability calendar ----------------------------------------
wb, ws = new_sheet("Monthly time off calendar", "Codes: PTO = approved, REQ = requested, UNAV = unavailable, HOL = paid holiday",
                   sheet_name="Time off")
label(ws, 3, 1, "First day of the month", MONTH, vcol=2, fmt="m/d/yyyy")
label(ws, 4, 1, "Most people off on one day", 2, vcol=2)
hr = 6
hdr(ws, hr, 1, "Employee")
hdr(ws, hr, 2, "Team")
hdr(ws, hr + 1, 1, "")
hdr(ws, hr + 1, 2, "")
month_headers(ws, hr, 3)
for j, t in enumerate(["PTO days", "Pending"]):
    hdr(ws, hr, 34 + j, t)
    hdr(ws, hr + 1, 34 + j, "")
r0, n = hr + 2, 25
for i in range(n):
    r = r0 + i
    put(ws, r, 1, "", zebra=i % 2)
    put(ws, r, 2, "", zebra=i % 2)
    for k in range(31):
        put(ws, r, 3 + k, "", center=True)
    put(ws, r, 34, f'=COUNTIF(C{r}:AG{r},"PTO")', fill=CALC, center=True)
    put(ws, r, 35, f'=COUNTIF(C{r}:AG{r},"REQ")', fill=CALC, center=True)
rl = r0 + n - 1
choices(ws, f"C{r0}:AG{rl}", ["PTO", "REQ", "UNAV", "HOL"])
shade(ws, f"C{r0}:AG{rl}", ["PTO", "REQ"])
rr = rl + 1
grey_row(ws, rr, 1, 2, "People off or pending")
for k in range(31):
    c = L(3 + k)
    put(ws, rr, 3 + k, f'=IF({c}${hr + 1}="","",COUNTIF({c}{r0}:{c}{rl},"PTO")+COUNTIF({c}{r0}:{c}{rl},"REQ")+COUNTIF({c}{r0}:{c}{rl},"UNAV"))',
        fill=CALC, center=True, bold=True)
grey_row(ws, rr + 1, 1, 2, "Over the limit?")
for k in range(31):
    c = L(3 + k)
    put(ws, rr + 1, 3 + k, f'=IF({c}{rr}="","",IF({c}{rr}>$B$4,"OVER",""))', fill=CALC, center=True)
ws.conditional_formatting.add(f"C{rr + 1}:AG{rr + 1}", CellIsRule(operator="equal", formula=['"OVER"'], fill=RED))
note(ws, rr + 3, "Mark approved time off, pending requests and days people cannot work before you build the month's schedule. "
                 "'Over the limit' flags days when more people are off than the yellow limit. Paid holidays (HOL) are not counted.")
finish(ws, [20, 12] + [5] * 31 + [7, 7], hr + 1, rr + 5, freeze="C8")
wb.save(d / "monthly-time-off-calendar.xlsx")
roster_pdf(d / "monthly-time-off-calendar.pdf", "Monthly time off calendar",
           "Month: ____________   Team: ____________   PTO = approved, REQ = requested, UNAV = unavailable, HOL = holiday", 18,
           ["PTO days"], "Count people off each day before you approve another request for it.")

report(d)
