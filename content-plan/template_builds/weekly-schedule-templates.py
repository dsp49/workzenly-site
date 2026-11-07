#!/usr/bin/env python3
"""Build the downloads for /templates/weekly-schedule-templates/.

Run from the site root:  python content-plan/template_builds/weekly-schedule-templates.py
Files go to assets/templates/weekly-schedule-templates/. Files are blank and ready to fill in, with working
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


# ============================================================== /templates/weekly-schedule-templates/
d = out("weekly-schedule-templates")
DAYS = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
WEEK = datetime.date(2026, 11, 16)  # a Monday; change it in the sheet
TIME = "h:mm AM/PM"


def day_headers(ws, r, c0, step=1, span=1, start_cell="$B$3"):
    """Day name + date headers driven by the week-start cell."""
    for k in range(7):
        c = c0 + k * step
        hdr(ws, r, c, f'=TEXT({start_cell}+{k},"ddd m/d")', merge_to=c + span - 1 if span > 1 else None)


# 1. Weekly employee schedule with automatic hours ----------------------------------------
wb, ws = new_sheet("Weekly employee schedule", "Business: ____________   Location: ____________   Manager: ____________",
                   sheet_name="Weekly schedule")
label(ws, 3, 1, "Week starting", WEEK, vcol=2, fmt="m/d/yyyy")
hr = 5
hdr(ws, hr, 1, "Employee")
hdr(ws, hr, 2, "Role")
hdr(ws, hr + 1, 1, "")
hdr(ws, hr + 1, 2, "")
day_headers(ws, hr, 3, step=2, span=2)
for k in range(7):
    hdr(ws, hr + 1, 3 + 2 * k, "In")
    hdr(ws, hr + 1, 4 + 2 * k, "Out")
for c, t in [(17, "Unpaid breaks (hours this week)"), (18, "Total hours"), (19, "Overtime hours (over 40)")]:
    hdr(ws, hr, c, t)
    hdr(ws, hr + 1, c, "")
r0, n = hr + 2, 20
for i in range(n):
    r = r0 + i
    put(ws, r, 1, "", zebra=i % 2)
    put(ws, r, 2, "", zebra=i % 2)
    for k in range(14):
        put(ws, r, 3 + k, None, fmt=TIME, center=True, zebra=i % 2)
    put(ws, r, 17, None, fmt="0.00", center=True, zebra=i % 2)
    terms = "+".join(hours_term(f"{L(3 + 2 * k)}{r}", f"{L(4 + 2 * k)}{r}") for k in range(7))
    put(ws, r, 18, f"=ROUND(24*({terms})-N(Q{r}),2)", fmt="0.00", fill=CALC, center=True)
    put(ws, r, 19, f"=MAX(0,R{r}-40)", fmt="0.00", fill=CALC, center=True)
rl = r0 + n - 1
put(ws, rl + 1, 1, "Totals", bold=True, fill=GREY)
for c in range(2, 18):
    put(ws, rl + 1, c, "", fill=GREY)
put(ws, rl + 1, 18, f"=SUM(R{r0}:R{rl})", fmt="0.00", fill=CALC, bold=True, center=True)
put(ws, rl + 1, 19, f"=SUM(S{r0}:S{rl})", fmt="0.00", fill=CALC, bold=True, center=True)
note(ws, rl + 3, "Type times like 9:00 AM or 21:30. A shift that ends after midnight (Out earlier than In) counts correctly. "
                 "Enter unpaid meal breaks as hours for the week (three 30-minute lunches = 1.5). Short rest breaks are paid "
                 "under federal law, so do not subtract them.")
finish(ws, [22, 14] + [9.5] * 14 + [12, 10, 11], hr + 1, rl + 5, freeze="C7")
wb.save(d / "weekly-employee-schedule-with-hours.xlsx")
pdf_table(d / "weekly-employee-schedule-with-hours.pdf", "Weekly employee schedule",
          ["Employee", "Role"] + [f"{x} in / out" for x in DAYS] + ["Breaks", "Total hours"], blank_rows=15,
          subtitle="Week starting: ____________   Location: ____________   Manager: ____________",
          notes="Write each shift as in / out times, for example 9:00a / 5:30p. Total hours = time worked minus unpaid meal breaks.")

# 2. Printable weekly work schedule (simple) ----------------------------------------------
h2 = ["Employee", "Position"] + DAYS + ["Weekly hours"]
xlsx_sheet(d / "printable-weekly-work-schedule.xlsx", title="Weekly work schedule",
           subtitle="Week of: ____________   Location: ____________   Posted on: ____________   Manager: ____________",
           headers=h2, rows=[], blank_rows=18, widths=[22, 14] + [13] * 7 + [10], sheet_name="Work schedule",
           notes="Write shifts as 9a-5p, OFF for a day off and PTO for approved time off. Post it where everyone can see it.")
pdf_table(d / "printable-weekly-work-schedule.pdf", "Weekly work schedule", h2, blank_rows=16,
          subtitle="Week of: ____________   Location: ____________   Posted on: ____________   Manager: ____________",
          notes="Write shifts as 9a-5p, OFF for a day off and PTO for approved time off.")

# 3. Weekly schedule by hour (time blocks) ------------------------------------------------
wb, ws = new_sheet("Weekly schedule by hour", "Name: ____________   Week starting:", sheet_name="By hour")
ws["C2"] = WEEK
ws["C2"].number_format = "mmm d, yyyy"
ws["C2"].fill = INPUT
hr = 4
hdr(ws, hr, 1, "Time")
for k in range(7):
    hdr(ws, hr, 2 + k, f'=TEXT($C$2+{k},"ddd m/d")')
slots = [datetime.time(h, m) for h in range(6, 22) for m in (0, 30)] + [datetime.time(22, 0)]
for i, t in enumerate(slots):
    r = hr + 1 + i
    put(ws, r, 1, t, fmt="h:mm AM/PM", bold=True, fill=GREY)
    for k in range(7):
        put(ws, r, 2 + k, "", zebra=i % 2)
r = hr + 1 + len(slots)
put(ws, r, 1, "Top 3 this week", bold=True, fill=GREY)
for k in range(7):
    put(ws, r, 2 + k, "")
finish(ws, [11] + [18] * 7, hr, r + 2, landscape=False)
wb.save(d / "weekly-schedule-by-hour.xlsx")
# Custom layout so all 33 half-hour rows and the "Top 3" row fit on one portrait page.
_rows = "".join("<tr><td>" + _esc(t.strftime("%I:%M %p").lstrip("0")) + "</td>" + "<td></td>" * 7 + "</tr>" for t in slots)
_rows += "<tr class='top3'><td><b>Top 3 this week</b></td>" + "<td></td>" * 7 + "</tr>"
_pdf(d / "weekly-schedule-by-hour.pdf",
     "<style>th,td{padding:3px 6px;height:13px;font-size:9pt} tr.top3 td{height:40px}</style>"
     "<h1>Weekly schedule by hour</h1><p class='sub'>Name: ____________   Week starting: ____________</p>"
     "<table><thead><tr>" + "".join("<th>" + _esc(h) + "</th>" for h in ["Time"] + DAYS) + "</tr></thead><tbody>"
     + _rows + "</tbody></table>", landscape=False)

# 4. Weekly schedule by position (coverage view) -----------------------------------------
wb, ws = new_sheet("Weekly schedule by position", "Location: ____________   Manager: ____________", sheet_name="By position")
label(ws, 3, 1, "Week starting", WEEK, vcol=2, fmt="m/d/yyyy")
hr = 5
hdr(ws, hr, 1, "Position")
hdr(ws, hr, 2, "Shift (start to end)")
day_headers(ws, hr, 3)
hdr(ws, hr, 10, "Shifts filled")
hdr(ws, hr, 11, "Open shifts")
r0, n = hr + 1, 24
for i in range(n):
    r = r0 + i
    put(ws, r, 1, "", bold=True, zebra=i % 2)
    put(ws, r, 2, "", zebra=i % 2)
    for k in range(7):
        put(ws, r, 3 + k, "", zebra=i % 2)
    put(ws, r, 10, f'=COUNTA(C{r}:I{r})-COUNTIF(C{r}:I{r},"OPEN")-COUNTIF(C{r}:I{r},"CLOSED")', fill=CALC, center=True)
    put(ws, r, 11, f'=COUNTIF(C{r}:I{r},"OPEN")', fill=CALC, center=True)
rl = r0 + n - 1
put(ws, rl + 1, 1, "People working", bold=True, fill=GREY)
put(ws, rl + 1, 2, "", fill=GREY)
for k in range(7):
    c = L(3 + k)
    put(ws, rl + 1, 3 + k, f'=COUNTA({c}{r0}:{c}{rl})-COUNTIF({c}{r0}:{c}{rl},"OPEN")-COUNTIF({c}{r0}:{c}{rl},"CLOSED")',
        fill=CALC, center=True, bold=True)
put(ws, rl + 2, 1, "Open shifts", bold=True, fill=GREY)
put(ws, rl + 2, 2, "", fill=GREY)
for k in range(7):
    c = L(3 + k)
    put(ws, rl + 2, 3 + k, f'=COUNTIF({c}{r0}:{c}{rl},"OPEN")', fill=CALC, center=True, bold=True)
note(ws, rl + 4, "One row per position and shift. Type the employee's name in each day, OPEN for a shift you still need "
                 "to fill, or CLOSED if the business is closed.")
finish(ws, [20, 16] + [14] * 7 + [9, 9], hr, rl + 6)
wb.save(d / "weekly-schedule-by-position.xlsx")
pdf_table(d / "weekly-schedule-by-position.pdf", "Weekly schedule by position", ["Position", "Shift"] + DAYS,
          blank_rows=18, subtitle="Week of: ____________   Location: ____________",
          notes="One row per position and shift. Write the employee's name in each day, or OPEN for a shift to fill.")

# 5. Weekly schedule with labor cost ------------------------------------------------------
wb, ws = new_sheet("Weekly schedule with labor cost", "Location: ____________   Manager: ____________", sheet_name="Labor cost")
label(ws, 3, 1, "Week starting", WEEK, vcol=2, fmt="m/d/yyyy")
label(ws, 3, 4, "Projected sales", None, vcol=6, fmt='"$"#,##0')
label(ws, 3, 8, "Labor budget %", 0.28, vcol=10, fmt="0%")
hr = 5
for c, t in enumerate(["Employee", "Role", "Hourly rate"], 1):
    hdr(ws, hr, c, t)
day_headers(ws, hr, 4)
for c, t in [(11, "Total hours"), (12, "Regular hours"), (13, "Overtime hours"), (14, "Estimated wages")]:
    hdr(ws, hr, c, t)
r0, n = hr + 1, 20
for i in range(n):
    r = r0 + i
    put(ws, r, 1, "", zebra=i % 2)
    put(ws, r, 2, "", zebra=i % 2)
    put(ws, r, 3, None, fmt='"$"#,##0.00', zebra=i % 2)
    for k in range(7):
        put(ws, r, 4 + k, None, fmt="General", center=True, zebra=i % 2)
    put(ws, r, 11, f"=SUM(D{r}:J{r})", fmt="General", fill=CALC, center=True)
    put(ws, r, 12, f"=MIN(40,K{r})", fmt="General", fill=CALC, center=True)
    put(ws, r, 13, f"=MAX(0,K{r}-40)", fmt="General", fill=CALC, center=True)
    put(ws, r, 14, f"=N(C{r})*L{r}+N(C{r})*1.5*M{r}", fmt='"$"#,##0.00', fill=CALC)
rl = r0 + n - 1
tr = rl + 1
put(ws, tr, 1, "Hours per day", bold=True, fill=GREY)
put(ws, tr, 2, "", fill=GREY)
put(ws, tr, 3, "", fill=GREY)
for k in range(7):
    c = L(4 + k)
    put(ws, tr, 4 + k, f"=SUM({c}{r0}:{c}{rl})", fmt="General", fill=CALC, center=True, bold=True)
put(ws, tr + 1, 1, "Straight-time wages per day", bold=True, fill=GREY)
put(ws, tr + 1, 2, "", fill=GREY)
put(ws, tr + 1, 3, "", fill=GREY)
for k in range(7):
    c = L(4 + k)
    put(ws, tr + 1, 4 + k, f"=SUMPRODUCT($C${r0}:$C${rl},{c}{r0}:{c}{rl})", fmt='"$"#,##0', fill=CALC, center=True)
for c, f, fm in [(11, f"=SUM(K{r0}:K{rl})", "General"), (12, f"=SUM(L{r0}:L{rl})", "General"),
                 (13, f"=SUM(M{r0}:M{rl})", "General"), (14, f"=SUM(N{r0}:N{rl})", '"$"#,##0.00')]:
    put(ws, tr, c, f, fmt=fm, fill=CALC, bold=True, center=True)
sr = tr + 3
for i, (lab, f, fm) in enumerate([("Estimated wages this week", f"=N{tr}", '"$"#,##0.00'),
                                  ("Labor budget (sales x budget %)", "=N(F3)*N(J3)", '"$"#,##0.00'),
                                  ("Over (+) or under (-) budget", f"=N{tr}-N(F3)*N(J3)", '"$"#,##0.00'),
                                  ("Wages as % of projected sales", '=IF(N(F3)=0,"",N' + str(tr) + '/F3)', "0.0%")]):
    put(ws, sr + i, 1, lab, bold=True, fill=GREY)
    ws.merge_cells(start_row=sr + i, start_column=1, end_row=sr + i, end_column=3)
    put(ws, sr + i, 4, f, fmt=fm, fill=CALC, bold=True)
    ws.merge_cells(start_row=sr + i, start_column=4, end_row=sr + i, end_column=5)
note(ws, sr + 5, "Enter scheduled hours per day as numbers (7.5 for a 9:00 to 5:00 shift with a 30-minute unpaid lunch). "
                 "Overtime is estimated at 1.5 x the hourly rate for hours over 40 in the week; bonuses, tips and shift "
                 "differentials can raise the regular rate. Payroll taxes and benefits are not included.")
finish(ws, [22, 14, 11] + [10] * 7 + [10, 10, 10, 13], hr, sr + 7)
wb.save(d / "weekly-schedule-with-labor-cost.xlsx")
pdf_table(d / "weekly-schedule-with-labor-cost.pdf", "Weekly schedule with labor cost",
          ["Employee", "Rate"] + DAYS + ["Hours", "OT hours", "Est. wages"], blank_rows=15,
          subtitle="Week of: ________   Projected sales: $________   Labor budget: ____ %",
          notes="Hours per day as numbers. Overtime = hours over 40 in the workweek, paid at 1.5 x the regular rate. "
                "Totals: hours ______   wages $________   wages as % of sales ______")

# 6. Weekly team schedule with time off and availability ---------------------------------
wb, ws = new_sheet("Weekly team schedule with time off", "Location: ____________   Manager: ____________",
                   sheet_name="Team schedule")
label(ws, 3, 1, "Week starting", WEEK, vcol=2, fmt="m/d/yyyy")
hr = 5
hdr(ws, hr, 1, "Employee")
hdr(ws, hr, 2, "Role")
day_headers(ws, hr, 3)
hdr(ws, hr, 10, "Shifts this week")
hdr(ws, hr, 11, "Days off or PTO")
r0, n = hr + 1, 20
codes_off = ["OFF", "PTO", "REQ", "UNAV"]
for i in range(n):
    r = r0 + i
    put(ws, r, 1, "", zebra=i % 2)
    put(ws, r, 2, "", zebra=i % 2)
    for k in range(7):
        put(ws, r, 3 + k, "", center=True, zebra=i % 2)
    rng = f"C{r}:I{r}"
    offs = "-".join(f'COUNTIF({rng},"{x}")' for x in codes_off)
    put(ws, r, 10, f"=COUNTA({rng})-{offs}", fill=CALC, center=True)
    put(ws, r, 11, f'=COUNTIF({rng},"OFF")+COUNTIF({rng},"PTO")', fill=CALC, center=True)
rl = r0 + n - 1
for j, (lab, f) in enumerate([("Working", None), ("On PTO", "PTO"), ("Requested off (pending)", "REQ"),
                              ("Unavailable", "UNAV")]):
    r = rl + 1 + j
    put(ws, r, 1, lab, bold=True, fill=GREY)
    put(ws, r, 2, "", fill=GREY)
    for k in range(7):
        c = L(3 + k)
        rng = f"{c}{r0}:{c}{rl}"
        if f:
            v = f'=COUNTIF({rng},"{f}")'
        else:
            v = f"=COUNTA({rng})-" + "-".join(f'COUNTIF({rng},"{x}")' for x in codes_off)
        put(ws, r, 3 + k, v, fill=CALC, center=True, bold=j == 0)
for code, color in [("PTO", "C9D7F2"), ("REQ", "FFE3A3"), ("UNAV", "E2E2DC"), ("OFF", "F2F2EE")]:
    from openpyxl.formatting.rule import CellIsRule
    ws.conditional_formatting.add(f"C{r0}:I{rl}", CellIsRule(operator="equal", formula=[f'"{code}"'],
                                                            fill=PatternFill("solid", fgColor=color)))
note(ws, rl + 7, "Type a shift (9a-5p) or a code: OFF = day off, PTO = approved time off, REQ = time off requested, "
                 "not yet approved, UNAV = employee not available that day. Resolve every REQ before you post the schedule.")
finish(ws, [22, 14] + [13] * 7 + [10, 10], hr, rl + 9)
wb.save(d / "weekly-team-schedule-with-time-off.xlsx")
pdf_table(d / "weekly-team-schedule-with-time-off.pdf", "Weekly team schedule with time off", ["Employee", "Role"] + DAYS + ["Shifts"],
          blank_rows=16, subtitle="Week of: ____________   Location: ____________",
          notes="Write a shift (9a-5p) or a code: OFF, PTO (approved time off), REQ (requested, pending), UNAV (not available).")

# 7. My weekly work schedule (one employee) -----------------------------------------------
wb, ws = new_sheet("My weekly work schedule", "Name: ____________   Employer: ____________", sheet_name="My week")
label(ws, 3, 1, "Week starting", WEEK, vcol=2, fmt="m/d/yyyy")
label(ws, 3, 4, "Hourly rate", None, vcol=5, fmt='"$"#,##0.00')
hr = 5
for c, t in enumerate(["Day", "Date", "Start", "End", "Unpaid break (minutes)", "Hours worked", "Notes"], 1):
    hdr(ws, hr, c, t)
for k in range(7):
    r = hr + 1 + k
    put(ws, r, 1, f'=TEXT($B$3+{k},"dddd")', bold=True, fill=GREY)
    put(ws, r, 2, f"=$B$3+{k}", fmt="mmm d", fill=GREY)
    put(ws, r, 3, None, fmt=TIME, center=True)
    put(ws, r, 4, None, fmt=TIME, center=True)
    put(ws, r, 5, None, center=True)
    put(ws, r, 6, f"=IF(COUNT(C{r}:D{r})=2,ROUND(MOD(D{r}-C{r},1)*24-N(E{r})/60,2),0)", fmt="0.00", fill=CALC, center=True)
    put(ws, r, 7, "")
s = hr + 9
for i, (lab, f, fm) in enumerate([("Total hours", f"=SUM(F{hr + 1}:F{hr + 7})", "0.00"),
                                  ("Regular hours (up to 40)", f"=MIN(40,F{s})", "0.00"),
                                  ("Overtime hours (over 40)", f"=MAX(0,F{s}-40)", "0.00"),
                                  ("Estimated gross pay", f"=N(E3)*F{s + 1}+N(E3)*1.5*F{s + 2}", '"$"#,##0.00')]):
    put(ws, s + i, 4, lab, bold=True, fill=GREY)
    ws.merge_cells(start_row=s + i, start_column=4, end_row=s + i, end_column=5)
    put(ws, s + i, 6, f, fmt=fm, fill=CALC, bold=True, center=True)
note(ws, s + 5, "For non-exempt employees, federal law requires overtime at 1.5 x the regular rate for hours over 40 in a "
                "workweek. Some states, such as California, also count daily overtime. Gross pay is before taxes and deductions.")
finish(ws, [14, 12, 12, 12, 14, 12, 30], hr, s + 7, landscape=False)
wb.save(d / "my-weekly-work-schedule.xlsx")
pdf_doc(d / "my-weekly-work-schedule.pdf", "My weekly work schedule", [
    ("fields", ["Name", "Employer", "Week starting", "Hourly rate"]),
    ("table", ["Day", "Date", "Start", "End", "Unpaid break", "Hours", "Notes"],
     [[x, "", "", "", "", "", ""] for x in ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]]),
    ("fields", ["Total hours", "Regular hours (up to 40)", "Overtime hours (over 40)", "Estimated gross pay"]),
])

report(d)
