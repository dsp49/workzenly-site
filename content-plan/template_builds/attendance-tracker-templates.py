#!/usr/bin/env python3
"""Build the downloads for /templates/attendance-tracker-templates/.

Run from the site root:  python content-plan/template_builds/attendance-tracker-templates.py
Files go to assets/templates/attendance-tracker-templates/. Files are blank and ready to fill in, with working
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
from openpyxl.formatting.rule import FormulaRule
from tplkit import out, xlsx_sheet, docx_doc, pdf_table, pdf_doc, _pdf, _esc, FOOT

THIN = Side(style="thin", color="C9C9C3")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
DARK = PatternFill("solid", fgColor="161616")
GREY = PatternFill("solid", fgColor="E9E9E3")
ZEBRA = PatternFill("solid", fgColor="F6F6F3")
CALC = PatternFill("solid", fgColor="F2FFD9")   # cells with formulas
INPUT = PatternFill("solid", fgColor="FFF8D6")  # cells the user types into at the top
MONEY = '"$"#,##0.00'
TIME = "h:mm AM/PM"
DATE = "m/d/yyyy"


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
    """A bold label with a yellow input cell next to it."""
    a = ws.cell(row=r, column=c, value=text)
    a.font = Font(bold=True)
    a.alignment = Alignment(horizontal="right", vertical="center")
    v = ws.cell(row=r, column=vcol or c + 1, value=value)
    v.fill, v.border = INPUT, BOX
    if fmt:
        v.number_format = fmt
    return v


def calc_label(ws, r, c, text, formula, fmt=None, vcol=None):
    a = ws.cell(row=r, column=c, value=text)
    a.font = Font(bold=True)
    a.alignment = Alignment(horizontal="right", vertical="center")
    return put(ws, r, vcol or c + 1, formula, fmt=fmt, fill=CALC, bold=True, center=True)


def note(ws, r, text, c=1):
    ws.cell(row=r, column=c, value=text).font = Font(italic=True, size=10, color="555555")


def sign_lines(ws, r, labels, c=1, gap=4):
    """Signature lines in the sheet: label row under a ruled line."""
    for i, lab in enumerate(labels):
        rr = r + i * 2
        for cc in range(c, c + gap):
            ws.cell(row=rr, column=cc).border = Border(bottom=Side(style="thin", color="333333"))
        ws.cell(row=rr + 1, column=c, value=lab).font = Font(size=9, color="444444")
        ws.cell(row=rr, column=c + gap + 1).border = Border(bottom=Side(style="thin", color="333333"))
        ws.cell(row=rr + 1, column=c + gap + 1, value="Date").font = Font(size=9, color="444444")
    return r + len(labels) * 2


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


def html_pdf(path, title, body_html, subtitle=None, landscape=True, extra_css=""):
    body = f"<style>{extra_css}</style><h1>{_esc(title)}</h1>"
    if subtitle:
        body += f"<p class='sub'>{_esc(subtitle)}</p>"
    return _pdf(path, body + body_html, landscape)


def grid_html(headers, rows, cls="", style=""):
    h = f"<table class='{cls}' style='{style}'><thead><tr>" + "".join(f"<th>{_esc(x)}</th>" for x in headers) + "</tr></thead><tbody>"
    h += "".join("<tr>" + "".join(f"<td>{_esc(v)}</td>" for v in r) + "</tr>" for r in rows)
    return h + "</tbody></table>"


def sign_html(labels):
    return "<div class='sign'>" + "".join(f"<div>{_esc(l)}</div><div>Date</div>" for l in labels) + "</div>"


def report(d):
    for f in sorted(d.iterdir()):
        print(f"{f.name:55s} {f.stat().st_size:>8,d} bytes")


# ============================================================== /templates/attendance-tracker-templates/
d = out("attendance-tracker-templates")
MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October",
          "November", "December"]
CODES = [("V", "Vacation or PTO"), ("S", "Sick"), ("P", "Personal day"), ("H", "Paid holiday"),
         ("U", "Unpaid or unexcused absence"), ("L", "Late arrival"), ("E", "Left early"),
         ("F", "FMLA or other protected leave")]
CODE_FILLS = {"V": "D9F2C2", "P": "D9F2C2", "S": "F8D0CC", "U": "F6B98E", "L": "FFF0A8", "E": "FFF0A8",
              "F": "CFE0F7", "H": "DADDE6", "X": "FFFFFF"}
PDF_CSS = (".info{display:grid;grid-template-columns:repeat(3,1fr);gap:8px 22px;margin:0 0 14px;font-size:10pt}"
           ".info div{border-bottom:1px solid #333;padding:12px 0 2px;color:#333}"
           ".note{font-size:9pt;color:#444;margin-top:12px} .foot{position:static;margin-top:22px}"
           ".grid{table-layout:fixed} .grid th,.grid td{padding:3px 1px;text-align:center;font-size:7.5pt} .grid td{height:22px}"
           ".grid td:first-child,.grid th:first-child{text-align:left;padding-left:5px;width:84px}"
           ".legend{font-size:8.5pt;color:#333;margin-top:10px} .sum{width:auto;min-width:46%;margin-top:14px}"
           ".sum th{background:#e9e9e3;color:#111}")


def info_html(labels):
    return "<div class='info'>" + "".join(f"<div>{_esc(l)}:</div>" for l in labels) + "</div>"


def legend_text(extra=()):
    return "Codes: " + "   ".join(f"{c} = {t}" for c, t in list(extra) + CODES)


def code_rules(ws, rng, codes):
    first = rng.split(":")[0]
    for c in codes:
        ws.conditional_formatting.add(rng, FormulaRule(formula=[f'UPPER({first})="{c}"'],
                                                       fill=PatternFill("solid", fgColor=CODE_FILLS[c])))


# 1. Yearly employee attendance tracker ----------------------------------------------------
wb, ws = new_sheet("Employee attendance tracker", "Company: ____________   Department: ____________   Manager: ____________",
                   sheet_name="Yearly attendance")
label(ws, 3, 1, "Employee", "", vcol=2)
ws.merge_cells("B3:H3")
label(ws, 4, 1, "Year", 2027, vcol=2)
ws.merge_cells("B4:C4")
hr = 6
hdr(ws, hr, 1, "Month")
for k in range(31):
    hdr(ws, hr, 2 + k, k + 1)
cc0 = 33
for i, (c, t) in enumerate(CODES):
    hdr(ws, hr, cc0 + i, c)
for m in range(12):
    r = hr + 1 + m
    put(ws, r, 1, MONTHS[m], bold=True, fill=GREY)
    for k in range(31):
        put(ws, r, 2 + k, "", center=True)
    for i, (c, t) in enumerate(CODES):
        put(ws, r, cc0 + i, f'=COUNTIF(B{r}:AF{r},"{c}")', fill=CALC, center=True)
rl = hr + 12
grid = f"B{hr + 1}:AF{rl}"
ws.conditional_formatting.add(grid, FormulaRule(formula=[f"DAY(DATE($B$4,ROW()-{hr},B${hr}))<>B${hr}"],
                                                 fill=PatternFill("solid", fgColor="8C8C86"), stopIfTrue=True))
code_rules(ws, grid, [c for c, _ in CODES])
ws.conditional_formatting.add(grid, FormulaRule(formula=[f"WEEKDAY(DATE($B$4,ROW()-{hr},B${hr}),2)>5"],
                                                 fill=PatternFill("solid", fgColor="E4E4DD")))
choices(ws, grid, [c for c, _ in CODES])
put(ws, rl + 1, 1, "Year total", bold=True, fill=GREY)
for k in range(31):
    put(ws, rl + 1, 2 + k, "", fill=GREY)
for i in range(len(CODES)):
    col = L(cc0 + i)
    put(ws, rl + 1, cc0 + i, f"=SUM({col}{hr + 1}:{col}{rl})", fill=CALC, bold=True, center=True)
s = rl + 3
calc_label(ws, s, 1, "Days absent (V+S+P+U+F)", f"=SUM({L(cc0)}{rl + 1}:{L(cc0 + 2)}{rl + 1})+{L(cc0 + 4)}{rl + 1}+{L(cc0 + 7)}{rl + 1}", vcol=8)
ws.merge_cells(start_row=s, start_column=1, end_row=s, end_column=7)
calc_label(ws, s + 1, 1, "Unplanned absences (S+U)", f"={L(cc0 + 1)}{rl + 1}+{L(cc0 + 4)}{rl + 1}", vcol=8)
ws.merge_cells(start_row=s + 1, start_column=1, end_row=s + 1, end_column=7)
calc_label(ws, s + 2, 1, "Late arrivals and early departures (L+E)", f"={L(cc0 + 5)}{rl + 1}+{L(cc0 + 6)}{rl + 1}", vcol=8)
ws.merge_cells(start_row=s + 2, start_column=1, end_row=s + 2, end_column=7)
note(ws, s + 4, legend_text())
note(ws, s + 5, "Type one code per day; leave normal workdays blank. Weekends are shaded and days that do not exist "
                "(like February 30) are dark. Change the year and the calendar updates.")
note(ws, s + 6, "Keep FMLA and other legally protected leave coded F so it is never counted against the employee.")
finish(ws, [12] + [3.6] * 31 + [4.2] * 8, hr, s + 8, freeze="B7")
wb.save(d / "yearly-employee-attendance-tracker.xlsx")
rows = [[m] + [""] * 31 + [""] * 8 for m in MONTHS] + [["Total"] + [""] * 39]
html_pdf(d / "yearly-employee-attendance-tracker.pdf", "Employee attendance tracker",
         info_html(["Employee", "Year", "Department"]) +
         grid_html(["Month"] + [str(i) for i in range(1, 32)] + [c for c, _ in CODES], rows, cls="grid") +
         f"<p class='legend'>{_esc(legend_text())}</p>"
         "<p class='note'>Write one code per day; leave normal workdays blank. Keep protected leave coded F.</p>",
         extra_css=PDF_CSS)

# 2. Monthly team attendance sheet ------------------------------------------------------------
wb, ws = new_sheet("Monthly attendance sheet", "Company: ____________   Manager: ____________",
                   sheet_name="Monthly attendance")
label(ws, 3, 1, "Location", "", vcol=2)
label(ws, 4, 1, "Month (first day)", datetime.date(2026, 11, 1), vcol=2, fmt="mmmm yyyy")
hr = 6
hdr(ws, hr, 1, "Employee")
hdr(ws, hr + 1, 1, "")
for k in range(31):
    c = 2 + k
    hdr(ws, hr, c, f'=IF($B$4+{k}>EOMONTH($B$4,0),"",DAY($B$4+{k}))')
    hdr(ws, hr + 1, c, f'=IF($B$4+{k}>EOMONTH($B$4,0),"",LEFT(TEXT($B$4+{k},"ddd"),2))')
TEAM_CODES = ["X", "V", "S", "P", "H", "U", "L", "E", "F"]
for i, c in enumerate(TEAM_CODES):
    hdr(ws, hr, cc0 + i, "Worked" if c == "X" else c)
    hdr(ws, hr + 1, cc0 + i, "X" if c == "X" else "")
hdr(ws, hr, cc0 + 9, "Attendance rate")
hdr(ws, hr + 1, cc0 + 9, "")
r0, n = hr + 2, 20
for i in range(n):
    r = r0 + i
    put(ws, r, 1, "", bold=True, zebra=i % 2)
    for k in range(31):
        put(ws, r, 2 + k, "", center=True)
    for j, c in enumerate(TEAM_CODES):
        fx = (f'=COUNTIF(B{r}:AF{r},"X")+COUNTIF(B{r}:AF{r},"L")+COUNTIF(B{r}:AF{r},"E")' if c == "X"
              else f'=COUNTIF(B{r}:AF{r},"{c}")')
        put(ws, r, cc0 + j, fx, fill=CALC, center=True)
    xs, ss, us = L(cc0), L(cc0 + 2), L(cc0 + 5)
    put(ws, r, cc0 + 9, f"=IF(({xs}{r}+{ss}{r}+{us}{r})=0,\"\",{xs}{r}/({xs}{r}+{ss}{r}+{us}{r}))", fmt="0%", fill=CALC, center=True)
rl = r0 + n - 1
grid = f"B{r0}:AF{rl}"
ws.conditional_formatting.add(grid, FormulaRule(formula=[f'B${hr}=""'], fill=PatternFill("solid", fgColor="8C8C86"), stopIfTrue=True))
code_rules(ws, grid, [c for c in TEAM_CODES if c != "X"])
ws.conditional_formatting.add(grid, FormulaRule(formula=[f'OR(B${hr + 1}="Sa",B${hr + 1}="Su")'],
                                                 fill=PatternFill("solid", fgColor="E4E4DD")))
choices(ws, grid, TEAM_CODES)
put(ws, rl + 1, 1, "Team total", bold=True, fill=GREY)
for k in range(31):
    col = L(2 + k)
    put(ws, rl + 1, 2 + k, f'=IF({col}${hr}="","",COUNTIF({col}{r0}:{col}{rl},"X")+COUNTIF({col}{r0}:{col}{rl},"L")'
                           f'+COUNTIF({col}{r0}:{col}{rl},"E"))', fill=CALC, center=True)
for j in range(len(TEAM_CODES)):
    col = L(cc0 + j)
    put(ws, rl + 1, cc0 + j, f"=SUM({col}{r0}:{col}{rl})", fill=CALC, bold=True, center=True)
xs, ss, us = L(cc0), L(cc0 + 2), L(cc0 + 5)
put(ws, rl + 1, cc0 + 9, f"=IF(({xs}{rl + 1}+{ss}{rl + 1}+{us}{rl + 1})=0,\"\",{xs}{rl + 1}/({xs}{rl + 1}+{ss}{rl + 1}+{us}{rl + 1}))",
    fmt="0%", fill=CALC, bold=True, center=True)
note(ws, rl + 3, legend_text([("X", "Worked as scheduled")]))
note(ws, rl + 4, "Leave days the person was not scheduled blank. Worked = X + L + E days. Attendance rate = days worked / "
                 "(days worked + sick days + unexcused absences). The team total row counts how many people worked each day.")
finish(ws, [20] + [3.6] * 31 + [6.5] + [3.8] * 8 + [9], hr + 1, rl + 6, freeze="B8")
wb.save(d / "monthly-attendance-sheet.xlsx")
rows = [[""] * 41 for _ in range(15)]
html_pdf(d / "monthly-attendance-sheet.pdf", "Monthly attendance sheet",
         info_html(["Month", "Location", "Manager"]) +
         grid_html(["Employee"] + [str(i) for i in range(1, 32)] + ["X", "V", "S", "P", "H", "U", "L", "E", "F"], rows, cls="grid") +
         f"<p class='legend'>{_esc(legend_text([('X', 'Worked as scheduled')]))}</p>", extra_css=PDF_CSS)

# 3. Daily attendance sign-in sheet ---------------------------------------------------------
wb, ws = new_sheet("Daily attendance sign-in sheet", "Company: ____________   Supervisor on duty: ____________",
                   sheet_name="Sign-in sheet")
label(ws, 3, 1, "Date", datetime.date(2026, 11, 9), vcol=2, fmt="dddd, mmmm d, yyyy")
ws.merge_cells("B3:C3")
label(ws, 4, 1, "Location or event", "", vcol=2)
ws.merge_cells("B4:C4")
hr = 6
heads = ["#", "Name (print)", "Time in", "Signature", "Time out", "Signature", "Hours", "Notes"]
for c, t in enumerate(heads, 1):
    hdr(ws, hr, c, t)
r0, n = hr + 1, 30
for i in range(n):
    r = r0 + i
    z = i % 2
    put(ws, r, 1, i + 1, center=True, zebra=z)
    put(ws, r, 2, "", zebra=z)
    put(ws, r, 3, None, fmt=TIME, zebra=z, center=True)
    put(ws, r, 4, "", zebra=z)
    put(ws, r, 5, None, fmt=TIME, zebra=z, center=True)
    put(ws, r, 6, "", zebra=z)
    put(ws, r, 7, f"=IF(COUNT(C{r},E{r})<2,\"\",ROUND(24*MOD(E{r}-C{r},1),2))", fmt="0.00", fill=CALC, center=True)
    put(ws, r, 8, "", zebra=z)
    ws.row_dimensions[r].height = 24
rl = r0 + n - 1
calc_label(ws, rl + 2, 5, "People signed in", f"=COUNTA(B{r0}:B{rl})", vcol=7)
note(ws, rl + 4, "Each person signs in and out themselves. Hours here are time on site; subtract unpaid meal breaks on the timesheet.")
finish(ws, [5, 26, 11, 22, 11, 22, 8, 22], hr, rl + 6, landscape=False)
wb.save(d / "daily-attendance-sign-in-sheet.xlsx")
html_pdf(d / "daily-attendance-sign-in-sheet.pdf", "Daily attendance sign-in sheet",
         info_html(["Date", "Location or event", "Supervisor on duty"]) +
         grid_html(["#", "Name (print)", "Time in", "Signature", "Time out", "Signature", "Notes"],
                   [[str(i)] + [""] * 6 for i in range(1, 23)]) +
         "<p class='note'>Each person signs in and out themselves. Hours on this sheet are time on site.</p>",
         landscape=False, extra_css=PDF_CSS + " td{height:19px}")

# 4. Attendance points tracker ----------------------------------------------------------------
wb, ws = new_sheet("Attendance points tracker", "Company: ____________   Manager: ____________",
                   sheet_name="Points tracker")
label(ws, 3, 1, "Employee", "", vcol=2)
label(ws, 4, 1, "Points as of", "=TODAY()", vcol=2, fmt=DATE)
label(ws, 5, 1, "Points expire after (months)", 12, vcol=2)
# policy table on the right
hdr(ws, 3, 9, "Occurrence type")
hdr(ws, 3, 10, "Points")
policy = [("Late arrival (15+ minutes)", 0.5), ("Left early without approval", 0.5), ("Absence with notice", 1),
          ("Absence without proper notice", 2), ("No-call no-show", 3), ("Protected leave (FMLA, sick leave law)", 0),
          ("Approved time off", 0)]
for i, (t, p) in enumerate(policy):
    put(ws, 4 + i, 9, t, fill=INPUT)
    put(ws, 4 + i, 10, p, fill=INPUT, center=True)
hdr(ws, 12, 9, "Step")
hdr(ws, 12, 10, "At points")
steps = [("Verbal warning", 4), ("Written warning", 6), ("Final written warning", 8), ("Termination review", 10)]
for i, (t, p) in enumerate(steps):
    put(ws, 13 + i, 9, t, fill=INPUT)
    put(ws, 13 + i, 10, p, fill=INPUT, center=True)
hr = 8
heads = ["Date", "Occurrence type", "Details", "Points", "Expires on", "Active?", "Recorded by"]
for c, t in enumerate(heads, 1):
    hdr(ws, hr, c, t)
r0, n = hr + 1, 30
for i in range(n):
    r = r0 + i
    z = i % 2
    put(ws, r, 1, None, fmt=DATE, zebra=z, center=True)
    put(ws, r, 2, "", zebra=z)
    put(ws, r, 3, "", zebra=z)
    put(ws, r, 4, f'=IF(B{r}="","",IFERROR(VLOOKUP(B{r},$I$4:$J$10,2,FALSE),0))', fmt="0.0", fill=CALC, center=True)
    put(ws, r, 5, f'=IF(A{r}="","",EDATE(A{r},$B$5))', fmt=DATE, fill=CALC, center=True)
    put(ws, r, 6, f'=IF(A{r}="","",IF(E{r}>$B$4,"Yes","No"))', fill=CALC, center=True)
    put(ws, r, 7, "", zebra=z)
rl = r0 + n - 1
choices(ws, f"B{r0}:B{rl}", [t for t, _ in policy])
ws.conditional_formatting.add(f"A{r0}:G{rl}", FormulaRule(formula=[f'$F{r0}="No"'], font=Font(color="999999", strike=True)))
s = rl + 2
calc_label(ws, s, 2, "Active points", f'=SUMIF(F{r0}:F{rl},"Yes",D{r0}:D{rl})', "0.0", vcol=4)
calc_label(ws, s + 1, 2, "Current step", f'=IF(D{s}>=J16,I16,IF(D{s}>=J15,I15,IF(D{s}>=J14,I14,IF(D{s}>=J13,I13,"Good standing"))))', vcol=4)
calc_label(ws, s + 2, 2, "Points until next step", f'=IF(D{s}>=J16,0,IF(D{s}>=J15,J16,IF(D{s}>=J14,J15,IF(D{s}>=J13,J14,J13))))-D{s}', "0.0", vcol=4)
note(ws, s + 4, "Edit the yellow policy tables to match your written attendance policy. Each occurrence drops off after the "
                "number of months you set (a rolling period). Expired rows turn grey.")
note(ws, s + 5, "Never assign points for FMLA leave or leave protected by a state or city sick leave law; record it as "
                "protected leave so it scores zero.")
finish(ws, [12, 34, 28, 8, 12, 9, 13, 3, 36, 10], hr, s + 7, landscape=False)
wb.save(d / "attendance-points-tracker.xlsx")
html_pdf(d / "attendance-points-tracker.pdf", "Attendance points tracker",
         info_html(["Employee", "Points as of", "Manager"]) +
         grid_html(["Date", "Occurrence type", "Details", "Points", "Expires on", "Active?", "Recorded by"],
                   [[""] * 7 for _ in range(16)]) +
         "<table class='sum'><thead><tr><th>Occurrence type</th><th>Points</th></tr></thead><tbody>" +
         "".join(f"<tr><td>{_esc(t)}</td><td>{p}</td></tr>" for t, p in policy) + "</tbody></table>"
         "<p class='note'>Active points: ______   Current step: ______________   Points drop off 12 months after the "
         "occurrence. Never assign points for FMLA leave or leave protected by a state or city sick leave law.</p>",
         landscape=False, extra_css=PDF_CSS)

# 5. Tardy and early departure log ------------------------------------------------------------
wb, ws = new_sheet("Tardy and early departure log", "Company: ____________   Location: ____________",
                   sheet_name="Tardy log")
label(ws, 3, 1, "Late after (minutes grace)", 0, vcol=2)
hr = 5
heads = ["Date", "Employee", "Scheduled start", "Arrived", "Minutes late", "Scheduled end", "Left", "Minutes left early",
         "Reason", "Manager initials"]
for c, t in enumerate(heads, 1):
    hdr(ws, hr, c, t)
r0, n = hr + 1, 40
for i in range(n):
    r = r0 + i
    z = i % 2
    put(ws, r, 1, None, fmt="m/d/yyyy", zebra=z, center=True)
    put(ws, r, 2, "", zebra=z)
    for c in (3, 4, 6, 7):
        put(ws, r, c, None, fmt=TIME, zebra=z, center=True)
    put(ws, r, 5, f'=IF(COUNT(C{r},D{r})<2,"",IF(ROUND((D{r}-C{r})*1440,0)>$B$3,ROUND((D{r}-C{r})*1440,0),0))', fmt="0",
        fill=CALC, center=True)
    put(ws, r, 8, f'=IF(COUNT(F{r},G{r})<2,"",MAX(0,ROUND((F{r}-G{r})*1440,0)))', fmt="0", fill=CALC, center=True)
    put(ws, r, 9, "", zebra=z)
    put(ws, r, 10, "", zebra=z, center=True)
rl = r0 + n - 1
s = rl + 2
calc_label(ws, s, 3, "Late arrivals", f'=COUNTIF(E{r0}:E{rl},">0")', vcol=5)
calc_label(ws, s + 1, 3, "Total minutes late", f"=SUM(E{r0}:E{rl})", vcol=5)
calc_label(ws, s, 6, "Early departures", f'=COUNTIF(H{r0}:H{rl},">0")', vcol=8)
calc_label(ws, s + 1, 6, "Total minutes left early", f"=SUM(H{r0}:H{rl})", vcol=8)
hdr(ws, s + 3, 2, "Employee")
hdr(ws, s + 3, 3, "Late arrivals")
hdr(ws, s + 3, 4, "Minutes late")
hdr(ws, s + 3, 5, "Early departures")
for i in range(10):
    r = s + 4 + i
    put(ws, r, 2, "", fill=INPUT)
    put(ws, r, 3, f'=IF(B{r}="","",COUNTIFS($B${r0}:$B${rl},B{r},$E${r0}:$E${rl},">0"))', fill=CALC, center=True)
    put(ws, r, 4, f'=IF(B{r}="","",SUMIF($B${r0}:$B${rl},B{r},$E${r0}:$E${rl}))', fill=CALC, center=True)
    put(ws, r, 5, f'=IF(B{r}="","",COUNTIFS($B${r0}:$B${rl},B{r},$H${r0}:$H${rl},">0"))', fill=CALC, center=True)
note(ws, s + 15, "Type a name in the yellow cells to total that person's rows. Pay hourly employees for the time they "
                 "actually worked; this log is for coaching and your attendance policy, not for docking pay beyond time not worked.")
finish(ws, [11, 20, 11, 11, 9, 11, 11, 10, 28, 10], hr, s + 17)
wb.save(d / "tardy-and-early-departure-log.xlsx")
html_pdf(d / "tardy-and-early-departure-log.pdf", "Tardy and early departure log",
         info_html(["Location", "Month", "Manager"]) +
         grid_html(["Date", "Employee", "Scheduled start", "Arrived", "Minutes late", "Scheduled end", "Left",
                    "Minutes early", "Reason", "Initials"], [[""] * 10 for _ in range(18)]),
         extra_css=PDF_CSS)

# 6. Employee absence report form (Word + PDF) -----------------------------------------------
absence = [
    ("fields", ["Employee name", "Job title and department", "Supervisor", "Date of this report"]),
    ("h", "Absence"),
    ("table", ["Date absent", "Scheduled shift", "Hours missed", "Paid leave used (hours and type)"],
     [["", "", "", ""] for _ in range(4)]),
    ("p", "Type of absence"),
    ("checks", ["Sick (self)", "Care of a family member", "Personal", "Bereavement", "Jury duty or court",
                "Weather or transportation", "No-call no-show", "Other"]),
    ("fields", ["How notice was given (call, text, app, in person)", "Date and time of notice", "Notice given to",
                "Expected or actual return-to-work date"]),
    ("p", "Optional: anything you want your supervisor to know. Do not include medical details; if certification is "
          "needed for protected leave, HR will ask for it separately."),
    ("fields", ["Employee comments"]),
    ("sign", ["Employee signature"]),
    ("h", "Manager use only"),
    ("checks", ["Excused", "Unexcused", "Refer to HR: may be FMLA or protected sick leave"]),
    ("fields", ["Points or occurrences recorded under policy", "Payroll: hours paid and leave balance updated"]),
    ("sign", ["Manager signature"]),
]
docx_doc(d / "employee-absence-report-form.docx", "Employee absence report", absence)
pdf_doc(d / "employee-absence-report-form.pdf", "Employee absence report", absence)

report(d)
