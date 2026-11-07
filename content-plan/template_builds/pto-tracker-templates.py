#!/usr/bin/env python3
"""Build the downloads for /templates/pto-tracker-templates/.

Run from the site root:  python content-plan/template_builds/pto-tracker-templates.py
Files go to assets/templates/pto-tracker-templates/. Files are blank and ready to fill in, with working
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


# ============================================================== /templates/pto-tracker-templates/
d = out("pto-tracker-templates")
D = datetime.date
MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October",
          "November", "December"]
# US federal holidays as observed (OPM), used as a default company holiday list. Edit to match your own.
FED = [(D(2026, 1, 1), "New Year's Day"), (D(2026, 1, 19), "Martin Luther King Jr. Day"), (D(2026, 2, 16), "Washington's Birthday"),
       (D(2026, 5, 25), "Memorial Day"), (D(2026, 6, 19), "Juneteenth"), (D(2026, 7, 3), "Independence Day (observed)"),
       (D(2026, 9, 7), "Labor Day"), (D(2026, 10, 12), "Columbus Day"), (D(2026, 11, 11), "Veterans Day"),
       (D(2026, 11, 26), "Thanksgiving Day"), (D(2026, 12, 25), "Christmas Day"),
       (D(2027, 1, 1), "New Year's Day"), (D(2027, 1, 18), "Martin Luther King Jr. Day"), (D(2027, 2, 15), "Washington's Birthday"),
       (D(2027, 5, 31), "Memorial Day"), (D(2027, 6, 18), "Juneteenth (observed)"), (D(2027, 7, 5), "Independence Day (observed)"),
       (D(2027, 9, 6), "Labor Day"), (D(2027, 10, 11), "Columbus Day"), (D(2027, 11, 11), "Veterans Day"),
       (D(2027, 11, 25), "Thanksgiving Day"), (D(2027, 12, 24), "Christmas Day (observed)")]
PDF_CSS = (".info{display:grid;grid-template-columns:repeat(3,1fr);gap:8px 22px;margin:0 0 14px;font-size:10pt}"
           ".info div{border-bottom:1px solid #333;padding:12px 0 2px;color:#333}"
           ".note{font-size:9pt;color:#444;margin-top:12px} .foot{position:static;margin-top:22px} td{height:18px}"
           ".sum{width:auto;min-width:46%;margin-top:14px} .sum th{background:#e9e9e3;color:#111;text-align:left}"
           ".grid{table-layout:fixed} .grid th,.grid td{padding:3px 1px;text-align:center;font-size:7.5pt} .grid td{height:20px}"
           ".grid td:first-child,.grid th:first-child{text-align:left;padding-left:5px;width:110px}")


def info_html(labels):
    return "<div class='info'>" + "".join(f"<div>{_esc(l)}:</div>" for l in labels) + "</div>"


def sum_html(rows):
    return "<table class='sum'><tbody>" + "".join(f"<tr><th>{_esc(a)}</th><td style='min-width:110px'>{_esc(b)}</td></tr>" for a, b in rows) + "</tbody></table>"


def pdf(path, title, body, landscape=False):
    return html_pdf(path, title, body, landscape=landscape, extra_css=PDF_CSS)


# 1. PTO tracker with accrual per pay period ------------------------------------------------
wb, ws = new_sheet("PTO tracker", "Company: ____________   Manager: ____________", sheet_name="PTO tracker")
label(ws, 3, 1, "Employee", "", vcol=2)
label(ws, 4, 1, "First pay period ends", D(2026, 1, 9), vcol=2, fmt=DATE)
label(ws, 5, 1, "Days between paydays (7, 14)", 14, vcol=2)
label(ws, 3, 4, "PTO hours accrued per pay period", 3.08, vcol=6, fmt="0.00")
label(ws, 4, 4, "Carryover from last year (hours)", 0, vcol=6, fmt="0.00")
label(ws, 5, 4, "Maximum balance (0 = no cap)", 0, vcol=6, fmt="0.00")
label(ws, 6, 1, "Show balance as of", "=TODAY()", vcol=2, fmt=DATE)
hr = 8
heads = ["Pay period", "Period ends", "Hours accrued", "PTO hours used", "Dates used and notes", "Balance after period"]
for c, t in enumerate(heads, 1):
    hdr(ws, hr, c, t)
r0, n = hr + 1, 27
for i in range(n):
    r = r0 + i
    z = i % 2
    prev = "$F$4" if i == 0 else f"N(F{r - 1})"
    put(ws, r, 1, i + 1, center=True, zebra=z)
    put(ws, r, 2, f"=$B$4+$B$5*{i}", fmt=DATE, center=True, zebra=z)
    put(ws, r, 3, f"=IF(B{r}>$B$6,0,ROUND(IF($F$5=0,$F$3,MAX(0,MIN($F$3,$F$5-{prev}))),2))", fmt="0.00", fill=CALC, center=True)
    put(ws, r, 4, None, fmt="0.00", center=True, zebra=z)
    put(ws, r, 5, "", zebra=z)
    put(ws, r, 6, f'=IF(AND(B{r}>$B$6,N(D{r})=0),"",ROUND({prev}+C{r}-N(D{r}),2))', fmt="0.00", fill=CALC, center=True)
rl = r0 + n - 1
ws.conditional_formatting.add(f"F{r0}:F{rl}", FormulaRule(formula=[f'AND(F{r0}<>"",F{r0}<0)'], fill=PatternFill("solid", fgColor="F8D0CC")))
s = rl + 2
calc_label(ws, s, 4, "Total accrued this year", f"=SUM(C{r0}:C{rl})", "0.00", vcol=6)
calc_label(ws, s + 1, 4, "Total PTO used this year", f"=SUM(D{r0}:D{rl})", "0.00", vcol=6)
calc_label(ws, s + 2, 4, "Balance as of the date above",
           f"=$F$4+F{s}-F{s + 1}", "0.00", vcol=6)
note(ws, s + 4, "Hours accrue only for pay periods that ended on or before the 'show balance as of' date, so future "
                "rows stay empty until you enter approved PTO in them. A negative balance turns red. Common rates: 80 hours a year "
                "= 3.08 per biweekly period, 120 = 4.62, 160 = 6.15.")
finish(ws, [12, 13, 11, 11, 36, 14], hr, s + 6, landscape=False)
wb.save(d / "pto-tracker-accrual-per-pay-period.xlsx")
pdf(d / "pto-tracker-accrual-per-pay-period.pdf", "PTO tracker",
    info_html(["Employee", "Year", "Accrual per pay period", "Carryover", "Maximum balance", "Manager"]) +
    grid_html(["Pay period", "Period ends", "Hours accrued", "PTO hours used", "Dates used and notes", "Balance"],
              [[str(i)] + [""] * 5 for i in range(1, 27)]))

# 2. Paid sick leave tracker (accrual per hours worked) ------------------------------------
wb, ws = new_sheet("Paid sick leave tracker", "Company: ____________   State or city law: ____________",
                   sheet_name="Sick leave tracker")
label(ws, 3, 1, "Employee", "", vcol=2)
label(ws, 4, 1, "Carryover from last year (hours)", 0, vcol=2, fmt="0.00")
label(ws, 3, 4, "Earn 1 hour for every ... hours worked", 30, vcol=6)
label(ws, 4, 4, "Maximum balance (accrual cap)", 80, vcol=6, fmt="0.00")
label(ws, 5, 4, "Annual use limit (hours, 0 = none)", 40, vcol=6, fmt="0.00")
hr = 7
heads = ["Period ends", "Hours worked", "Sick hours earned", "Sick hours used", "Balance", "Used this year", "Check"]
for c, t in enumerate(heads, 1):
    hdr(ws, hr, c, t)
r0, n = hr + 1, 30
for i in range(n):
    r = r0 + i
    z = i % 2
    prev = "$B$4" if i == 0 else f"E{r - 1}"
    used_prev = "0" if i == 0 else f"F{r - 1}"
    put(ws, r, 1, None, fmt=DATE, center=True, zebra=z)
    put(ws, r, 2, None, fmt="0.00", center=True, zebra=z)
    put(ws, r, 3, f"=ROUND(MAX(0,MIN(N(B{r})/$F$3,$F$4-{prev})),2)", fmt="0.00", fill=CALC, center=True)
    put(ws, r, 4, None, fmt="0.00", center=True, zebra=z)
    put(ws, r, 5, f"=ROUND({prev}+C{r}-N(D{r}),2)", fmt="0.00", fill=CALC, center=True)
    put(ws, r, 6, f"={used_prev}+N(D{r})", fmt="0.00", fill=CALC, center=True)
    put(ws, r, 7, f'=IF(E{r}<0,"Balance below zero",IF(AND($F$5>0,F{r}>$F$5),"Over annual use limit",""))', fill=CALC)
rl = r0 + n - 1
s = rl + 2
calc_label(ws, s, 4, "Sick hours earned this year", f"=SUM(C{r0}:C{rl})", "0.00", vcol=6)
calc_label(ws, s + 1, 4, "Sick hours used this year", f"=SUM(D{r0}:D{rl})", "0.00", vcol=6)
calc_label(ws, s + 2, 4, "Balance available", f"=$B$4+F{s}-F{s + 1}", "0.00", vcol=6)
note(ws, s + 4, "Defaults follow California's accrual method: 1 hour per 30 hours worked, a cap of 80 hours, and use that "
                "may be limited to 40 hours a year. Other states and cities set different rates and caps; change the yellow cells.")
note(ws, s + 5, "Many laws require you to show the available balance on the pay stub or a notice given with it, and to keep "
                "these records for several years.")
finish(ws, [14, 12, 12, 12, 11, 12, 22], hr, s + 7, landscape=False)
wb.save(d / "paid-sick-leave-tracker.xlsx")
pdf(d / "paid-sick-leave-tracker.pdf", "Paid sick leave tracker",
    info_html(["Employee", "Year", "Accrual rate (1 hour per ... worked)", "Carryover", "Accrual cap", "Annual use limit"]) +
    grid_html(["Period ends", "Hours worked", "Sick hours earned", "Sick hours used", "Balance", "Used this year"],
              [[""] * 6 for _ in range(26)]))

# 3. Team PTO calendar (12 month tabs) -------------------------------------------------------
wb = Workbook()
sm = wb.active
sm.title = "Year summary"
sm["A1"] = "Team PTO calendar"
sm["A1"].font = Font(bold=True, size=16)
sm["A2"] = "Type names once here; they appear on every month tab. Mark days off on the month tabs."
sm["A2"].font = Font(size=11, color="444444")
label(sm, 3, 1, "Year", 2027, vcol=2)
label(sm, 4, 1, "Hours in a workday", 8, vcol=2)
hr = 6
PCODES = [("V", "Vacation or PTO"), ("S", "Sick"), ("P", "Personal"), ("F", "FMLA or other leave"), ("U", "Unpaid")]
for c, t in enumerate(["Employee", "Team or role", "Annual PTO (days)"] + [f"{c} days" for c, _ in PCODES] +
                      ["PTO days left", "PTO hours used"], 1):
    hdr(sm, hr, c, t)
NE = 20
for i in range(NE):
    r = hr + 1 + i
    put(sm, r, 1, "", fill=INPUT)
    put(sm, r, 2, "", fill=INPUT)
    put(sm, r, 3, None, fill=INPUT, center=True)
    for j, (code, _) in enumerate(PCODES):
        refs = "+".join(f"'{m[:3]}'!{L(33 + j)}{r}" for m in MONTHS)
        put(sm, r, 4 + j, f"={refs}", fill=CALC, center=True)
    put(sm, r, 9, f'=IF(C{r}="","",C{r}-D{r}-F{r})', fill=CALC, center=True)
    put(sm, r, 10, f"=(D{r}+F{r})*$B$4", fill=CALC, center=True)
hl0 = hr + NE + 3
hdr(sm, hl0, 1, "Company holidays (edit)")
hdr(sm, hl0, 2, "Holiday")
for i, (dt_, nm) in enumerate(FED):
    put(sm, hl0 + 1 + i, 1, dt_, fmt="ddd m/d/yyyy", fill=INPUT)
    put(sm, hl0 + 1 + i, 2, nm, fill=INPUT)
hl1 = hl0 + len(FED) + 8
for i in range(len(FED) + 1, len(FED) + 8):
    put(sm, hl0 + i, 1, None, fmt="ddd m/d/yyyy", fill=INPUT)
    put(sm, hl0 + i, 2, "", fill=INPUT)
HOL = f"'Year summary'!$A${hl0 + 1}:$A${hl1}"
note(sm, hl1 + 2, "Codes on the month tabs: " + "   ".join(f"{c} = {t}" for c, t in PCODES) +
     ". PTO days left = annual PTO - vacation/PTO days - personal days. Half days: type the code and adjust the total by hand.")
note(sm, hl1 + 3, "The holiday list starts with the 2026 and 2027 US federal holidays as observed; replace it with the "
                  "holidays your business actually closes for.")
finish(sm, [22, 16, 12, 9, 9, 9, 9, 9, 12, 12], hr, hl1 + 5, landscape=False)
for mi, mname in enumerate(MONTHS):
    ws = wb.create_sheet(mname[:3])
    ws["A1"] = f"Team PTO calendar: {mname}"
    ws["A1"].font = Font(bold=True, size=16)
    ws["A2"] = f"='Year summary'!B3"
    ws["A2"].font = Font(size=11, color="444444")
    ws["A2"].alignment = Alignment(horizontal="left")
    # rows: 4 day numbers, 5 weekday, 6 holiday marker, 7.. employees
    hdr(ws, 4, 1, "Employee")
    hdr(ws, 5, 1, "")
    put(ws, 6, 1, "Holiday", bold=True, fill=GREY)
    for k in range(31):
        c = 2 + k
        dexpr = f"DATE('Year summary'!$B$3,{mi + 1},{k + 1})"
        hdr(ws, 4, c, f'=IF(MONTH({dexpr})<>{mi + 1},"",{k + 1})')
        hdr(ws, 5, c, f'=IF(MONTH({dexpr})<>{mi + 1},"",LEFT(TEXT({dexpr},"ddd"),2))')
        put(ws, 6, c, f'=IF({L(c)}$4="","",IF(COUNTIF({HOL},{dexpr})>0,"H",""))', fill=GREY, center=True)
    for j, (code, _) in enumerate(PCODES):
        hdr(ws, 4, 33 + j, code)
        hdr(ws, 5, 33 + j, "")
        put(ws, 6, 33 + j, "", fill=GREY)
    for i in range(NE):
        r = 7 + i
        sr = hr + 1 + i
        put(ws, r, 1, f"=IF('Year summary'!A{sr}=\"\",\"\",'Year summary'!A{sr})", bold=True, zebra=i % 2)
        for k in range(31):
            put(ws, r, 2 + k, "", center=True)
        for j, (code, _) in enumerate(PCODES):
            put(ws, r, 33 + j, f'=COUNTIF(B{r}:AF{r},"{code}")', fill=CALC, center=True)
    # summary rows must sit on the same row numbers as the Year summary employees (7..26) -> they do (hr + 1 = 7)
    rl = 7 + NE - 1
    put(ws, rl + 1, 1, "People off", bold=True, fill=GREY)
    for k in range(31):
        col = L(2 + k)
        put(ws, rl + 1, 2 + k, f'=IF({col}$4="","",COUNTA({col}7:{col}{rl}))', fill=CALC, center=True, bold=True)
    grid = f"B7:AF{rl}"
    ws.conditional_formatting.add(grid, FormulaRule(formula=['B$4=""'], fill=PatternFill("solid", fgColor="8C8C86"), stopIfTrue=True))
    for code, col in [("V", "D9F2C2"), ("P", "D9F2C2"), ("S", "F8D0CC"), ("F", "CFE0F7"), ("U", "F6B98E")]:
        ws.conditional_formatting.add(grid, FormulaRule(formula=[f'UPPER(B7)="{code}"'], fill=PatternFill("solid", fgColor=col)))
    ws.conditional_formatting.add(grid, FormulaRule(formula=['OR(B$5="Sa",B$5="Su",B$6="H")'],
                                                     fill=PatternFill("solid", fgColor="E4E4DD")))
    ws.conditional_formatting.add(f"B{rl + 1}:AF{rl + 1}", FormulaRule(formula=[f"AND(B{rl + 1}<>\"\",B{rl + 1}>=3)"],
                                                                       fill=PatternFill("solid", fgColor="F8D0CC")))
    choices(ws, grid, [c for c, _ in PCODES])
    note(ws, rl + 3, "Codes: " + "   ".join(f"{c} = {t}" for c, t in PCODES) +
         ". Weekends and company holidays are shaded. 'People off' turns red at 3 or more.")
    finish(ws, [20] + [3.6] * 31 + [4] * 5, 5, rl + 5, freeze="B7")
wb.save(d / "team-pto-calendar.xlsx")
pdf(d / "team-pto-calendar.pdf", "Team PTO calendar",
    info_html(["Month and year", "Team", "Manager"]) +
    grid_html(["Employee"] + [str(i) for i in range(1, 32)], [[""] * 32 for _ in range(15)] + [["People off"] + [""] * 31], cls="grid") +
    "<p class='note'>Codes: V = Vacation or PTO   S = Sick   P = Personal   F = FMLA or other leave   U = Unpaid   H = Company holiday</p>",
    landscape=True)

# 4. Team PTO balance summary --------------------------------------------------------------
wb, ws = new_sheet("Team PTO balance summary", "Company: ____________   Prepared by: ____________",
                   sheet_name="PTO balances")
label(ws, 3, 1, "Balances as of", D(2026, 11, 7), vcol=2, fmt=DATE)
hr = 5
heads = ["Employee", "Hire date", "Annual PTO (hours)", "Carryover (hours)", "Earned so far this year", "Used this year",
         "Approved, not yet taken", "Balance today", "Available to request", "Share of year's PTO used", "Balance cap", "Flag"]
for c, t in enumerate(heads, 1):
    hdr(ws, hr, c, t)
r0, n = hr + 1, 25
for i in range(n):
    r = r0 + i
    z = i % 2
    put(ws, r, 1, "", zebra=z)
    put(ws, r, 2, None, fmt=DATE, center=True, zebra=z)
    for c in (3, 4, 5, 6, 7):
        put(ws, r, c, None, fmt="0.00", center=True, zebra=z)
    put(ws, r, 8, f'=IF(A{r}="","",N(D{r})+N(E{r})-N(F{r}))', fmt="0.00", fill=CALC, center=True)
    put(ws, r, 9, f'=IF(A{r}="","",H{r}-N(G{r}))', fmt="0.00", fill=CALC, center=True)
    put(ws, r, 10, f'=IF(OR(A{r}="",N(C{r})+N(D{r})=0),"",N(F{r})/(N(C{r})+N(D{r})))', fmt="0%", fill=CALC, center=True)
    put(ws, r, 11, None, fmt="0.00", center=True, zebra=z)
    put(ws, r, 12, f'=IF(A{r}="","",IF(I{r}<0,"Over-approved",IF(AND(N(K{r})>0,H{r}>=K{r}),"At cap, not earning",'
                   f'IF(AND(MONTH($B$3)>=10,J{r}<0.5),"Low use, plan time off",""))))', fill=CALC)
rl = r0 + n - 1
put(ws, rl + 1, 1, "Team totals", bold=True, fill=GREY)
put(ws, rl + 1, 2, "", fill=GREY)
for c in range(3, 10):
    put(ws, rl + 1, c, f"=SUM({L(c)}{r0}:{L(c)}{rl})", fmt="0.00", fill=CALC, bold=True, center=True)
for c in (10, 11, 12):
    put(ws, rl + 1, c, "", fill=GREY)
note(ws, rl + 3, "Balance today = carryover + earned so far - used. Available to request = balance - approved time not yet "
                 "taken. Flags: over-approved, at the balance cap (so not earning), or under half of the year's PTO used from October on.")
note(ws, rl + 4, "Unused PTO may have to be paid out at termination in some states, and California does not allow "
                 "use-it-or-lose-it vacation policies. Check your state before you set an expiry rule.")
finish(ws, [20, 11, 11, 11, 11, 10, 11, 10, 11, 11, 9, 22], hr, rl + 6)
wb.save(d / "team-pto-balance-summary.xlsx")
pdf(d / "team-pto-balance-summary.pdf", "Team PTO balance summary",
    info_html(["Balances as of", "Company", "Prepared by"]) +
    grid_html(["Employee", "Hire date", "Annual PTO", "Carryover", "Earned this year", "Used this year",
               "Approved, not taken", "Balance today", "Available"], [[""] * 9 for _ in range(18)] + [["Totals"] + [""] * 8]),
    landscape=True)

# 5. Vacation tracker (days, separate banks) -----------------------------------------------
wb, ws = new_sheet("Vacation tracker", "Company: ____________   Manager: ____________", sheet_name="Vacation tracker")
label(ws, 3, 1, "Employee", "", vcol=2)
label(ws, 4, 1, "Year", 2026, vcol=2)
hdr(ws, 6, 1, "Leave bank")
hdr(ws, 6, 2, "Days for the year")
hdr(ws, 6, 3, "Carryover days")
hdr(ws, 6, 4, "Days used")
hdr(ws, 6, 5, "Days remaining")
BANKS = ["Vacation", "Sick", "Personal", "Unpaid"]
for i, b in enumerate(BANKS):
    r = 7 + i
    put(ws, r, 1, b, bold=True, fill=GREY)
    put(ws, r, 2, None if b != "Unpaid" else "", fmt="0.0", fill=INPUT, center=True)
    put(ws, r, 3, None if b != "Unpaid" else "", fmt="0.0", fill=INPUT, center=True)
    put(ws, r, 4, f'=SUMIF($C$14:$C$53,A{r},$F$14:$F$53)', fmt="0.0", fill=CALC, center=True)
    put(ws, r, 5, f'=IF(A{r}="Unpaid","",N(B{r})+N(C{r})-D{r})', fmt="0.0", fill=CALC, center=True, bold=True)
hr = 13
heads = ["First day", "Last day", "Type", "Workdays (auto)", "Days (override, e.g. 0.5)", "Days taken", "Notes",
         "Approved by"]
for c, t in enumerate(heads, 1):
    hdr(ws, hr, c, t)
# holiday list to the right
hdr(ws, 5, 10, "Company holidays (edit)")
for i, (dt_, nm) in enumerate(FED):
    put(ws, 6 + i, 10, dt_, fmt="ddd m/d/yyyy", fill=INPUT)
    put(ws, 6 + i, 11, nm, fill=INPUT)
for i in range(len(FED), len(FED) + 6):
    put(ws, 6 + i, 10, None, fmt="ddd m/d/yyyy", fill=INPUT)
    put(ws, 6 + i, 11, "", fill=INPUT)
hend = 6 + len(FED) + 5
r0, n = hr + 1, 40
for i in range(n):
    r = r0 + i
    z = i % 2
    put(ws, r, 1, None, fmt="ddd m/d/yyyy", center=True, zebra=z)
    put(ws, r, 2, None, fmt="ddd m/d/yyyy", center=True, zebra=z)
    put(ws, r, 3, "", zebra=z, center=True)
    put(ws, r, 4, f'=IF(COUNT(A{r})=0,"",NETWORKDAYS(A{r},IF(COUNT(B{r})=0,A{r},B{r}),$J$6:$J${hend}))', fmt="0", fill=CALC, center=True)
    put(ws, r, 5, None, fmt="0.0", zebra=z, center=True)
    put(ws, r, 6, f'=IF(COUNT(A{r})=0,0,IF(COUNT(E{r})=1,E{r},D{r}))', fmt="0.0", fill=CALC, center=True)
    put(ws, r, 7, "", zebra=z)
    put(ws, r, 8, "", zebra=z)
rl = r0 + n - 1
choices(ws, f"C{r0}:C{rl}", BANKS)
note(ws, rl + 2, "Workdays count Monday to Friday and skip the company holidays on the right. Type a number in the override "
                 "column for half days or weekend shifts. Remaining days turn the bank balance for each type.")
finish(ws, [15, 15, 11, 10, 12, 10, 28, 12, 3, 16, 28], hr, rl + 4, freeze="A14")
wb.save(d / "vacation-tracker.xlsx")
pdf(d / "vacation-tracker.pdf", "Vacation tracker",
    info_html(["Employee", "Year", "Manager"]) +
    "<table class='sum' style='width:100%'><thead><tr><th>Leave bank</th><th>Days for the year</th><th>Carryover</th>"
    "<th>Days used</th><th>Days remaining</th></tr></thead><tbody>" +
    "".join(f"<tr><th>{b}</th><td></td><td></td><td></td><td></td></tr>" for b in BANKS[:3]) + "</tbody></table><br>" +
    grid_html(["First day", "Last day", "Type", "Days taken", "Notes", "Approved by"], [[""] * 6 for _ in range(18)]))

# 6. PTO by years of service (tiered accrual) -----------------------------------------------
wb, ws = new_sheet("PTO accrual by years of service", "Company: ____________", sheet_name="PTO by tenure")
label(ws, 3, 1, "Calculate as of", D(2027, 1, 1), vcol=2, fmt=DATE)
label(ws, 4, 1, "Pay periods per year", 26, vcol=2)
hdr(ws, 3, 10, "Years of service (at least)")
hdr(ws, 3, 11, "PTO hours per year")
tiers = [(0, 80), (2, 120), (5, 160), (10, 200)]
for i, (y, h) in enumerate(tiers):
    put(ws, 4 + i, 10, y, fill=INPUT, center=True)
    put(ws, 4 + i, 11, h, fill=INPUT, center=True)
for i in range(len(tiers), 6):
    put(ws, 4 + i, 10, None, fill=INPUT, center=True)
    put(ws, 4 + i, 11, None, fill=INPUT, center=True)
hr = 11
heads = ["Employee", "Hire date", "Full years of service", "PTO hours per year", "Accrual per pay period",
         "PTO days per year (8-hour days)", "Next increase on", "New annual hours"]
for c, t in enumerate(heads, 1):
    hdr(ws, hr, c, t)
r0, n = hr + 1, 25
for i in range(n):
    r = r0 + i
    z = i % 2
    put(ws, r, 1, "", zebra=z)
    put(ws, r, 2, None, fmt=DATE, center=True, zebra=z)
    put(ws, r, 3, f'=IF(COUNT(B{r})=0,"",DATEDIF(B{r},$B$3,"y"))', fill=CALC, center=True)
    put(ws, r, 4, f'=IF(C{r}="","",LOOKUP(C{r},$J$4:$J$9,$K$4:$K$9))', fill=CALC, center=True)
    put(ws, r, 5, f'=IF(D{r}="","",ROUND(D{r}/$B$4,2))', fmt="0.00", fill=CALC, center=True)
    put(ws, r, 6, f'=IF(D{r}="","",D{r}/8)', fmt="0.0", fill=CALC, center=True)
    put(ws, r, 7, f'=IF(C{r}="","",IFERROR(EDATE(B{r},12*SMALL($J$4:$J$9,COUNTIF($J$4:$J$9,"<="&C{r})+1)),"Top tier"))',
        fmt=DATE, fill=CALC, center=True)
    put(ws, r, 8, f'=IF(C{r}="","",IFERROR(INDEX($K$4:$K$9,MATCH(SMALL($J$4:$J$9,COUNTIF($J$4:$J$9,"<="&C{r})+1),$J$4:$J$9,0)),""))',
        fill=CALC, center=True)
rl = r0 + n - 1
note(ws, rl + 2, "Edit the tier table (top right) to match your policy; keep it sorted from fewest to most years. "
                 "Accrual per pay period = annual hours / pay periods per year, rounded to 2 decimals.")
finish(ws, [20, 12, 11, 11, 12, 13, 14, 12, 3, 14, 13], hr, rl + 4)
wb.save(d / "pto-accrual-by-years-of-service.xlsx")
pdf(d / "pto-accrual-by-years-of-service.pdf", "PTO accrual by years of service",
    info_html(["Company", "Calculate as of", "Pay periods per year"]) +
    "<table class='sum'><thead><tr><th>Years of service (at least)</th><th>PTO hours per year</th></tr></thead><tbody>" +
    "".join("<tr><td></td><td></td></tr>" for _ in range(5)) + "</tbody></table><br>" +
    grid_html(["Employee", "Hire date", "Years of service", "PTO hours per year", "Accrual per pay period", "Next increase on"],
              [[""] * 6 for _ in range(16)]), landscape=True)

report(d)
