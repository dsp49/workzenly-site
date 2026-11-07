#!/usr/bin/env python3
"""Build the downloads for /templates/timesheet-templates/.

Run from the site root:  python content-plan/template_builds/timesheet-templates.py
Files go to assets/templates/timesheet-templates/. Files are blank and ready to fill in, with working
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


# ============================================================== /templates/timesheet-templates/
d = out("timesheet-templates")
DAYS7 = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
PDF_CSS = (".info{display:grid;grid-template-columns:repeat(3,1fr);gap:8px 22px;margin:0 0 14px;font-size:10pt}"
           ".info div{border-bottom:1px solid #333;padding:14px 0 2px;color:#333}"
           ".sum{width:auto;min-width:46%;margin-top:14px} .sum th{background:#e9e9e3;color:#111;text-align:left;width:60%}"
           ".sum td{min-width:90px} .note{font-size:9pt;color:#444;margin-top:12px} td{height:20px} .foot{position:static;margin-top:22px}")


def info_html(labels):
    return "<div class='info'>" + "".join(f"<div>{_esc(l)}:</div>" for l in labels) + "</div>"


def sum_html(rows):
    return "<table class='sum'><tbody>" + "".join(f"<tr><th>{_esc(a)}</th><td>{_esc(b)}</td></tr>" for a, b in rows) + "</tbody></table>"


def sheet_pdf(path, title, info, headers, rows, summary=None, sign=None, notes=None, landscape=True, subtitle=None,
              compact=False):
    body = info_html(info) + grid_html(headers, rows)
    if summary:
        body += sum_html(summary)
    if notes:
        body += f"<p class='note'>{_esc(notes)}</p>"
    if sign:
        body += sign_html(sign)
    return html_pdf(path, title, body, subtitle=subtitle, landscape=landscape,
                    extra_css=PDF_CSS + (" td,th{padding:2px 6px;font-size:8.5pt} td{height:11px} .info div{padding-top:8px} h1{font-size:17pt} .sign{margin-top:18px;gap:16px 36px}" if compact else ""))


def hours_in_out(a, b, brk=None):
    """Excel: decimal hours from two time cells, overnight-safe, minus an unpaid break in minutes."""
    minus = f"-N({brk})/60" if brk else ""
    return f"ROUND(IF(COUNT({a},{b})<2,0,24*MOD({b}-{a},1){minus}),2)"


CERT = ("I certify that this timesheet is a true and complete record of all the time I worked, "
        "including any work before or after my shift and during meal breaks.")
WEEK = datetime.date(2026, 11, 9)       # a Monday
PERIOD = datetime.date(2026, 11, 1)     # a Sunday

# 1. Weekly timesheet with lunch breaks and overtime -----------------------------------------
wb, ws = new_sheet("Weekly timesheet", "Company: ____________   Department: ____________   Supervisor: ____________",
                   sheet_name="Weekly timesheet")
label(ws, 3, 1, "Employee", "", vcol=2)
label(ws, 4, 1, "Week starting", WEEK, vcol=2, fmt=DATE)
label(ws, 3, 5, "Hourly rate", None, vcol=7, fmt=MONEY)
label(ws, 4, 5, "Overtime after (hours/week)", 40, vcol=7)
hr = 6
for c, t in enumerate(["Day", "Date", "Time in", "Lunch out", "Lunch in", "Time out", "Hours worked", "Notes"], 1):
    hdr(ws, hr, c, t)
for k in range(7):
    r = hr + 1 + k
    z = k % 2
    put(ws, r, 1, f'=TEXT(B{r},"dddd")', zebra=z, bold=True)
    put(ws, r, 2, f"=$B$4+{k}", fmt="m/d", zebra=z, center=True)
    for c in range(3, 7):
        put(ws, r, c, None, fmt=TIME, zebra=z, center=True)
    put(ws, r, 7, "=ROUND(24*IF(COUNT(C{0},F{0})<2,0,IF(COUNT(D{0},E{0})=2,MOD(D{0}-C{0},1)+MOD(F{0}-E{0},1),"
                  "MOD(F{0}-C{0},1))),2)".format(r), fmt="0.00", fill=CALC, center=True)
    put(ws, r, 8, "", zebra=z)
put(ws, 14, 1, "Total hours", bold=True, fill=GREY)
for c in range(2, 7):
    put(ws, 14, c, "", fill=GREY)
put(ws, 14, 7, "=SUM(G7:G13)", fmt="0.00", fill=CALC, bold=True, center=True)
put(ws, 14, 8, "", fill=GREY)
calc_label(ws, 16, 6, "Regular hours", "=MIN(G14,G4)", "0.00", vcol=7)
calc_label(ws, 17, 6, "Overtime hours (1.5x)", "=MAX(0,G14-G4)", "0.00", vcol=7)
calc_label(ws, 18, 6, "Regular pay", "=ROUND(G16*G3,2)", MONEY, vcol=7)
calc_label(ws, 19, 6, "Overtime pay", "=ROUND(G17*G3*1.5,2)", MONEY, vcol=7)
calc_label(ws, 20, 6, "Gross pay", "=G18+G19", MONEY, vcol=7)
note(ws, 22, "Type times like 7:00 AM or 15:30. Leave the lunch cells blank on days without an unpaid meal break. "
             "Short rest breaks are paid under federal law, so do not clock out for them.")
note(ws, 23, CERT)
end = sign_lines(ws, 25, ["Employee signature", "Supervisor approval"], gap=3)
finish(ws, [14, 10, 11, 11, 11, 11, 13, 28], hr, end + 1, landscape=False)
wb.save(d / "weekly-timesheet-with-lunch-breaks.xlsx")
sheet_pdf(d / "weekly-timesheet-with-lunch-breaks.pdf", "Weekly timesheet",
          ["Employee", "Week starting", "Hourly rate", "Company", "Department", "Supervisor"],
          ["Day", "Date", "Time in", "Lunch out", "Lunch in", "Time out", "Hours worked"],
          [[x, "", "", "", "", "", ""] for x in DAYS7] + [["Total hours", "", "", "", "", "", ""]],
          summary=[("Regular hours", ""), ("Overtime hours (over 40)", ""), ("Gross pay", "")],
          notes=CERT, sign=["Employee signature", "Supervisor approval"], landscape=False)

# 2. Biweekly timesheet with overtime by workweek ---------------------------------------------
wb, ws = new_sheet("Biweekly timesheet", "Company: ____________   Department: ____________   Supervisor: ____________",
                   sheet_name="Biweekly timesheet")
label(ws, 3, 1, "Employee", "", vcol=2)
label(ws, 4, 1, "Pay period starts", PERIOD, vcol=2, fmt=DATE)
label(ws, 3, 6, "Hourly rate", None, vcol=7, fmt=MONEY)
label(ws, 4, 6, "Overtime after (hours/week)", 40, vcol=7)
hr = 6
for c, t in enumerate(["Week", "Day", "Date", "Time in", "Time out", "Unpaid break (minutes)", "Hours worked", "Notes"], 1):
    hdr(ws, hr, c, t)
rows_wk = {}
r = hr + 1
for w in (1, 2):
    first = r
    for k in range(7):
        z = k % 2
        put(ws, r, 1, f"Week {w}", zebra=z, center=True)
        put(ws, r, 2, f'=TEXT(C{r},"ddd")', zebra=z, center=True, bold=True)
        put(ws, r, 3, f"=$B$4+{(w - 1) * 7 + k}", fmt="m/d", zebra=z, center=True)
        put(ws, r, 4, None, fmt=TIME, zebra=z, center=True)
        put(ws, r, 5, None, fmt=TIME, zebra=z, center=True)
        put(ws, r, 6, None, fmt="0", zebra=z, center=True)
        put(ws, r, 7, "=" + hours_in_out(f"D{r}", f"E{r}", f"F{r}"), fmt="0.00", fill=CALC, center=True)
        put(ws, r, 8, "", zebra=z)
        r += 1
    put(ws, r, 1, f"Week {w} total", bold=True, fill=GREY)
    for c in range(2, 7):
        put(ws, r, c, "", fill=GREY)
    put(ws, r, 7, f"=SUM(G{first}:G{r - 1})", fmt="0.00", fill=CALC, bold=True, center=True)
    put(ws, r, 8, "", fill=GREY)
    rows_wk[w] = r
    r += 1
s = r + 1
for c, t in [(5, ""), (6, "Week 1"), (7, "Week 2"), (8, "Pay period")]:
    hdr(ws, s, c, t)
w1, w2 = rows_wk[1], rows_wk[2]
lines = [("Regular hours", f"=MIN(G{w1},$G$4)", f"=MIN(G{w2},$G$4)", "0.00"),
         ("Overtime hours (1.5x)", f"=MAX(0,G{w1}-$G$4)", f"=MAX(0,G{w2}-$G$4)", "0.00")]
for i, (t, a, b, f) in enumerate(lines):
    rr = s + 1 + i
    put(ws, rr, 5, t, bold=True, fill=GREY)
    put(ws, rr, 6, a, fmt=f, fill=CALC, center=True)
    put(ws, rr, 7, b, fmt=f, fill=CALC, center=True)
    put(ws, rr, 8, f"=F{rr}+G{rr}", fmt=f, fill=CALC, center=True, bold=True)
calc_label(ws, s + 4, 6, "Regular pay", f"=ROUND(H{s + 1}*G3,2)", MONEY, vcol=8)
calc_label(ws, s + 5, 6, "Overtime pay", f"=ROUND(H{s + 2}*G3*1.5,2)", MONEY, vcol=8)
calc_label(ws, s + 6, 6, "Gross pay", f"=H{s + 4}+H{s + 5}", MONEY, vcol=8)
note(ws, s + 8, "Overtime is counted for each workweek on its own, never averaged across the two weeks. "
                "Set 'Pay period starts' to the first day of your workweek.")
note(ws, s + 9, CERT)
end = sign_lines(ws, s + 11, ["Employee signature", "Supervisor approval"], gap=3)
finish(ws, [13, 8, 9, 11, 11, 12, 12, 26], hr, end + 1, landscape=False)
wb.save(d / "biweekly-timesheet.xlsx")
bi_rows = []
for w in (1, 2):
    bi_rows += [[f"Week {w}", x[:3], "", "", "", "", ""] for x in ["Sunday"] + DAYS7[:6]]
    bi_rows.append([f"Week {w} total", "", "", "", "", "", ""])
sheet_pdf(d / "biweekly-timesheet.pdf", "Biweekly timesheet",
          ["Employee", "Pay period", "Hourly rate", "Company", "Department", "Supervisor"],
          ["Week", "Day", "Date", "Time in", "Time out", "Unpaid break (min)", "Hours worked"], bi_rows,
          summary=[("Week 1 overtime hours (over 40)", ""), ("Week 2 overtime hours (over 40)", ""),
                   ("Total regular hours", ""), ("Gross pay", "")],
          notes="Overtime is counted for each workweek on its own, never averaged across the two weeks. " + CERT,
          sign=["Employee signature", "Supervisor approval"], landscape=False)

# 3. California daily overtime timesheet ------------------------------------------------------
wb, ws = new_sheet("California daily overtime timesheet", "Company: ____________   Location: ____________",
                   sheet_name="CA overtime timesheet")
label(ws, 3, 1, "Employee", "", vcol=2)
label(ws, 4, 1, "Week starting", WEEK, vcol=2, fmt=DATE)
label(ws, 3, 6, "Regular hourly rate", None, vcol=7, fmt=MONEY)
hr = 6
heads = ["Day", "Date", "Time in", "Time out", "Unpaid meal breaks (minutes)", "7th consecutive day? (Y)",
         "Hours worked", "Regular", "Overtime 1.5x", "Double time 2x"]
for c, t in enumerate(heads, 1):
    hdr(ws, hr, c, t)
for k in range(7):
    r = hr + 1 + k
    z = k % 2
    put(ws, r, 1, f'=TEXT(B{r},"dddd")', zebra=z, bold=True)
    put(ws, r, 2, f"=$B$4+{k}", fmt="m/d", zebra=z, center=True)
    put(ws, r, 3, None, fmt=TIME, zebra=z, center=True)
    put(ws, r, 4, None, fmt=TIME, zebra=z, center=True)
    put(ws, r, 5, None, fmt="0", zebra=z, center=True)
    put(ws, r, 6, "", zebra=z, center=True)
    put(ws, r, 7, "=" + hours_in_out(f"C{r}", f"D{r}", f"E{r}"), fmt="0.00", fill=CALC, center=True)
    put(ws, r, 8, f'=IF(UPPER(F{r})="Y",0,MIN(8,G{r}))', fmt="0.00", fill=CALC, center=True)
    put(ws, r, 9, f'=IF(UPPER(F{r})="Y",MIN(8,G{r}),MIN(4,MAX(0,G{r}-8)))', fmt="0.00", fill=CALC, center=True)
    put(ws, r, 10, f'=IF(UPPER(F{r})="Y",MAX(0,G{r}-8),MAX(0,G{r}-12))', fmt="0.00", fill=CALC, center=True)
choices(ws, "F7:F13", ["Y"])
put(ws, 14, 1, "Daily totals", bold=True, fill=GREY)
for c in range(2, 7):
    put(ws, 14, c, "", fill=GREY)
for c in range(7, 11):
    put(ws, 14, c, f"=SUM({L(c)}7:{L(c)}13)", fmt="0.00", fill=CALC, bold=True, center=True)
put(ws, 15, 1, "Weekly overtime: regular hours over 40 move to overtime", bold=True, fill=GREY)
for c in range(2, 8):
    put(ws, 15, c, "", fill=GREY)
put(ws, 15, 8, "=-MAX(0,H14-40)", fmt="0.00", fill=CALC, center=True)
put(ws, 15, 9, "=MAX(0,H14-40)", fmt="0.00", fill=CALC, center=True)
put(ws, 15, 10, 0, fmt="0.00", fill=CALC, center=True)
put(ws, 16, 1, "Hours to pay", bold=True, fill=GREY)
for c in range(2, 7):
    put(ws, 16, c, "", fill=GREY)
put(ws, 16, 7, "=H16+I16+J16", fmt="0.00", fill=CALC, bold=True, center=True)
for c in (8, 9, 10):
    put(ws, 16, c, f"={L(c)}14+{L(c)}15", fmt="0.00", fill=CALC, bold=True, center=True)
calc_label(ws, 18, 7, "Regular pay", "=ROUND(H16*G3,2)", MONEY, vcol=10)
calc_label(ws, 19, 7, "Overtime pay (1.5x)", "=ROUND(I16*G3*1.5,2)", MONEY, vcol=10)
calc_label(ws, 20, 7, "Double-time pay (2x)", "=ROUND(J16*G3*2,2)", MONEY, vcol=10)
calc_label(ws, 21, 7, "Gross pay", "=J18+J19+J20", MONEY, vcol=10)
note(ws, 23, "California: overtime after 8 hours in a workday or 40 in a workweek, double time after 12 in a workday. "
             "On the 7th consecutive workday in the workweek, the first 8 hours are overtime and the rest double time (mark it Y).")
note(ws, 24, "Daily overtime hours are not counted again toward the 40. Alternative workweek schedules and exempt employees follow different rules.")
note(ws, 25, CERT)
end = sign_lines(ws, 27, ["Employee signature", "Supervisor approval"], gap=3)
finish(ws, [13, 8, 11, 11, 12, 12, 11, 10, 11, 11], hr, end + 1)
wb.save(d / "california-daily-overtime-timesheet.xlsx")
sheet_pdf(d / "california-daily-overtime-timesheet.pdf", "California daily overtime timesheet",
          ["Employee", "Week starting", "Regular hourly rate", "Company", "Location", "Supervisor"],
          ["Day", "Date", "Time in", "Time out", "Unpaid meal breaks (min)", "7th consecutive day?", "Hours worked",
           "Regular", "Overtime 1.5x", "Double time 2x"],
          [[x] + [""] * 9 for x in DAYS7] + [["Totals"] + [""] * 9],
          summary=[("Regular hours over 40 in the week (move to overtime)", ""), ("Regular hours to pay", ""),
                   ("Overtime hours to pay (1.5x)", ""), ("Double-time hours to pay (2x)", "")],
          notes="California: overtime after 8 hours in a workday or 40 in a workweek; double time after 12 hours in a "
                "workday; on the 7th consecutive workday, the first 8 hours are overtime and the rest double time.",
          sign=["Employee signature", "Supervisor approval"])

# 4. Semimonthly timesheet ------------------------------------------------------------------
wb, ws = new_sheet("Semimonthly timesheet", "Company: ____________   Department: ____________   Supervisor: ____________",
                   sheet_name="Semimonthly timesheet")
label(ws, 3, 1, "Employee", "", vcol=2)
label(ws, 4, 1, "Period starts (1st or 16th)", PERIOD, vcol=2, fmt=DATE)
label(ws, 3, 5, "Hourly rate", None, vcol=6, fmt=MONEY)
label(ws, 4, 5, "Workweek starts (1=Sun ... 7=Sat)", 1, vcol=6)
hr = 6
for c, t in enumerate(["Date", "Day", "Time in", "Time out", "Unpaid break (minutes)", "Hours worked",
                       "Workweek starting", "Notes"], 1):
    hdr(ws, hr, c, t)
for k in range(16):
    r = hr + 1 + k
    z = k % 2
    put(ws, r, 1, f'=IF($B$4+{k}>IF(DAY($B$4)=1,$B$4+14,EOMONTH($B$4,0)),"",$B$4+{k})', fmt="m/d", zebra=z, center=True)
    put(ws, r, 2, f'=IF(A{r}="","",TEXT(A{r},"ddd"))', zebra=z, center=True, bold=True)
    put(ws, r, 3, None, fmt=TIME, zebra=z, center=True)
    put(ws, r, 4, None, fmt=TIME, zebra=z, center=True)
    put(ws, r, 5, None, fmt="0", zebra=z, center=True)
    put(ws, r, 6, f'=IF(A{r}="","",' + hours_in_out(f"C{r}", f"D{r}", f"E{r}") + ")", fmt="0.00", fill=CALC, center=True)
    put(ws, r, 7, f'=IF(A{r}="","",A{r}-MOD(WEEKDAY(A{r},1)-$F$4,7))', fmt="m/d", fill=CALC, center=True)
    put(ws, r, 8, "", zebra=z)
put(ws, 23, 1, "Total hours", bold=True, fill=GREY)
for c in range(2, 6):
    put(ws, 23, c, "", fill=GREY)
put(ws, 23, 6, "=SUM(F7:F22)", fmt="0.00", fill=CALC, bold=True, center=True)
put(ws, 23, 7, "", fill=GREY)
put(ws, 23, 8, "", fill=GREY)
for c, t in [(1, "Workweek starting"), (2, "Hours in this period"), (3, "Overtime hours (over 40)")]:
    hdr(ws, 25, c, t)
for i in range(4):
    r = 26 + i
    put(ws, r, 1, "=G7" if i == 0 else f'=IF(A{r - 1}="","",IF(A{r - 1}+7>MAX($A$7:$A$22),"",A{r - 1}+7))',
        fmt="m/d/yyyy", fill=CALC, center=True)
    put(ws, r, 2, f'=IF(A{r}="","",SUMIF($G$7:$G$22,A{r},$F$7:$F$22))', fmt="0.00", fill=CALC, center=True)
    put(ws, r, 3, f'=IF(A{r}="","",MAX(0,B{r}-40))', fmt="0.00", fill=CALC, center=True)
calc_label(ws, 31, 4, "Regular hours", "=F23-SUM(C26:C29)", "0.00", vcol=6)
calc_label(ws, 32, 4, "Overtime hours (1.5x)", "=SUM(C26:C29)", "0.00", vcol=6)
calc_label(ws, 33, 4, "Gross pay", "=ROUND(F31*F3+F32*F3*1.5,2)", MONEY, vcol=6)
note(ws, 35, "Semimonthly periods do not line up with workweeks. A workweek that starts in the previous period or ends "
             "in the next one needs those hours too before you decide on overtime.")
note(ws, 36, CERT)
end = sign_lines(ws, 38, ["Employee signature", "Supervisor approval"], gap=3)
finish(ws, [16, 9, 11, 11, 12, 11, 12, 24], hr, end + 1, landscape=False)
wb.save(d / "semimonthly-timesheet.xlsx")
sheet_pdf(d / "semimonthly-timesheet.pdf", "Semimonthly timesheet",
          ["Employee", "Pay period", "Hourly rate", "Company", "Department", "Supervisor"],
          ["Date", "Day", "Time in", "Time out", "Unpaid break (min)", "Hours worked", "Notes"],
          [[""] * 7 for _ in range(16)] + [["Total hours", "", "", "", "", "", ""]],
          summary=[("Hours over 40 in any workweek (overtime)", ""), ("Regular hours", ""), ("Gross pay", "")],
          notes="Count overtime by workweek, including days from the previous or next pay period. " + CERT,
          sign=["Employee signature", "Supervisor approval"], landscape=False)

# 5. Monthly timesheet with leave columns -----------------------------------------------------
wb, ws = new_sheet("Monthly timesheet", "Company: ____________   Department: ____________",
                   sheet_name="Monthly timesheet")
label(ws, 3, 1, "Employee", "", vcol=2)
label(ws, 4, 1, "Month (first day)", PERIOD, vcol=2, fmt="mmmm yyyy")
hr = 6
heads = ["Date", "Day", "Hours worked", "PTO", "Sick", "Holiday", "Unpaid leave", "Total paid hours", "Notes"]
for c, t in enumerate(heads, 1):
    hdr(ws, hr, c, t)
for k in range(31):
    r = hr + 1 + k
    z = k % 2
    put(ws, r, 1, f'=IF($B$4+{k}>EOMONTH($B$4,0),"",$B$4+{k})', fmt="m/d", zebra=z, center=True)
    put(ws, r, 2, f'=IF(A{r}="","",TEXT(A{r},"ddd"))', zebra=z, center=True, bold=True)
    for c in range(3, 8):
        put(ws, r, c, None, fmt="0.00", zebra=z, center=True)
    put(ws, r, 8, f'=IF(A{r}="","",SUM(C{r}:F{r}))', fmt="0.00", fill=CALC, center=True)
    put(ws, r, 9, "", zebra=z)
rl = hr + 31
ws.conditional_formatting.add(f"A{hr + 1}:G{rl}", FormulaRule(formula=[f'AND($A{hr + 1}<>"",WEEKDAY($A{hr + 1},2)>5)'],
                                                                 fill=PatternFill("solid", fgColor="E4E4DD")))
put(ws, rl + 1, 1, "Monthly totals", bold=True, fill=GREY)
put(ws, rl + 1, 2, "", fill=GREY)
for c in range(3, 9):
    put(ws, rl + 1, c, f"=SUM({L(c)}{hr + 1}:{L(c)}{rl})", fmt="0.00", fill=CALC, bold=True, center=True)
put(ws, rl + 1, 9, "", fill=GREY)
note(ws, rl + 3, "Enter decimal hours (7.5 = 7 hours 30 minutes). Weekends are shaded. Unpaid leave is tracked but not "
                 "added to paid hours. Non-exempt employees: overtime is still owed for each workweek over 40 hours worked.")
note(ws, rl + 4, CERT)
end = sign_lines(ws, rl + 6, ["Employee signature", "Supervisor approval"], gap=3)
finish(ws, [16, 8, 11, 9, 9, 9, 10, 12, 26], hr, end + 1, landscape=False)
wb.save(d / "monthly-timesheet.xlsx")
sheet_pdf(d / "monthly-timesheet.pdf", "Monthly timesheet",
          ["Employee", "Month", "Department"],
          ["Date", "Day", "Hours worked", "PTO", "Sick", "Holiday", "Unpaid", "Total paid", "Notes"],
          [[str(i), "", "", "", "", "", "", "", ""] for i in range(1, 32)] + [["Totals"] + [""] * 8],
          notes=CERT, sign=["Employee signature", "Supervisor approval"], landscape=False, compact=True)

# 6. Team weekly timesheet --------------------------------------------------------------------
wb, ws = new_sheet("Team weekly timesheet", "Business: ____________   Approved by: ____________",
                   sheet_name="Team timesheet")
label(ws, 3, 1, "Location", "", vcol=2)
label(ws, 4, 1, "Week starting", WEEK, vcol=2, fmt=DATE)
hr = 6
hdr(ws, hr, 1, "Employee")
hdr(ws, hr, 2, "Role")
for k in range(7):
    hdr(ws, hr, 3 + k, f'=TEXT($B$4+{k},"ddd m/d")')
for c, t in [(10, "Total hours"), (11, "Regular"), (12, "Overtime"), (13, "Hourly rate"), (14, "Gross pay")]:
    hdr(ws, hr, c, t)
r0, n = hr + 1, 20
for i in range(n):
    r = r0 + i
    z = i % 2
    put(ws, r, 1, "", zebra=z)
    put(ws, r, 2, "", zebra=z)
    for k in range(7):
        put(ws, r, 3 + k, None, fmt="0.00", zebra=z, center=True)
    put(ws, r, 10, f"=SUM(C{r}:I{r})", fmt="0.00", fill=CALC, center=True)
    put(ws, r, 11, f"=MIN(40,J{r})", fmt="0.00", fill=CALC, center=True)
    put(ws, r, 12, f"=MAX(0,J{r}-40)", fmt="0.00", fill=CALC, center=True)
    put(ws, r, 13, None, fmt=MONEY, zebra=z, center=True)
    put(ws, r, 14, f"=ROUND(K{r}*M{r}+L{r}*M{r}*1.5,2)", fmt=MONEY, fill=CALC, center=True)
rl = r0 + n - 1
put(ws, rl + 1, 1, "Totals", bold=True, fill=GREY)
put(ws, rl + 1, 2, "", fill=GREY)
for c in list(range(3, 13)) + [14]:
    put(ws, rl + 1, c, f"=SUM({L(c)}{r0}:{L(c)}{rl})", fmt=MONEY if c == 14 else "0.00", fill=CALC, bold=True, center=True)
put(ws, rl + 1, 13, "", fill=GREY)
note(ws, rl + 3, "Enter each person's daily hours worked in decimals (8.5 = 8 hours 30 minutes), after unpaid meal breaks. "
                 "Keep each employee's own time records (in and out times) as the source for these totals.")
end = sign_lines(ws, rl + 5, ["Manager approval"], gap=3)
finish(ws, [22, 14] + [9] * 7 + [10, 9, 9, 10, 12], hr, end + 1, freeze="C7")
wb.save(d / "team-weekly-timesheet.xlsx")
sheet_pdf(d / "team-weekly-timesheet.pdf", "Team weekly timesheet",
          ["Location", "Week starting", "Approved by"],
          ["Employee", "Role", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun", "Total", "Overtime", "Rate", "Gross pay"],
          [[""] * 13 for _ in range(15)] + [["Totals"] + [""] * 12],
          notes="Daily hours worked in decimals, after unpaid meal breaks. Overtime = hours over 40 in the week.",
          sign=["Manager approval"])

# 7. Project and job timesheet ----------------------------------------------------------------
wb, ws = new_sheet("Project and job timesheet", "Company: ____________   Supervisor: ____________",
                   sheet_name="Job timesheet")
label(ws, 3, 1, "Employee", "", vcol=2)
label(ws, 4, 1, "Week starting", WEEK, vcol=2, fmt=DATE)
hr = 6
heads = ["Date", "Client or job", "Job number", "Task", "Start", "End", "Break (minutes)", "Hours",
         "Billable? (Y/N)", "Bill rate", "Billable amount"]
for c, t in enumerate(heads, 1):
    hdr(ws, hr, c, t)
r0, n = hr + 1, 30
for i in range(n):
    r = r0 + i
    z = i % 2
    put(ws, r, 1, None, fmt="m/d", zebra=z, center=True)
    for c in (2, 3, 4):
        put(ws, r, c, "", zebra=z)
    put(ws, r, 5, None, fmt=TIME, zebra=z, center=True)
    put(ws, r, 6, None, fmt=TIME, zebra=z, center=True)
    put(ws, r, 7, None, fmt="0", zebra=z, center=True)
    put(ws, r, 8, "=" + hours_in_out(f"E{r}", f"F{r}", f"G{r}"), fmt="0.00", fill=CALC, center=True)
    put(ws, r, 9, "", zebra=z, center=True)
    put(ws, r, 10, None, fmt=MONEY, zebra=z, center=True)
    put(ws, r, 11, f'=IF(UPPER(I{r})="Y",ROUND(H{r}*J{r},2),0)', fmt=MONEY, fill=CALC, center=True)
rl = r0 + n - 1
choices(ws, f"I{r0}:I{rl}", ["Y", "N"])
put(ws, rl + 1, 1, "Totals", bold=True, fill=GREY)
for c in range(2, 8):
    put(ws, rl + 1, c, "", fill=GREY)
put(ws, rl + 1, 8, f"=SUM(H{r0}:H{rl})", fmt="0.00", fill=CALC, bold=True, center=True)
put(ws, rl + 1, 9, "", fill=GREY)
put(ws, rl + 1, 10, "", fill=GREY)
put(ws, rl + 1, 11, f"=SUM(K{r0}:K{rl})", fmt=MONEY, fill=CALC, bold=True, center=True)
s = rl + 3
calc_label(ws, s, 9, "Total hours worked", f"=H{rl + 1}", "0.00", vcol=11)
calc_label(ws, s + 1, 9, "Billable hours", f'=SUMIF(I{r0}:I{rl},"Y",H{r0}:H{rl})', "0.00", vcol=11)
calc_label(ws, s + 2, 9, "Non-billable hours", f"=K{s}-K{s + 1}", "0.00", vcol=11)
calc_label(ws, s + 3, 9, "Billable share", f"=IF(K{s}=0,0,K{s + 1}/K{s})", "0%", vcol=11)
note(ws, s + 5, "One row per job or task block; a day can have several rows. Pay hourly employees for all hours worked, "
                "billable or not, and count all of them toward weekly overtime.")
note(ws, s + 6, CERT)
end = sign_lines(ws, s + 8, ["Employee signature", "Supervisor approval"], gap=3)
finish(ws, [9, 22, 11, 22, 10, 10, 9, 9, 10, 10, 13], hr, end + 1)
wb.save(d / "project-job-timesheet.xlsx")
sheet_pdf(d / "project-job-timesheet.pdf", "Project and job timesheet",
          ["Employee", "Week starting", "Supervisor"],
          ["Date", "Client or job", "Job no.", "Task", "Start", "End", "Break (min)", "Hours", "Billable?", "Notes"],
          [[""] * 10 for _ in range(18)] + [["Totals"] + [""] * 9],
          notes="Pay hourly employees for all hours worked, billable or not. " + CERT,
          sign=["Employee signature", "Supervisor approval"])

# 8. Printable weekly time card (Word + PDF) -------------------------------------------------
card = [
    ("fields", ["Employee name", "Employee ID", "Department or job site", "Week of"]),
    ("table", ["Day", "Date", "Time in", "Lunch out", "Lunch in", "Time out", "Total hours"],
     [[x, "", "", "", "", "", ""] for x in DAYS7] + [["Week total", "", "", "", "", "", ""]]),
    ("fields", ["Regular hours", "Overtime hours (over 40)", "Paid leave hours and type"]),
    ("p", CERT),
    ("sign", ["Employee signature"]),
    ("sign", ["Supervisor signature"]),
    ("p", "Payroll use only: entered by ________   date ________"),
]
docx_doc(d / "printable-weekly-time-card.docx", "Weekly time card", card)
pdf_doc(d / "printable-weekly-time-card.pdf", "Weekly time card", card)

report(d)
