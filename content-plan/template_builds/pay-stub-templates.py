#!/usr/bin/env python3
"""Build the downloads for /templates/pay-stub-templates/.

Run from the site root:  python content-plan/template_builds/pay-stub-templates.py
Files go to assets/templates/pay-stub-templates/. Files are blank and ready to fill in, with working
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


# ============================================================== /templates/pay-stub-templates/
d = out("pay-stub-templates")
SS_BASE_2026 = 184500   # IRS Topic 751 / SSA, 2026 Social Security wage base
PDF_CSS = (".hd{display:grid;grid-template-columns:1fr 1fr;gap:6px 30px;margin:0 0 14px;font-size:9.5pt}"
           ".hd div{border-bottom:1px solid #333;padding:10px 0 2px;color:#333}"
           "table{margin-top:10px} td{height:16px} th:nth-child(n+2),td:nth-child(n+2){text-align:right}"
           ".tot td,.tot th{font-weight:700;background:#e9e9e3 !important;color:#111}"
           ".net{margin-top:12px;font-size:13pt;font-weight:700;border:2px solid #111;padding:8px 12px;display:inline-block}"
           ".note{font-size:8.5pt;color:#444;margin-top:12px} .foot{position:static;margin-top:18px}")


def build_stub(name, title, subtitle, info, earnings, deductions, extra=None, notes=(), settings=(), paid_gross=None,
               pdf_info=None, landscape=False):
    """One pay stub sheet.

    info: [(label, default)] pairs placed two per row in the header (yellow input cells).
    earnings / deductions: [(key, description, B, C, D)] where B = hours or basis, C = rate, D = current amount.
      Use None for an input cell, "" for an unused cell, or a formula string with {key} placeholders that are
      replaced by that line's row number (and {gross}, {fica}, {ded}, {set_<n>} for totals and settings).
    settings: [(key, label, value, fmt)] small table of tax settings at the bottom (yellow).
    paid_gross: formula for the gross actually paid in the check when it differs from taxable gross (tips).
    """
    wb, ws = new_sheet(title, subtitle, sheet_name="Pay stub")
    r = 3
    info_cells = {}
    for i in range(0, len(info), 2):
        for j, (lab, default) in enumerate(info[i:i + 2]):
            c = 1 if j == 0 else 4
            fmt = DATE if isinstance(default, datetime.date) else None
            cell = label(ws, r, c, lab, default, vcol=c + 1, fmt=fmt)
            if j == 0:
                ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
            else:
                ws.merge_cells(start_row=r, start_column=5, end_row=r, end_column=6)
            info_cells[lab] = cell.coordinate
        r += 1
    rows = {}
    r += 1
    eh = r
    for c, t in enumerate(["Earnings", "Hours or units", "Rate", "This period", "Prior year to date", "Year to date"], 1):
        hdr(ws, eh, c, t)
    r += 1
    for key, *_ in earnings:
        rows[key] = r
        r += 1
    rows["gross"] = r
    r += 1
    if paid_gross:
        rows["paid"] = r
        r += 1
    rows["fica"] = r
    r += 2
    dh = r
    for c, t in enumerate(["Deductions", "Wages or basis", "Rate", "This period", "Prior year to date", "Year to date"], 1):
        hdr(ws, dh, c, t)
    r += 1
    for key, *_ in deductions:
        rows[key] = r
        r += 1
    rows["ded"] = r
    rows["net"] = r + 2
    r = rows["net"] + 2
    sr = r + 1
    for i, (key, *_rest) in enumerate(settings):
        rows["set_" + key] = sr + i
    rows.update({k: v for k, v in info_cells.items()})

    def fx(v):
        if isinstance(v, str) and v.startswith("="):
            return v.format(**{k: v_ for k, v_ in rows.items() if isinstance(v_, int)})
        return v

    def line(rr, desc, b, c_, dd, z):
        put(ws, rr, 1, desc, bold=True, zebra=z)
        for col, v, fmt in ((2, b, "0.00"), (3, c_, None), (4, dd, MONEY)):
            if v is None:
                put(ws, rr, col, None, fmt=fmt or "0.00##", zebra=z, center=True).fill = INPUT
            elif isinstance(v, str) and v.startswith("="):
                put(ws, rr, col, fx(v), fmt=fmt or "0.00##", fill=CALC, center=True)
            elif isinstance(v, tuple):  # (value, fmt) input with a preset value
                put(ws, rr, col, v[0], fmt=v[1], zebra=z, center=True).fill = INPUT
            else:
                put(ws, rr, col, v, fmt=fmt, zebra=z, center=True)
        put(ws, rr, 5, None, fmt=MONEY, center=True).fill = INPUT
        put(ws, rr, 6, f"=N(D{rr})+N(E{rr})", fmt=MONEY, fill=CALC, center=True)

    for i, (key, desc, b, c_, dd) in enumerate(earnings):
        line(rows[key], desc, b, c_, dd if dd != "x" else f"=ROUND(N(B{rows[key]})*N(C{rows[key]}),2)", i % 2)
    e0, e1 = eh + 1, rows["gross"] - 1
    g = rows["gross"]
    put(ws, g, 1, "Gross pay", bold=True, fill=GREY)
    put(ws, g, 2, "", fill=GREY)
    put(ws, g, 3, "", fill=GREY)
    for col in (4, 5, 6):
        put(ws, g, col, f"=SUM({L(col)}{e0}:{L(col)}{e1})", fmt=MONEY, fill=CALC, bold=True, center=True)
    if paid_gross:
        p = rows["paid"]
        put(ws, p, 1, "Gross paid in this check", bold=True, fill=GREY)
        put(ws, p, 2, "", fill=GREY)
        put(ws, p, 3, "", fill=GREY)
        for col in (4, 5, 6):
            put(ws, p, col, fx(paid_gross).replace("#", L(col)), fmt=MONEY, fill=CALC, bold=True, center=True)
    f_ = rows["fica"]
    put(ws, f_, 1, "Wages subject to Social Security and Medicare", fill=GREY)
    put(ws, f_, 2, "", fill=GREY)
    put(ws, f_, 3, "", fill=GREY)
    exempt = [k for k, *_ in deductions if k.startswith("pre")]
    minus = "".join(f"-N(D{rows[k]})" for k in exempt)
    put(ws, f_, 4, f"=D{g}{minus}", fmt=MONEY, fill=CALC, center=True)
    put(ws, f_, 5, None, fmt=MONEY, center=True).fill = INPUT
    put(ws, f_, 6, f"=N(D{f_})+N(E{f_})", fmt=MONEY, fill=CALC, center=True)
    for i, (key, desc, b, c_, dd) in enumerate(deductions):
        line(rows[key], desc, b, c_, dd if dd != "x" else f"=ROUND(N(B{rows[key]})*N(C{rows[key]}),2)", i % 2)
    t = rows["ded"]
    put(ws, t, 1, "Total deductions", bold=True, fill=GREY)
    put(ws, t, 2, "", fill=GREY)
    put(ws, t, 3, "", fill=GREY)
    for col in (4, 5, 6):
        put(ws, t, col, f"=SUM({L(col)}{dh + 1}:{L(col)}{t - 1})", fmt=MONEY, fill=CALC, bold=True, center=True)
    n_ = rows["net"]
    base = rows["paid"] if paid_gross else g
    put(ws, n_, 1, "NET PAY", bold=True, fill=GREY)
    put(ws, n_, 2, "", fill=GREY)
    put(ws, n_, 3, "", fill=GREY)
    for col in (4, 5, 6):
        put(ws, n_, col, f"={L(col)}{base}-{L(col)}{t}", fmt=MONEY, fill=CALC, bold=True, center=True)
    ws.cell(row=n_, column=4).font = Font(bold=True, size=13)
    if settings:
        hdr(ws, sr - 1, 1, "Tax and policy settings")
        hdr(ws, sr - 1, 2, "Value")
        for i, (key, lab, val, fmt) in enumerate(settings):
            put(ws, sr + i, 1, lab, fill=GREY)
            put(ws, sr + i, 2, val, fmt=fmt, center=True).fill = INPUT
    r = sr + len(settings) + 1
    for k, (lab, formula, fmt) in enumerate(extra or []):
        calc_label(ws, r + k, 1, lab, fx(formula), fmt, vcol=4)
        ws.merge_cells(start_row=r + k, start_column=1, end_row=r + k, end_column=3)
    r += len(extra or []) + 1
    for k, text in enumerate(notes):
        note(ws, r + k, text)
    for col in (2, 3):
        for rr in list(range(e0, e1 + 1)) + list(range(dh + 1, t)):
            ws.cell(row=rr, column=col).number_format = ws.cell(row=rr, column=col).number_format
    finish(ws, [40, 14, 12, 14, 15, 15], eh, r + len(notes) + 1, landscape=landscape, freeze="A3")
    wb.save(d / f"{name}.xlsx")
    # printable PDF (blank)
    hd = "<div class='hd'>" + "".join(f"<div>{_esc(l)}:</div>" for l in (pdf_info or [x for x, _ in info])) + "</div>"
    et = grid_html(["Earnings", "Hours or units", "Rate", "This period", "Year to date"],
                   [[e[1], "", "", "", ""] for e in earnings] + [["Gross pay", "", "", "", ""]])
    dt_ = grid_html(["Deductions", "This period", "Year to date"],
                    [[x[1], "", ""] for x in deductions] + [["Total deductions", "", ""]])
    body = hd + et + dt_ + "<div class='net'>Net pay: ______________</div>"
    if notes:
        body += "<p class='note'>" + "<br>".join(_esc(n) for n in notes) + "</p>"
    html_pdf(d / f"{name}.pdf", title, body, landscape=False, extra_css=PDF_CSS)
    return rows


FED_NOTE = ("Federal and state income tax withholding are entered by you, from the employee's Form W-4 and IRS "
            "Publication 15-T (and your state's tables). Social Security (6.2%, 2026 wage base $184,500) and Medicare "
            "(1.45%) are calculated; Additional Medicare Tax (0.9% on wages over $200,000 in a year) is not.")
PRETAX_NOTE = ("Lines starting 'Pre-tax' are subtracted from Social Security and Medicare wages; use them only for "
               "deductions that are exempt from FICA, such as qualifying Section 125 health premiums. 401(k) deferrals are not.")
HONEST = "Use this template only to record wages actually paid. A pay stub showing income that was not paid is fraud."
SS = ("ss", "Social Security", "={fica}", "=C{set_ssrate}", "=ROUND(MIN(D{fica},MAX(0,B{set_ssbase}-E{fica}))*C{ss},2)")
MED = ("med", "Medicare", "={fica}", "=C{set_medrate}", "=ROUND(D{fica}*C{med},2)")
TAXSET = [("ssrate", "Social Security employee rate", 0.062, "0.00%"), ("ssbase", "Social Security wage base (2026)", SS_BASE_2026, '"$"#,##0'),
          ("medrate", "Medicare employee rate", 0.0145, "0.00%")]


def fix_refs(lines):
    """SS/Medicare reference rates in column B of the settings table; adjust the placeholders."""
    outl = []
    for key, desc, b, c_, dd in lines:
        b = b.replace("={fica}", "=D{fica}") if isinstance(b, str) else b
        c_ = c_.replace("=C{set_", "=B{set_") if isinstance(c_, str) else c_
        dd = dd.replace("B{set_ssbase}", "B{set_ssbase}") if isinstance(dd, str) else dd
        outl.append((key, desc, b, c_, dd))
    return outl


BASIC_INFO = [("Employer (legal name)", ""), ("Pay date", None), ("Employer address", ""), ("Pay period start", None),
              ("Employee", ""), ("Pay period end", None), ("Employee ID or last 4 of SSN", ""), ("Check or deposit no.", ""),
              ("Pay basis", "Hourly"), ("PTO or sick hours available", "")]

# 1. Hourly pay stub with overtime and YTD --------------------------------------------------
build_stub("hourly-pay-stub-with-overtime", "Pay stub (hourly)", "Earnings statement. Keep this for your records.",
           BASIC_INFO,
           [("reg", "Regular", None, None, "x"),
            ("ot", "Overtime (1.5x)", None, "=ROUND(C{reg}*1.5,4)", "x"),
            ("hol", "Holiday", None, "=C{reg}", "x"),
            ("pto", "PTO or sick leave paid", None, "=C{reg}", "x"),
            ("bonus", "Bonus or commission", "", "", None)],
           fix_refs([("pre_health", "Pre-tax health insurance (Section 125)", "", "", None),
                     ("k401", "401(k) deferral", "=D{gross}", (0.04, "0%"), "=ROUND(B{k401}*C{k401},2)"),
                     SS, MED,
                     ("fit", "Federal income tax", "", "", None),
                     ("sit", "State income tax", "", "", None),
                     ("other", "Other (garnishment, local tax, after-tax)", "", "", None)]),
           settings=TAXSET, notes=[FED_NOTE, PRETAX_NOTE, HONEST])

# 2. Salaried pay stub ----------------------------------------------------------------------
sal_info = [x if x[0] != "Pay basis" else ("Pay basis", "Salary, exempt") for x in BASIC_INFO]
build_stub("salaried-pay-stub", "Pay stub (salaried)", "Earnings statement. Keep this for your records.", sal_info,
           [("sal", "Salary for this pay period", "", "=ROUND(B{set_annual}/B{set_periods},2)", "=C{sal}"),
            ("ptomemo", "PTO hours taken (paid in salary)", None, "", ("", MONEY)),
            ("bonus", "Bonus or commission", "", "", None)],
           fix_refs([("pre_health", "Pre-tax health insurance (Section 125)", "", "", None),
                     ("pre_dental", "Pre-tax dental and vision (Section 125)", "", "", None),
                     ("k401", "401(k) deferral", "=D{gross}", (0.05, "0%"), "=ROUND(B{k401}*C{k401},2)"),
                     SS, MED,
                     ("fit", "Federal income tax", "", "", None),
                     ("sit", "State income tax", "", "", None),
                     ("other", "Other (after-tax)", "", "", None)]),
           settings=[("annual", "Annual salary", None, MONEY), ("periods", "Pay periods per year (52, 26, 24, 12)", 24, "0")] + TAXSET,
           notes=["Salary per period = annual salary / pay periods. PTO hours are shown for the employee's balance and do "
                  "not change exempt pay.", FED_NOTE, PRETAX_NOTE, HONEST])

# 3. Tipped employee pay stub ---------------------------------------------------------------
tip_info = [x if x[0] != "Pay basis" else ("Pay basis", "Hourly, tipped") for x in BASIC_INFO]
build_stub("tipped-employee-pay-stub", "Pay stub (tipped employee)", "Earnings statement. Keep this for your records.",
           tip_info,
           [("cash", "Cash wage", None, "=B{set_cash}", "x"),
            ("otcash", "Overtime cash wage (1.5 x minimum wage, less tip credit)", None,
             "=ROUND(B{set_min}*1.5-(B{set_min}-B{set_cash}),4)", "x"),
            ("ctips", "Charged tips paid through payroll", "", "", None),
            ("rtips", "Cash tips reported by employee (already received)", "", "", None),
            ("makeup", "Pay to reach minimum wage", "",
             "", "=MAX(0,ROUND((N(B{cash})+N(B{otcash}))*B{set_min}+N(B{otcash})*B{set_min}*0.5-(N(D{cash})+N(D{otcash})+N(D{ctips})+N(D{rtips})),2))")],
           fix_refs([SS, MED,
                     ("fit", "Federal income tax", "", "", None),
                     ("sit", "State income tax", "", "", None),
                     ("other", "Other (uniforms, after-tax)", "", "", None)]),
           settings=[("min", "Minimum wage", 7.25, MONEY), ("cash", "Cash wage paid by employer", 2.13, MONEY)] + TAXSET,
           paid_gross="=#{gross}-#{rtips}",
           extra=[("Tip credit taken this period", "=ROUND((N(B{cash})+N(B{otcash}))*(B{set_min}-B{set_cash}),2)", MONEY),
                  ("Cash wage + tips per hour worked", "=IF(N(B{cash})+N(B{otcash})=0,0,ROUND((D{cash}+D{otcash}+D{ctips}+D{rtips})/(B{cash}+B{otcash}),2))", MONEY)],
           notes=["Defaults are the federal figures: $7.25 minimum wage and a $2.13 cash wage (a $5.12 tip credit). Many "
                  "states require a higher cash wage or allow no tip credit; change the settings to your state's rates.",
                  "Reported cash tips count as wages for tax but are not paid again in the check, so net pay is figured "
                  "from 'gross paid in this check'. If cash wage plus tips falls short of minimum wage, the make-up line pays the difference.",
                  FED_NOTE, HONEST])

# 4. California pay stub (Labor Code 226) ---------------------------------------------------
ca_info = [("Employer legal name", ""), ("Pay date", None), ("Employer address", ""), ("Pay period start (inclusive)", None),
           ("Employee name", ""), ("Pay period end (inclusive)", None), ("Last 4 of SSN or employee ID", ""),
           ("Paid sick leave available (hours or days)", ""), ("Pay basis", "Hourly, non-exempt"), ("Check or deposit no.", "")]
build_stub("california-pay-stub", "California itemized wage statement", "Pay stub under California Labor Code section 226(a).",
           ca_info,
           [("reg", "Regular", None, None, "x"),
            ("ot", "Overtime (1.5x)", None, "=ROUND(C{reg}*1.5,4)", "x"),
            ("dt", "Double time (2x)", None, "=ROUND(C{reg}*2,4)", "x"),
            ("sick", "Paid sick leave", None, None, "x"),
            ("piece", "Piece-rate units (if paid by the piece)", None, None, "x")],
           fix_refs([SS, MED,
                     ("fit", "Federal income tax", "", "", None),
                     ("sit", "California income tax (PIT)", "", "", None),
                     ("sdi", "California SDI", "=D{gross}", "=B{set_sdi}", "=ROUND(B{sdi}*C{sdi},2)"),
                     ("other", "Deductions on the employee's written order (may be combined)", "", "", None)]),
           settings=[("sdi", "California SDI rate (2026)", 0.013, "0.0%")] + TAXSET,
           notes=["California requires, on each stub: gross wages, total hours worked (not for salaried exempt employees), "
                  "piece-rate units and rates if any, all deductions, net wages, the inclusive pay period dates, the "
                  "employee's name and only the last 4 digits of the SSN or an employee ID, the employer's legal name and "
                  "address, and every hourly rate with the hours worked at it. Show paid sick leave available here or on a "
                  "notice given with the pay.",
                  "The SDI rate is set each year by the EDD; 1.3% is the 2026 rate, with no wage cap.", FED_NOTE, HONEST])

# 5. New York pay stub (Labor Law 195.3) -----------------------------------------------------
ny_info = [("Employer name", ""), ("Pay date", None), ("Employer address", ""), ("Pay period start", None),
           ("Employer phone", ""), ("Pay period end", None), ("Employee name", ""), ("Basis of pay (hourly, salary, other)", "Hourly"),
           ("Check or deposit no.", ""), ("Vacation or sick hours available", "")]
build_stub("new-york-pay-stub", "New York wage statement", "Pay stub under New York Labor Law section 195(3).", ny_info,
           [("reg", "Regular", None, None, "x"),
            ("ot", "Overtime (1.5x)", None, "=ROUND(C{reg}*1.5,4)", "x"),
            ("allow", "Allowances or credits (meals, lodging, tips, uniforms)", "", "", None),
            ("other_e", "Other earnings (PTO, holiday, bonus)", None, None, "x")],
           fix_refs([SS, MED,
                     ("fit", "Federal income tax", "", "", None),
                     ("sit", "New York State income tax", "", "", None),
                     ("dbl", "NY disability (DBL)", "", "", None),
                     ("pfl", "NY Paid Family Leave", "=D{gross}", "=B{set_pflrate}",
                      "=ROUND(MIN(B{pfl}*C{pfl},MAX(0,B{set_pflmax}-E{pfl})),2)"),
                     ("other", "Other deductions (garnishment, after-tax)", "", "", None)]),
           settings=[("pflrate", "NY Paid Family Leave rate (2026)", 0.00432, "0.000%"),
                     ("pflmax", "NY Paid Family Leave annual maximum (2026)", 411.91, MONEY)] + TAXSET,
           notes=["New York requires, on each stub: the employer's name, address and phone number, the employee's name, "
                  "the dates covered, the basis of pay, regular and overtime rates, regular and overtime hours, any "
                  "allowances or credits claimed, gross wages, deductions and net wages.",
                  "PFL rate and cap are the 2026 figures; update them each January. DBL is usually set by your carrier.",
                  FED_NOTE, HONEST])

# 6. Pay stub with shift differential and multiple rates -------------------------------------
build_stub("pay-stub-with-shift-differential", "Pay stub (shift differential and bonus)",
           "Earnings statement. Overtime uses the regular rate, which includes differentials and nondiscretionary bonuses.",
           BASIC_INFO,
           [("reg", "Straight time, all hours worked", None, None, "x"),
            ("diff", "Shift differential (hours on the shift)", None, None, "x"),
            ("bonus", "Nondiscretionary bonus (attendance, production)", "", "", None),
            ("otp", "Overtime premium (0.5 x regular rate)", "=MAX(0,N(B{reg})-40)",
             "=IF(N(B{reg})=0,0,ROUND(0.5*(D{reg}+D{diff}+N(D{bonus}))/B{reg},4))", "x")],
           fix_refs([("pre_health", "Pre-tax health insurance (Section 125)", "", "", None), SS, MED,
                     ("fit", "Federal income tax", "", "", None),
                     ("sit", "State income tax", "", "", None),
                     ("other", "Other (local tax, after-tax)", "", "", None)]),
           settings=TAXSET,
           extra=[("Regular rate for overtime (straight-time pay / hours)",
                   "=IF(N(B{reg})=0,0,ROUND((D{reg}+D{diff}+N(D{bonus}))/B{reg},4))", '"$"#,##0.0000')],
           notes=["Enter ALL hours worked in the first line at the base rate; overtime hours (over 40 in the workweek) are "
                  "then paid an extra half of the regular rate. This weekly layout assumes a one-week pay period; for "
                  "longer periods, work out the premium for each workweek.", FED_NOTE, PRETAX_NOTE, HONEST])

report(d)
