#!/usr/bin/env python3
"""Build the downloads for /templates/shift-schedule-templates/.

Run from the site root:  python content-plan/template_builds/shift-schedule-templates.py
Files go to assets/templates/shift-schedule-templates/. Pattern sheets fill themselves from a start date
and a pattern string, so they are ready to use; staffing sheets are blank.
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


# ============================================================== /templates/shift-schedule-templates/
from openpyxl.formatting.rule import CellIsRule

d = out("shift-schedule-templates")
START = datetime.date(2027, 1, 4)  # a Monday; change it in the sheet
SHADES = {"D": "FFF1B8", "E": "FFD9B3", "N": "C9D7F2"}
DOW = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]


def shade_codes(ws, rng):
    for code, color in SHADES.items():
        ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=[f'"{code}"'],
                                                      fill=PatternFill("solid", fgColor=color)))


def pattern_sheet(path, title, subtitle, crews, days, shift_hours, codes, notes, pdf_title=None, blank_crews=0):
    """crews: [(crew, pattern, offset)]. Each day cell = MID(pattern, MOD(date - start + offset, LEN) + 1, 1)."""
    wb, ws = new_sheet(title, subtitle, sheet_name="Schedule")
    label(ws, 3, 1, "Start date (day 1 of your workweek)", START, vcol=2, fmt="ddd, mmm d, yyyy")
    label(ws, 4, 1, "Hours per shift", shift_hours, vcol=2, fmt="0.0")
    weeks = days // 7
    first = 4                         # column D
    last = first + days - 1
    hr = 6
    hdr(ws, hr, 1, "Crew")
    hdr(ws, hr, 2, "Pattern (D, E, N; - = off)")
    hdr(ws, hr, 3, "Offset (days)")
    hdr(ws, hr + 1, 1, "")
    hdr(ws, hr + 1, 2, "")
    hdr(ws, hr + 1, 3, "")
    for k in range(days):
        c = first + k
        hdr(ws, hr, c, f'=TEXT({L(c)}{hr + 1},"ddd")')
        cell = hdr(ws, hr + 1, c, f"=$B$3+{k}")
        cell.number_format = 'm"/"d'
    wcol = last + 1
    for w in range(weeks):
        hdr(ws, hr, wcol + w, f"Week {w + 1} hours")
        hdr(ws, hr + 1, wcol + w, "")
    otc = wcol + weeks
    hdr(ws, hr, otc, "FLSA overtime hours")
    hdr(ws, hr + 1, otc, "(over 40 each week)")
    rows = list(crews) + [("", "", 0)] * blank_crews
    r0 = hr + 2
    for i, (crew, pat, off) in enumerate(rows):
        r = r0 + i
        put(ws, r, 1, crew, bold=True)
        put(ws, r, 2, pat).font = Font(name="Consolas", size=9)
        put(ws, r, 3, off, center=True)
        for k in range(days):
            c = first + k
            put(ws, r, c, f'=IF($B{r}="","",MID($B{r},MOD({L(c)}${hr + 1}-$B$3+$C{r},LEN($B{r}))+1,1))',
                center=True)
        for w in range(weeks):
            a, b = L(first + 7 * w), L(first + 7 * w + 6)
            blk = f"{a}{r}:{b}{r}"
            f = "+".join(f'COUNTIF({blk},"{x}")' for x in codes)
            put(ws, r, wcol + w, f"=({f})*$B$4", fill=CALC, center=True, fmt="General")
        ot = "+".join(f"MAX(0,{L(wcol + w)}{r}-40)" for w in range(weeks))
        put(ws, r, otc, f"={ot}", fill=CALC, center=True, fmt="General")
    rl = r0 + len(rows) - 1
    shade_codes(ws, f"{L(first)}{r0}:{L(last)}{rl}")
    names = {"D": "Day shifts working", "E": "Evening shifts working", "N": "Night shifts working"}
    r = rl + 1
    for x in codes:
        put(ws, r, 1, names[x], bold=True, fill=GREY)
        put(ws, r, 2, "", fill=GREY)
        put(ws, r, 3, "", fill=GREY)
        for k in range(days):
            c = first + k
            put(ws, r, c, f'=COUNTIF({L(c)}{r0}:{L(c)}{rl},"{x}")', fill=CALC, center=True)
        r += 1
    note(ws, r + 1, notes)
    finish(ws, [30, 26, 9] + [5.2] * days + [8] * weeks + [11], hr + 1, r + 3,
           freeze=ws.cell(row=hr + 2, column=first))
    wb.save(path)

    # PDF: the pattern itself, one table per 28 days, Monday start
    def seq(p, o):
        return [p[(i + o) % len(p)] if p else "" for i in range(days)]
    body = ""
    for part in range(0, days, 28):
        span = range(part, min(days, part + 28))
        heads = ["Crew"] + [f"{DOW[i % 7]} {i + 1}" for i in span] + [f"Hours ({len(span)} days)"]
        prow = []
        for crew, p, o in rows:
            s = seq(p, o)
            hrs = sum(shift_hours for i in span if s[i] in codes)
            prow.append([crew] + [("" if s[i] == "-" else s[i]) for i in span] + [hrs if crew else ""])
        body += f"<h2>Days {span[0] + 1} to {span[-1] + 1}</h2>" + grid_html(heads, prow, "pat")
        body += "<p class='sub' style='margin-top:6px'>Dates: " + " ".join(["_____"] * 7) + "</p>"
    body += f"<p class='sub' style='margin-top:12px'>{_esc(notes)}</p>"
    html_pdf(path.with_suffix(".pdf"), pdf_title or title, body, subtitle=subtitle, landscape=True,
             extra_css=".pat th,.pat td{padding:4px 2px;text-align:center;font-size:8pt}.pat td:first-child,.pat th:first-child{text-align:left;padding-left:5px;width:70px}")


LEG = "D = day shift, N = night shift, - = off. "

# 1. Pitman / 2-2-3 ---------------------------------------------------------------------------
P = "DD--DDD--DD---"
PN = P.replace("D", "N")
pattern_sheet(d / "pitman-2-2-3-12-hour-shift-schedule.xlsx", "Pitman (2-2-3) 12-hour shift schedule",
              "Four crews, 14-day cycle: 2 on, 2 off, 3 on, 2 off, 2 on, 3 off. Crews A and B work days, C and D work nights.",
              [("Crew A (days, 7 a.m. to 7 p.m.)", P, 2), ("Crew B (days, 7 a.m. to 7 p.m.)", P, 9),
               ("Crew C (nights, 7 p.m. to 7 a.m.)", PN, 2), ("Crew D (nights, 7 p.m. to 7 a.m.)", PN, 9)],
              28, 12, ["D", "N"],
              LEG + "Each crew works 84 hours every 14 days (42 a week on average). With these offsets each week is 36 or 48 hours; "
                    "change the offset and the overtime column shows the effect. Night shifts are shown on the day they start.")

# 2. DuPont -----------------------------------------------------------------------------------
DP = "NNNN---DDD-NNN---DDDD-------"
pattern_sheet(d / "dupont-12-hour-shift-schedule.xlsx", "DuPont 12-hour shift schedule",
              "Four crews, 28-day cycle: 4 nights, 3 off, 3 days, 1 off, 3 nights, 3 off, 4 days, 7 off.",
              [("Crew A", DP, 0), ("Crew B", DP, 7), ("Crew C", DP, 14), ("Crew D", DP, 21)],
              28, 12, ["D", "N"],
              LEG + "Each crew works 168 hours every 28 days (42 a week on average), in weeks of 48, 72, 48 and 0 hours. "
                    "Night shifts are shown on the day they start.")

# 3. 4 on, 4 off --------------------------------------------------------------------------------
pattern_sheet(d / "4-on-4-off-12-hour-shift-schedule.xlsx", "4-on, 4-off 12-hour shift schedule",
              "Four crews, 8-day cycle: 4 shifts on, 4 days off. Crews A and B work days, C and D work nights. Eight weeks shown.",
              [("Crew A (days)", "DDDD----", 0), ("Crew B (days)", "DDDD----", 4),
               ("Crew C (nights)", "NNNN----", 0), ("Crew D (nights)", "NNNN----", 4)],
              56, 12, ["D", "N"],
              LEG + "The 8-day cycle does not line up with a 7-day week, so a crew works 48 hours some weeks and 36 in others "
                    "(42 a week on average). Night shifts are shown on the day they start.")

# 4. 7 on, 7 off --------------------------------------------------------------------------------
S7 = "DDDDDDD-------"
pattern_sheet(d / "7-on-7-off-12-hour-shift-schedule.xlsx", "7-on, 7-off 12-hour shift schedule",
              "Four crews, 14-day cycle: 7 shifts on (Friday to Thursday), 7 days off. Crews A and B work days, C and D work nights.",
              [("Crew A (days)", S7, 3), ("Crew B (days)", S7, 10),
               ("Crew C (nights)", S7.replace("D", "N"), 3), ("Crew D (nights)", S7.replace("D", "N"), 10)],
              28, 12, ["D", "N"],
              LEG + "Crews change over on Friday, so with a Monday workweek each week is 48 or 36 hours (8 overtime hours every "
                    "two weeks). If the workweek starts on the changeover day, a week on is 84 hours and 44 of them are overtime.")

# 5. Custom pattern builder ---------------------------------------------------------------------
pattern_sheet(d / "custom-shift-pattern-builder.xlsx", "Shift pattern builder",
              "Type any repeating pattern for each crew using D (day), E (evening), N (night) and - (off). The calendar, coverage and hours fill in.",
              [("Crew A", "DDD---", 0), ("Crew B", "DDD---", 3), ("Crew C", "NNN---", 0), ("Crew D", "NNN---", 3)],
              56, 12, ["D", "E", "N"],
              "Example shown: 3 on, 3 off with 12-hour shifts. Replace the patterns, offsets and hours per shift with your own; "
              "add crews in the blank rows. The pattern can be any length, for example 14 or 28 letters.",
              pdf_title="Shift pattern builder (worksheet)", blank_crews=2)

# 6. Rotating 8-hour, 3-shift weekly rotation ------------------------------------------------
R = "DDDDD--EEEEE--NNNNN--"
pattern_sheet(d / "rotating-8-hour-3-shift-schedule.xlsx", "Rotating 8-hour 3-shift schedule",
              "Three crews rotate forward each week: days (7 a.m. to 3 p.m.), then evenings (3 p.m. to 11 p.m.), then nights (11 p.m. to 7 a.m.). Monday to Friday.",
              [("Crew A", R, 0), ("Crew B", R, 14), ("Crew C", R, 7)],
              21, 8, ["D", "E", "N"],
              "D = day, E = evening, N = night, - = off. Each crew works 40 hours a week with weekends off. "
              "Forward rotation (days, evenings, nights) is shown; set the offsets to change the order.")

# 7. Two-week 12-hour nurse schedule (3 x 12) ----------------------------------------------
wb, ws = new_sheet("12-hour nurse schedule (two weeks)",
                   "Unit: ____________   Codes: D = day 7a-7p, N = night 7p-7a, PTO, blank = off", sheet_name="Schedule")
label(ws, 3, 1, "Start date (day 1 of the workweek)", START, vcol=2, fmt="m/d/yyyy")
label(ws, 4, 1, "Hours per shift", 12, vcol=2, fmt="0.0")
hr = 6
for c, t in enumerate(["Name", "Role (RN, LPN, CNA)"], 1):
    hdr(ws, hr, c, t)
    hdr(ws, hr + 1, c, "")
for k in range(14):
    c = 3 + k
    hdr(ws, hr, c, f'=TEXT({L(c)}{hr + 1},"ddd")')
    hdr(ws, hr + 1, c, f"=$B$3+{k}").number_format = 'm"/"d'
extra = ["Week 1 shifts", "Week 2 shifts", "Week 1 hours", "Week 2 hours", "Overtime hours (FLSA, over 40 a week)", "14-day hours"]
for j, t in enumerate(extra):
    hdr(ws, hr, 17 + j, t)
    hdr(ws, hr + 1, 17 + j, "")
r0, n = hr + 2, 20
for i in range(n):
    r = r0 + i
    put(ws, r, 1, "")
    put(ws, r, 2, "", center=True)
    for k in range(14):
        put(ws, r, 3 + k, "", center=True, zebra=i % 2 == 1)
    w1, w2 = f"C{r}:I{r}", f"J{r}:P{r}"
    put(ws, r, 17, f'=COUNTIF({w1},"D")+COUNTIF({w1},"N")', fill=CALC, center=True)
    put(ws, r, 18, f'=COUNTIF({w2},"D")+COUNTIF({w2},"N")', fill=CALC, center=True)
    put(ws, r, 19, f"=Q{r}*$B$4", fill=CALC, center=True)
    put(ws, r, 20, f"=R{r}*$B$4", fill=CALC, center=True)
    put(ws, r, 21, f"=MAX(0,S{r}-40)+MAX(0,T{r}-40)", fill=CALC, center=True)
    put(ws, r, 22, f"=S{r}+T{r}", fill=CALC, center=True)
rl = r0 + n - 1
choices(ws, f"C{r0}:P{rl}", ["D", "N", "PTO"])
choices(ws, f"B{r0}:B{rl}", ["RN", "LPN", "CNA"])
shade_codes(ws, f"C{r0}:P{rl}")
r = rl + 1
for lab, role, code, need in [("RN days needed", "RN", "D", 4), ("RN days scheduled", "RN", "D", None),
                              ("RN nights needed", "RN", "N", 3), ("RN nights scheduled", "RN", "N", None),
                              ("CNA days needed", "CNA", "D", 2), ("CNA days scheduled", "CNA", "D", None),
                              ("CNA nights needed", "CNA", "N", 2), ("CNA nights scheduled", "CNA", "N", None)]:
    put(ws, r, 1, lab, bold=True, fill=GREY)
    put(ws, r, 2, "", fill=GREY)
    for k in range(14):
        c = 3 + k
        if need is None:
            put(ws, r, c, f'=COUNTIFS($B${r0}:$B${rl},"{role}",{L(c)}{r0}:{L(c)}{rl},"{code}")', fill=CALC, center=True)
        else:
            put(ws, r, c, need, fill=INPUT, center=True)
    if need is None:
        ws.conditional_formatting.add(f"C{r}:P{r}", CellIsRule(operator="lessThan", formula=[f"C{r - 1}"],
                                                              fill=PatternFill("solid", fgColor="F8C9C0")))
    r += 1
note(ws, r + 1, "Yellow rows: how many staff each shift needs. Scheduled counts turn red when a shift is short. "
                "Hospitals using the 8-and-80 system with a prior agreement compute overtime over 8 hours a day and 80 hours "
                "in the 14-day period instead of the weekly column.")
finish(ws, [30, 12] + [5.4] * 14 + [9, 9, 9, 9, 13, 9], hr + 1, r + 3, freeze="C8")
wb.save(d / "12-hour-nurse-schedule-two-weeks.xlsx")
heads = ["Name", "Role"] + [f"{DOW[k % 7]} {k + 1}" for k in range(14)] + ["Wk 1 hrs", "Wk 2 hrs"]
rows = [[""] * len(heads) for _ in range(14)]
rows += [["RN days needed / scheduled", ""] + [""] * 16, ["RN nights needed / scheduled", ""] + [""] * 16]
html_pdf(d / "12-hour-nurse-schedule-two-weeks.pdf", "12-hour nurse schedule (two weeks)",
         grid_html(heads, rows, "pat") + "<p class='sub' style='margin-top:10px'>Codes: D = day 7a-7p, N = night 7p-7a, PTO. "
         "Dates: ______ to ______. Unit: ______________</p>",
         landscape=True, extra_css=".pat th,.pat td{padding:6px 3px;font-size:8.5pt;text-align:center;height:18px}.pat td:first-child{width:150px}")

report(d)
