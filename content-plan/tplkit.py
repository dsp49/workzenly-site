#!/usr/bin/env python3
"""Make the downloadable files for /templates/ pages: Excel (.xlsx), Word (.docx) and printable PDF.

Use from a small script (keep build scripts in content-plan/template_builds/<slug>.py so files can be rebuilt):

    import sys; sys.path.insert(0, "content-plan")
    from tplkit import out, xlsx_sheet, docx_doc, pdf_table, pdf_doc

    d = out("weekly-schedule-templates")            # -> assets/templates/weekly-schedule-templates/
    xlsx_sheet(d / "weekly-employee-schedule.xlsx",
               title="Weekly employee schedule", subtitle="Week of: ________   Location: ________",
               headers=["Employee", "Role", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun", "Total hours"],
               rows=[["Maya Torres", "Cashier", "9a–5p", "", "9a–5p", "9a–5p", "", "10a–6p", "", 32]],
               blank_rows=12, widths=[22, 14] + [11] * 7 + [12],
               formulas={"Total hours": None})       # or a template like "=SUM(C{r}:I{r})"
    pdf_table(d / "weekly-employee-schedule.pdf", "Weekly employee schedule", headers, rows, blank_rows=12,
              subtitle="Week of: ________", landscape=True)
    docx_doc(d / "time-off-request-form.docx", "Time off request form", [
        ("p", "Complete this form and give it to your manager at least two weeks before the first day off."),
        ("fields", ["Employee name", "Department", "Manager"]),
        ("table", ["Date", "Hours", "Type"], [["", "", ""]] * 5),
        ("checks", ["PTO", "Sick leave", "Unpaid"]),
        ("sign", ["Employee signature", "Manager approval"]),
    ])
    pdf_doc(d / "time-off-request-form.pdf", "Time off request form", same_blocks)

Block types for docx_doc / pdf_doc: ("h", text) heading, ("p", text) paragraph, ("bullets", [..]),
("fields", [labels]) label + blank line, ("table", headers, rows), ("checks", [options]) checkboxes,
("sign", [labels]) signature + date lines, ("letter", text) a full letter, keeps line breaks.

All files carry a small "Free template from workzenly.com" footer. Keep sample names realistic and US-based.
"""
import html as _html
import os
import pathlib
import shutil
import subprocess
import tempfile

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, RGBColor, Inches

ROOT = pathlib.Path(__file__).resolve().parent.parent
FOOT = "Free template from workzenly.com"
CHROME = [r"C:\Program Files\Google\Chrome\Application\chrome.exe",
          r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"]


def out(slug):
    d = ROOT / "assets" / "templates" / slug
    d.mkdir(parents=True, exist_ok=True)
    return d


# ---------------------------------------------------------------- Excel
def xlsx_sheet(path, title, headers, rows=(), subtitle=None, blank_rows=10, widths=None, formulas=None,
               sheet_name=None, notes=None, landscape=True, extra_sheets=None):
    """One formatted sheet. formulas: {header: "=SUM(C{r}:I{r})"} applied to every data row (sample + blank)."""
    wb = Workbook()
    _fill_sheet(wb.active, title, headers, rows, subtitle, blank_rows, widths, formulas, sheet_name, notes, landscape)
    for spec in extra_sheets or []:
        _fill_sheet(wb.create_sheet(), **spec)
    wb.save(path)
    return path


def _fill_sheet(ws, title, headers, rows=(), subtitle=None, blank_rows=10, widths=None, formulas=None,
                sheet_name=None, notes=None, landscape=True):
    ws.title = (sheet_name or title)[:31]
    thin = Side(style="thin", color="C9C9C3")
    box = Border(left=thin, right=thin, top=thin, bottom=thin)
    ncol = len(headers)
    ws.cell(row=1, column=1, value=title).font = Font(bold=True, size=16)
    r = 2
    if subtitle:
        ws.cell(row=2, column=1, value=subtitle).font = Font(size=11, color="444444")
        r = 3
    hr = r + 1
    for c, h in enumerate(headers, 1):
        cell = ws.cell(row=hr, column=c, value=h)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="161616")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = box
    data = list(rows) + [[""] * ncol for _ in range(blank_rows)]
    for i, row in enumerate(data):
        rr = hr + 1 + i
        for c, h in enumerate(headers, 1):
            v = row[c - 1] if c - 1 < len(row) else ""
            if formulas and h in formulas and formulas[h]:
                v = formulas[h].format(r=rr)
            cell = ws.cell(row=rr, column=c, value=v)
            cell.border = box
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            if i % 2 == 1:
                cell.fill = PatternFill("solid", fgColor="F6F6F3")
    last = hr + len(data)
    if notes:
        ws.cell(row=last + 2, column=1, value=notes).font = Font(italic=True, size=10, color="555555")
    ws.cell(row=last + (4 if notes else 2), column=1, value=FOOT).font = Font(size=9, color="888888")
    for c in range(1, ncol + 1):
        ws.column_dimensions[get_column_letter(c)].width = (widths[c - 1] if widths and c - 1 < len(widths) else 14)
    ws.row_dimensions[hr].height = 30
    ws.freeze_panes = ws.cell(row=hr + 1, column=1)
    ws.page_setup.orientation = "landscape" if landscape else "portrait"
    ws.page_setup.fitToWidth = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = f"{hr}:{hr}"


# ---------------------------------------------------------------- Word
def docx_doc(path, title, blocks):
    doc = Document()
    st = doc.styles["Normal"]
    st.font.name = "Calibri"
    st.font.size = Pt(11)
    for s in doc.sections:
        s.left_margin = s.right_margin = Inches(0.9)
        s.top_margin = s.bottom_margin = Inches(0.8)
        fp = s.footer.paragraphs[0]
        fp.text = FOOT
        fp.runs[0].font.size = Pt(8)
        fp.runs[0].font.color.rgb = RGBColor(0x88, 0x88, 0x88)
    h = doc.add_heading(title, level=1)
    for run in h.runs:
        run.font.color.rgb = RGBColor(0x11, 0x11, 0x11)
    for b in blocks:
        kind = b[0]
        if kind == "h":
            doc.add_heading(b[1], level=2)
        elif kind == "p":
            doc.add_paragraph(b[1])
        elif kind == "letter":
            for para in b[1].split("\n\n"):
                p = doc.add_paragraph()
                lines = para.split("\n")
                for i, ln in enumerate(lines):
                    run = p.add_run(ln)
                    if i < len(lines) - 1:
                        run.add_break()
        elif kind == "bullets":
            for item in b[1]:
                doc.add_paragraph(item, style="List Bullet")
        elif kind == "fields":
            for label in b[1]:
                p = doc.add_paragraph()
                p.add_run(f"{label}: ").bold = True
                p.add_run("_" * max(10, 60 - len(label)))
        elif kind == "checks":
            p = doc.add_paragraph()
            p.add_run("   ".join(f"☐ {o}" for o in b[1]))
        elif kind == "table":
            headers, rows = b[1], b[2]
            t = doc.add_table(rows=1, cols=len(headers))
            t.style = "Table Grid"
            for i, hh in enumerate(headers):
                t.rows[0].cells[i].text = hh
                for r in t.rows[0].cells[i].paragraphs[0].runs:
                    r.bold = True
            for row in rows:
                cells = t.add_row().cells
                for i, v in enumerate(row):
                    cells[i].text = str(v)
            doc.add_paragraph()
        elif kind == "sign":
            for label in b[1]:
                p = doc.add_paragraph()
                p.paragraph_format.space_before = Pt(18)
                p.add_run("_" * 34 + "        " + "_" * 16)
                q = doc.add_paragraph()
                q.add_run(f"{label}" + " " * max(1, 48 - len(label)) + "Date").font.size = Pt(9)
        else:
            raise ValueError(f"unknown block {kind}")
    doc.save(path)
    return path


# ---------------------------------------------------------------- PDF (rendered by Chrome)
_CSS = """
@page { size: %s; margin: 14mm; }
body { font-family: Arial, Helvetica, sans-serif; color: #111; font-size: 11pt; }
h1 { font-size: 20pt; margin: 0 0 4px; } h2 { font-size: 13pt; margin: 18px 0 6px; }
.sub { color: #444; margin: 0 0 12px; }
table { border-collapse: collapse; width: 100%%; font-size: 9.5pt; }
th, td { border: 1px solid #bbb; padding: 6px 7px; text-align: left; vertical-align: top; height: 16px; }
th { background: #161616; color: #fff; }
tr:nth-child(even) td { background: #f5f5f2; }
.field { margin: 10px 0; } .field b { display: inline-block; min-width: 160px; }
.line { display: inline-block; border-bottom: 1px solid #333; width: 60%%; height: 14px; }
.sign { margin-top: 34px; display: grid; grid-template-columns: 2.4fr 1fr; gap: 26px 36px; }
.sign div { border-top: 1px solid #333; padding-top: 4px; font-size: 9pt; color: #333; }
.letter { white-space: pre-wrap; font-family: Georgia, serif; font-size: 11.5pt; line-height: 1.55; }
.foot { position: fixed; bottom: 0; left: 0; right: 0; font-size: 8pt; color: #888; }
"""


def _pdf(path, body, landscape=False):
    chrome = next((c for c in CHROME if os.path.exists(c)), None)
    if not chrome:
        raise RuntimeError("Chrome or Edge is needed to render PDFs")
    tmp = tempfile.mkdtemp(prefix="wz-pdf-")
    try:
        src = pathlib.Path(tmp) / "doc.html"
        css = _CSS % ("letter landscape" if landscape else "letter portrait")
        src.write_text(f"<!doctype html><meta charset='utf-8'><style>{css}</style>{body}<div class='foot'>{FOOT}</div>",
                       encoding="utf-8")
        subprocess.run([chrome, "--headless=new", "--disable-gpu", "--no-first-run", f"--user-data-dir={tmp}/p",
                        "--no-pdf-header-footer", f"--print-to-pdf={pathlib.Path(path).resolve()}", src.as_uri()],
                       capture_output=True, timeout=120)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    if not pathlib.Path(path).exists():
        raise RuntimeError(f"PDF was not created: {path}")
    return path


def _esc(s):
    return _html.escape(str(s))


def pdf_table(path, title, headers, rows=(), blank_rows=10, subtitle=None, landscape=True, notes=None):
    data = list(rows) + [[""] * len(headers) for _ in range(blank_rows)]
    body = f"<h1>{_esc(title)}</h1>" + (f"<p class='sub'>{_esc(subtitle)}</p>" if subtitle else "")
    body += "<table><thead><tr>" + "".join(f"<th>{_esc(h)}</th>" for h in headers) + "</tr></thead><tbody>"
    body += "".join("<tr>" + "".join(f"<td>{_esc(v)}</td>" for v in r) + "</tr>" for r in data) + "</tbody></table>"
    if notes:
        body += f"<p class='sub' style='margin-top:12px'>{_esc(notes)}</p>"
    return _pdf(path, body, landscape)


def pdf_doc(path, title, blocks, landscape=False):
    body = f"<h1>{_esc(title)}</h1>"
    for b in blocks:
        k = b[0]
        if k == "h":
            body += f"<h2>{_esc(b[1])}</h2>"
        elif k == "p":
            body += f"<p>{_esc(b[1])}</p>"
        elif k == "letter":
            body += f"<div class='letter'>{_esc(b[1])}</div>"
        elif k == "bullets":
            body += "<ul>" + "".join(f"<li>{_esc(i)}</li>" for i in b[1]) + "</ul>"
        elif k == "fields":
            body += "".join(f"<div class='field'><b>{_esc(l)}:</b> <span class='line'></span></div>" for l in b[1])
        elif k == "checks":
            body += "<p>" + " &nbsp; ".join(f"☐ {_esc(o)}" for o in b[1]) + "</p>"
        elif k == "table":
            body += "<table><thead><tr>" + "".join(f"<th>{_esc(h)}</th>" for h in b[1]) + "</tr></thead><tbody>"
            body += "".join("<tr>" + "".join(f"<td>{_esc(v)}</td>" for v in r) + "</tr>" for r in b[2]) + "</tbody></table>"
        elif k == "sign":
            body += "<div class='sign'>" + "".join(f"<div>{_esc(l)}</div><div>Date</div>" for l in b[1]) + "</div>"
    return _pdf(path, body, landscape)
