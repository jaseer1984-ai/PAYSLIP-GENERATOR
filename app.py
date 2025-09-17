import io
import re
import zipfile
from datetime import datetime

import pandas as pd
import streamlit as st
from num2words import num2words

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, LETTER
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle

# -------------------------- Settings --------------------------
DEFAULT_COMPANY = "AL Glazo Interiors and Décor LLC"
DEFAULT_TITLE = "PAYSLIP"          # PDF heading text
DEFAULT_DAYS_IN_MONTH = 30
PAGE_SIZES = {"A4": A4, "Letter": LETTER}
FOOTER_SPACER_PT = 48              # move footer lower (increase if needed)

# -------------------------- Helpers --------------------------
def clean(s):
    return re.sub(r"\s+", " ", str(s).strip())

def parse_number(x):
    """float or None. Accepts '1,200.50', '-100', '(100)'."""
    if x is None:
        return None
    s = str(x).strip()
    if s in ["", "-", "–"]:
        return None
    s = s.replace(",", "")
    m = re.fullmatch(r"\((\d+(\.\d+)?)\)", s)  # (100) -> -100
    if m:
        s = "-" + m.group(1)
    try:
        return float(s)
    except:
        return None

def fmt_amount(x):
    """int if whole, else up to 2 dp, no commas."""
    v = parse_number(x) if not isinstance(x, (int, float)) else float(x)
    if v is None:
        return ""
    return f"{int(v)}" if abs(v - int(v)) < 1e-9 else re.sub(r"\.?0+$", "", f"{v:.2f}")

def amount_in_words(x):
    v = parse_number(x)
    if v is None:
        return ""
    sign = "minus " if v < 0 else ""
    v = abs(v)
    whole = int(v)
    frac = int(round((v - whole) * 100))
    words = num2words(whole, lang="en").replace("-", " ")
    if frac:
        words = f"{words} and {frac:02d}/100"
    return (sign + words).strip().capitalize() + " only"

def safe_filename(s):
    s = re.sub(r"[\\/:*?\"<>|]+", " ", str(s)).strip()
    return re.sub(r"\s+", " ", s) or "Payslip"

def auto_absent_pay(row, days_in_month):
    """If 'Absent Pay (Deduction)' is blank, derive from Basic Pay and Absent Days."""
    ap = parse_number(row.get("Absent Pay (Deduction)", ""))
    if ap is not None:
        return ap
    basic = parse_number(row.get("Basic Pay", ""))
    days = parse_number(row.get("Absent Days", ""))
    if basic is None or days is None:
        return None
    return round(basic / float(days_in_month) * days, 2)

def calc_totals(row, days_in_month):
    """Compute Total Earnings, Total Deductions, Net Pay (if not provided)."""
    ot_header = "Over time" if "Over time" in row.index else ("Overtime" if "Overtime" in row.index else None)
    earn_cols = ["Basic Pay", "Other Allowance", "Housing Allowance", ot_header,
                 "Reward (Full Day Attendance)", "Incentive"]
    earn_cols = [c for c in earn_cols if c]
    ded_cols = ["Absent Pay (Deduction)", "Salary Advance (Deduction)",
                "Ticket / Other Ded. (Deduction)", "Extra Leave / Punishment (Deduction)"]

    r = row.copy()
    ap = auto_absent_pay(r, days_in_month)
    if ap is not None:
        r["Absent Pay (Deduction)"] = ap

    def total(cols):
        s = 0.0
        used = False
        for c in cols:
            if c in r.index:
                v = parse_number(r[c])
                if v is not None:
                    s += v
                    used = True
        return 0.0 if not used else s

    te_v = parse_number(r.get("Total Earnings (optional)", ""))
    td_v = parse_number(r.get("Total Deductions (optional)", ""))
    if te_v is None:
        te_v = total(earn_cols)
    if td_v is None:
        td_v = total(ded_cols)

    np_v = parse_number(r.get("Net Pay (optional)", ""))
    if np_v is None:
        np_v = te_v - td_v

    return fmt_amount(te_v), fmt_amount(td_v), fmt_amount(np_v), fmt_amount(ap)

# -------------------------- PDF builder --------------------------
def build_pdf_for_row(row, company_name, title, page_size, days_in_month) -> bytes:
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=page_size,
        leftMargin=0.6 * inch,
        rightMargin=0.6 * inch,
        topMargin=0.6 * inch,
        bottomMargin=0.6 * inch,
    )

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle("TitleBig", parent=styles["Title"], fontSize=18, leading=22, alignment=1)
    company_style = ParagraphStyle("Company", parent=styles["Heading2"], fontSize=14, leading=18, alignment=1)
    label_style = ParagraphStyle("Label", parent=styles["Normal"], fontSize=11, leading=14)

    elems = []
    elems.append(Paragraph(title, title_style))
    elems.append(Paragraph(company_name, company_style))
    elems.append(Spacer(1, 8))

    # Header table (Label | Value)
    hdr_rows = [
        ["Employee Name", clean(row.get("Employee Name", ""))],
        ["Employee Code", clean(row.get("Employee Code", ""))],
        ["Pay Period",    clean(row.get("Pay Period", ""))],
        ["Designation",   clean(row.get("Designation", ""))],
        ["Absent Days",   clean(row.get("Absent Days", ""))],
    ]
    hdr_tbl = Table(hdr_rows, colWidths=[2.0 * inch, None])
    hdr_tbl.setStyle(TableStyle([
        ("FONTNAME", (0,0), (-1,-1), "Helvetica"),
        ("FONTSIZE", (0,0), (-1,-1), 11),
        ("GRID", (0,0), (-1,-1), 0.6, colors.black),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("FONTNAME", (0,0), (0,-1), "Helvetica-Bold"),
        ("FONTNAME", (1,0), (1,-1), "Helvetica-Bold"),
    ]))
    elems.append(hdr_tbl)
    elems.append(Spacer(1, 8))

    # Earnings / Deductions table
    overtime_header = "Over time" if "Over time" in row.index else ("Overtime" if "Overtime" in row.index else "Over time")
    earnings_order = [
        ("Basic Pay", "Basic Pay"),
        ("Other Allowance", "Other Allowance"),
        ("Housing Allowance", "Housing Allowance"),
        ("Over time", overtime_header),
        ("Reward for Full Day Attendance", "Reward (Full Day Attendance)"),
        ("Incentive", "Incentive"),
    ]
    deductions_order = [
        ("Absent Pay", "Absent Pay (Deduction)"),
        ("Salary Adv Pay", "Salary Advance (Deduction)"),
        ("Ticket / Other Ded.", "Ticket / Other Ded. (Deduction)"),
        ("Extra Leave Punishment Ded", "Extra Leave / Punishment (Deduction)"),
    ]

    rows = [["Earnings", "Amount", "Deductions", "Amount"]]
    max_rows = max(len(earnings_order), len(deductions_order))
    for i in range(max_rows):
        Llbl = Lval = Rlbl = Rval = ""
        if i < len(earnings_order):
            lbl, xl = earnings_order[i]
            Llbl = lbl
            Lval = fmt_amount(row.get(xl, ""))
        if i < len(deductions_order):
            lbl, xl = deductions_order[i]
            Rlbl = lbl
            val = row.get(xl, "")
            if xl == "Absent Pay (Deduction)":
                ap = auto_absent_pay(row, days_in_month)
                if ap is not None:
                    val = ap
            Rval = fmt_amount(val)
        rows.append([Llbl, Lval, Rlbl, Rval])

    te, td, np_, _ap = calc_totals(row, days_in_month)
    rows.append(["Total Earnings", te, "Total Deductions", td])
    rows.append(["", "", "Net Pay", np_])

    col_widths = [2.6 * inch, 1.2 * inch, 2.9 * inch, 1.2 * inch]
    tbl = Table(rows, colWidths=col_widths, repeatRows=1)
    tbl.setStyle(TableStyle([
        ("FONTNAME", (0,0), (-1,-1), "Helvetica"),
        ("FONTSIZE", (0,0), (-1,-1), 11),
        ("BACKGROUND", (0,0), (3,0), colors.lightgrey),
        ("GRID", (0,0), (-1,-1), 0.6, colors.black),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("ALIGN", (1,1), (1,-1), "RIGHT"),
        ("ALIGN", (3,1), (3,-1), "RIGHT"),
        ("FONTNAME", (0,0), (0,0), "Helvetica-Bold"),
        ("FONTNAME", (2,0), (2,0), "Helvetica-Bold"),
        ("FONTNAME", (0,-2), (1,-2), "Helvetica-Bold"),
        ("FONTNAME", (2,-2), (3,-2), "Helvetica-Bold"),
        ("FONTNAME", (2,-1), (3,-1), "Helvetica-Bold"),
    ]))
    elems.append(tbl)

    # Net to pay (in words)
    elems.append(Spacer(1, 12))
    np_words = amount_in_words(np_)
    if np_words:
        elems.append(Paragraph(f"<b>Net to pay (in words):</b> {np_words}", label_style))

    # Footer moved downward
    elems.append(Spacer(1, FOOTER_SPACER_PT))
    foot = Table([["Accounts", "Employee Signature"]], colWidths=[3.5 * inch, 3.5 * inch])
    foot.setStyle(TableStyle([
        ("FONTNAME", (0,0), (-1,-1), "Helvetica"),
        ("FONTSIZE", (0,0), (-1,-1), 11),
        ("ALIGN", (0,0), (0,0), "LEFT"),
        ("ALIGN", (1,0), (1,0), "RIGHT"),
    ]))
    elems.append(foot)

    doc.build(elems)
    buf.seek(0)
    return buf.read()

# -------------------------- Streamlit UI --------------------------
st.set_page_config(page_title="PAYSLIP", page_icon="🧾", layout="centered")  # browser tab title
st.title("PAYSLIP")  # main heading

# Settings bar
with st.expander("Settings", expanded=True):
    colA, colB, colC = st.columns([2, 1, 1])
    company_name = colA.text_input("Company name", value=DEFAULT_COMPANY)
    days_in_month = colB.number_input("Days in month", min_value=1, max_value=31, value=DEFAULT_DAYS_IN_MONTH, step=1)
    page_size_label = colC.selectbox("Page size", list(PAGE_SIZES.keys()), index=0)
    # PDF document heading (kept editable; default = PAYSLIP)
    title = st.text_input("PDF heading text", value=DEFAULT_TITLE)

# Move help & template to collapsed sidebar (kept off the main dashboard)
def make_template_xlsx() -> bytes:
    cols = [
        "Employee Code","Employee Name","Pay Period","Designation","Absent Days",
        "Basic Pay","Other Allowance","Housing Allowance","Over time",
        "Reward (Full Day Attendance)","Incentive",
        "Absent Pay (Deduction)","Salary Advance (Deduction)",
        "Ticket / Other Ded. (Deduction)","Extra Leave / Punishment (Deduction)",
        "Total Earnings (optional)","Total Deductions (optional)","Net Pay (optional)"
    ]
    data = [
        ["AG-0213","AJIT KUMAR AJABDAYAL RAM","AUGUST 2025","FITTER",0,1200,300,"",325,100,0,"","","","", "", "", ""],
        ["AG-0401","RAHUL SHARMA","AUGUST 2025","ELECTRICIAN",1,1500,250,300,200,0,50,"100",0,0,0,"","",""],
    ]
    df = pd.DataFrame(data, columns=cols)
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name="Data")
    bio.seek(0)
    return bio.read()

with st.sidebar.expander("Help & Template", expanded=False):
    st.markdown(
        "**Required Excel headers**: `Employee Code`, `Employee Name`  \n"
        "**Optional**: `Pay Period`, `Designation`, `Absent Days`, "
        "`Basic Pay`, `Other Allowance`, `Housing Allowance`, `Over time` (or `Overtime`), "
        "`Reward (Full Day Attendance)`, `Incentive`, `Absent Pay (Deduction)`, "
        "`Salary Advance (Deduction)`, `Ticket / Other Ded. (Deduction)`, "
        "`Extra Leave / Punishment (Deduction)`, `Total Earnings (optional)`, "
        "`Total Deductions (optional)`, `Net Pay (optional)`"
    )
    st.download_button(
        "⬇️ Download Excel template",
        data=make_template_xlsx(),
        file_name="Payslip_Input_Template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

# Upload Excel
excel_file = st.file_uploader("Upload Excel (.xlsx)", type=["xlsx"])
if excel_file:
    df = pd.read_excel(excel_file)
    df.columns = [str(c).strip() for c in df.columns]
    missing = [c for c in ["Employee Code", "Employee Name"] if c not in df.columns]
    if missing:
        st.error(f"Missing required columns: {missing}")
    else:
        st.success(f"Loaded {len(df)} rows.")
        st.dataframe(df.head(10))
        if st.button("Generate PDFs"):
            zbuf = io.BytesIO()
            with zipfile.ZipFile(zbuf, "w", zipfile.ZIP_DEFLATED) as zf:
                page_size = PAGE_SIZES[page_size_label]
                prog = st.progress(0.0)
                for i, (_, row) in enumerate(df.iterrows(), start=1):
                    try:
                        pdf_bytes = build_pdf_for_row(row, company_name, title, page_size, days_in_month)
                        emp_code = safe_filename(row.get("Employee Code", ""))
                        emp_name = safe_filename(row.get("Employee Name", ""))
                        fname = f"{emp_code} - {emp_name}".strip() or f"Payslip_{i}"
                        zf.writestr(f"{fname}.pdf", pdf_bytes)
                    except Exception as e:
                        zf.writestr(f"row_{i}_ERROR.txt", f"Row {i}: {e}")
                    prog.progress(i / len(df))
            zbuf.seek(0)
            run_id = datetime.now().strftime("%Y%m%d-%H%M%S")
            st.download_button(
                "⬇️ Download ZIP of PDFs",
                data=zbuf.read(),
                file_name=f"Payslips_PDF_{run_id}.zip",
                mime="application/zip",
            )
