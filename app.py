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
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
)

# -------------------------- Settings --------------------------
DEFAULT_COMPANY = "AL Glazo Interiors and Décor LLC"
DEFAULT_TITLE = "PAYSLIP"
DEFAULT_DAYS_IN_MONTH = 30
PAGE_SIZES = {"A4": A4, "Letter": LETTER}
FOOTER_SPACER_PT = 48
UI_POWERED_BY_TEXT = 'Powered By <b>Jaseer</b>'  # UI-only footer

# -------------------------- Helpers --------------------------
def clean(s):
    """Trim spaces and hide NaN/None."""
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return ""
    return re.sub(r"\s+", " ", str(s).strip())

def parse_number(x):
    """float or None. Accepts '1,200.50', '-100', '(100)'. Returns None if blank/invalid."""
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return None
    s = str(x).strip()
    if s in ["", "-", "–", "nan", "NaN", "None"]:
        return None
    s = s.replace(",", "")
    m = re.fullmatch(r"\((\d+(\.\d+)?)\)", s)  # (100) -> -100
    if m:
        s = "-" + m.group(1)
    try:
        return float(s)
    except:
        return None

def fmt_amount_basic(x):
    """Format number with no commas; return '0' for blanks/NaN."""
    # x might be a real float (including NaN) or a string
    v = float(x) if isinstance(x, (int, float)) else parse_number(x)
    if v is None or pd.isna(v):
        return "0"
    return f"{int(v)}" if abs(v - int(v)) < 1e-9 else re.sub(r"\.?0+$", "", f"{v:.2f}")

def fmt_amount_commas(x, decimals=2):
    """Format with thousand separators; return '0' for blanks/NaN."""
    v = float(x) if isinstance(x, (int, float)) else parse_number(x)
    if v is None or pd.isna(v):
        return "0"
    if decimals == 0:
        return f"{int(round(v)):,}"
    return f"{v:,.{decimals}f}"

def amount_in_words(x):
    v = float(x) if isinstance(x, (int, float)) else parse_number(x)
    if v is None or pd.isna(v):
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
    s = "" if (s is None or (isinstance(s, float) and pd.isna(s))) else str(s)
    s = re.sub(r"[\\/:*?\"<>|]+", " ", s).strip()
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

    return te_v, td_v, np_v, ap

# -------------------------- PDF builder --------------------------
def build_pdf_for_row(
    row, company_name, title, page_size, days_in_month,
    logo_bytes=None, logo_width=1.2*inch, use_commas=True, decimals=2,
    currency_label="SAR", include_words=True
) -> bytes:
    # choose formatter
    fmt = (lambda x: fmt_amount_commas(x, decimals=decimals)) if use_commas else fmt_amount_basic

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

    # Header with optional logo
    if logo_bytes:
        img = Image(io.BytesIO(logo_bytes))
        img._restrictSize(logo_width, logo_width * 1.2)
        header_tbl = Table(
            [[img, Paragraph(f"<b>{title}</b><br/>{company_name}", ParagraphStyle("hdr", parent=styles["Normal"], fontSize=14, leading=18, alignment=1))]],
            colWidths=[logo_width, None],
        )
        header_tbl.setStyle(TableStyle([
            ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
            ("ALIGN", (1,0), (1,0), "CENTER"),
        ]))
        elems.append(header_tbl)
        elems.append(Spacer(1, 8))
    else:
        elems.append(Paragraph(title, title_style))
        elems.append(Paragraph(company_name, company_style))
        elems.append(Spacer(1, 8))

    # Header table
    hdr_rows = [
        ["Employee Name", clean(row.get("Employee Name", ""))],
        ["Employee Code", clean(row.get("Employee Code", ""))],
        ["Pay Period",    clean(row.get("Pay Period", ""))],
        ["Designation",   clean(row.get("Designation", ""))],
        ["Absent Days",   fmt_amount_basic(row.get("Absent Days", ""))],  # show 0 if blank
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

    rows = [[f"Earnings ({currency_label})", "Amount", f"Deductions ({currency_label})", "Amount"]]
    max_rows = max(len(earnings_order), len(deductions_order))
    for i in range(max_rows):
        Llbl = Lval = Rlbl = Rval = ""
        if i < len(earnings_order):
            lbl, xl = earnings_order[i]
            Llbl = lbl
            Lval = fmt(row.get(xl, ""))
        if i < len(deductions_order):
            lbl, xl = deductions_order[i]
            Rlbl = lbl
            val = row.get(xl, "")
            if xl == "Absent Pay (Deduction)":
                ap = auto_absent_pay(row, days_in_month)
                if ap is not None:
                    val = ap
            Rval = fmt(val)
        rows.append([Llbl, Lval, Rlbl, Rval])

    te, td, np_, _ap = calc_totals(row, days_in_month)
    rows.append(["Total Earnings", fmt(te), "Total Deductions", fmt(td)])
    rows.append(["", "", "Net Pay", fmt(np_)])

    col_widths = [2.8 * inch, 1.2 * inch, 2.8 * inch, 1.2 * inch]
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
    if include_words and np_ is not None and not pd.isna(np_):
        np_words = amount_in_words(np_)
        if np_words:
            elems.append(Paragraph(f"<b>Net to pay (in words):</b> {np_words}", label_style))

    # Signature row
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
st.set_page_config(page_title="PAYSLIP", page_icon="🧾", layout="centered")
st.title("PAYSLIP")

with st.expander("Settings", expanded=True):
    colA, colB, colC = st.columns([2, 1, 1])
    company_name = colA.text_input("Company name", value=DEFAULT_COMPANY)
    days_in_month = colB.number_input("Days in month", min_value=1, max_value=31, value=DEFAULT_DAYS_IN_MONTH, step=1)
    page_size_label = colC.selectbox("Page size", list(PAGE_SIZES.keys()), index=0)
    title = st.text_input("PDF heading text", value=DEFAULT_TITLE)

with st.expander("Formatting", expanded=False):
    c1, c2, c3, c4 = st.columns([1,1,1,1])
    use_commas = c1.checkbox("Use comma separators", value=True)
    fixed_decimals = c2.selectbox("Decimal places", options=[0,2], index=1)
    currency_label = c3.text_input("Currency label", value="SAR")
    include_words = c4.checkbox("Show amount in words", value=True)

with st.expander("Branding", expanded=False):
    logo_file = st.file_uploader("Optional logo (PNG/JPG)", type=["png","jpg","jpeg"])

excel_file = st.file_uploader("Upload Excel (.xlsx)", type=["xlsx"])

# Optional: template button (comment out if not using a bundled template file)
# st.download_button("📥 Download Payslip Template (Excel)", data=open("payslip_template.xlsx","rb").read(), file_name="payslip_template.xlsx")

if excel_file:
    try:
        df = pd.read_excel(excel_file)
    except Exception as e:
        st.error(f"Failed to read Excel file: {e}")
        st.stop()

    df.columns = [str(c).strip() for c in df.columns]
    missing = [c for c in ["Employee Code", "Employee Name"] if c not in df.columns]
    if missing:
        st.error(f"Missing required columns: {missing}")
    else:
        st.success(f"Loaded {len(df)} rows.")
        if st.button("Generate PDFs"):
            zbuf = io.BytesIO()
            page_size = PAGE_SIZES[page_size_label]
            prog = st.progress(0.0)

            logo_bytes = None
            if logo_file:
                logo_bytes = logo_file.read()

            errors = []
            with zipfile.ZipFile(zbuf, "w", zipfile.ZIP_DEFLATED) as zf:
                total = len(df)
                for i, (_, row) in enumerate(df.iterrows(), start=1):
                    try:
                        pdf_bytes = build_pdf_for_row(
                            row=row,
                            company_name=company_name,
                            title=title,
                            page_size=page_size,
                            days_in_month=days_in_month,
                            logo_bytes=logo_bytes,
                            use_commas=use_commas,
                            decimals=fixed_decimals,
                            currency_label=currency_label,
                            include_words=include_words,
                        )
                        emp_code = safe_filename(row.get("Employee Code", ""))
                        emp_name = safe_filename(row.get("Employee Name", ""))
                        fname = f"{emp_code} - {emp_name}".strip() or f"Payslip_{i}"
                        zf.writestr(f"{fname}.pdf", pdf_bytes)
                    except Exception as e:
                        err_name = f"row_{i}_ERROR.txt"
                        zf.writestr(err_name, f"Row {i}: {e}")
                        errors.append((i, str(e)))
                    prog.progress(i / max(1, total))

            zbuf.seek(0)
            run_id = datetime.now().strftime("%Y%m%d-%H%M%S")
            st.download_button(
                "⬇️ Download ZIP of PDFs",
                data=zbuf.read(),
                file_name=f"Payslips_PDF_{run_id}.zip",
                mime="application/zip",
            )

            if errors:
                st.warning(f"{len(errors)} rows had issues. An error file per row was added to the ZIP.")
else:
    st.info("Upload an Excel file to generate payslip PDFs.")

# --- UI footer: Powered By Jaseer (NOT in PDFs) ---
st.markdown(
    """
    <style>
      .stApp { padding-bottom: 60px; }
      .custom-footer {
        position: fixed; left: 0; right: 0; bottom: 0;
        text-align: center; padding: 10px 0;
        color: #6b7280; font-size: 12px; background: rgba(255,255,255,0.7);
      }
    </style>
    <div class="custom-footer">""" + UI_POWERED_BY_TEXT + """</div>
    """,
    unsafe_allow_html=True,
)
