import io
import re
import zipfile
from datetime import datetime

import pandas as pd
import streamlit as st
from num2words import num2words

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, LETTER, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
from openpyxl.utils.cell import column_index_from_string

# -------------------------- Settings --------------------------
DEFAULT_COMPANY = "AL Glazo Interiors and Décor LLC"
DEFAULT_TITLE = "PAYSLIP"
PAGE_SIZES = {"A4 (Landscape)": landscape(A4), "Letter (Landscape)": landscape(LETTER)}
FOOTER_SPACER_PT = 80
UI_POWERED_BY_TEXT = 'Powered By <b>Jaseer</b>'

# ---- Fuzzy header candidates (case-insensitive) ----
EMP_NAME_CANDIDATES = [
    "name","employee name","emp name","staff name","worker name"
]
EMP_CODE_CANDIDATES = [
    "code","employee code","emp code","emp id","employee id","id","staff id","worker id"
]
DESIGNATION_CANDIDATES = [
    "designation","title","position","proffession","profession","job title"
]
ABSENT_DAYS_CANDIDATES = [
    "leave/days","leave days","absent days","absent","leave"
]

# ---- Amount mapping by Excel LETTER (your spec) ----
EARNINGS_LETTERS = {
    "Basic Pay": "F",
    "Other Allowance": "G",
    "Housing Allowance": "H",
    "Over time": "AK",
    "Reward for Full Day Attendance": "AD",
    "Incentive": "AE",
}
DEDUCTIONS_LETTERS = {
    "Absent Pay": ["AJ"],
    "Extra Leave Punishment Ded": ["R"],
    "Air Ticket Deduction": ["S"],
    "Other Fine Ded": ["T"],
    "Medical Deduction": ["U"],
    "Mob Bill Deduction": ["V"],
    "I LOE Insurance Deduction": ["W"],   # LOE caps, Insurance normal
    "Sal Advance Deduction": ["X"],
}
TOTAL_EARNINGS_LETTER   = "AG"
TOTAL_DEDUCTIONS_LETTER = "AH"
NET_PAY_LETTER          = "AL"

# -------------------------- Helpers --------------------------
def clean(s):
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return ""
    return re.sub(r"\s+", " ", str(s).strip())

def parse_number(x):
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return None
    s = str(x).strip()
    if s in ["", "-", "–", "nan", "NaN", "None"]:
        return None
    s = s.replace(",", "")
    m = re.fullmatch(r"\((\d+(\.\d+)?)\)", s)
    if m:
        s = "-" + m.group(1)
    try:
        return float(s)
    except:
        return None

def fmt_amount(x, decimals=2):
    v = float(x) if isinstance(x, (int, float)) else parse_number(x)
    if v is None or pd.isna(v):
        v = 0.0
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
    return re.sub(r"\s+", " ", s) or ""

def letter_value(row_values, letter):
    try:
        idx = column_index_from_string(letter) - 1
        if 0 <= idx < len(row_values):
            return row_values[idx]
    except Exception:
        pass
    return None

def sum_letters(row_values, letters):
    total = 0.0
    any_val = False
    for L in letters:
        v = parse_number(letter_value(row_values, L))
        if v is not None:
            total += v
            any_val = True
    return total if any_val else 0.0

def build_lookup(columns):
    """Return a dict for case-insensitive, trimmed lookup and a normalized list."""
    norm_map = {str(c).strip().lower(): c for c in columns}
    return norm_map

def get_value(row, norm_map, candidates):
    for key in candidates:
        if key in norm_map:
            return row[norm_map[key]]
    # fuzzy: first column containing the candidate token
    cols_lower = {k: v for k, v in norm_map.items()}
    for token in candidates:
        for k in cols_lower:
            if token in k:
                return row[cols_lower[k]]
    return ""

def get_absent_days(row, norm_map):
    v = get_value(row, norm_map, ABSENT_DAYS_CANDIDATES)
    num = parse_number(v)
    return num if num is not None else 0

# -------------------------- PDF builder --------------------------
def build_pdf_for_row(
    std, company_name, title, page_size,
    currency_label="AED", include_words=True, logo_bytes=None, logo_width=1.2*inch
) -> bytes:
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=page_size,
        leftMargin=0.8*inch, rightMargin=0.8*inch,
        topMargin=0.6*inch, bottomMargin=0.6*inch
    )

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle("TitleBig", parent=styles["Title"], fontSize=18, leading=22, alignment=1)
    company_style = ParagraphStyle("Company", parent=styles["Heading2"], fontSize=14, leading=18, alignment=1)
    label_style = ParagraphStyle("Label", parent=styles["Normal"], fontSize=11, leading=14)

    elems = []

    # Header with optional logo
    if logo_bytes:
        img = Image(io.BytesIO(logo_bytes))
        img._restrictSize(logo_width, logo_width*1.2)
        head_tbl = Table(
            [[img, Paragraph(f"<b>{title}</b><br/>{company_name}",
                             ParagraphStyle("hdr", parent=styles["Normal"], fontSize=14, leading=18, alignment=1))]],
            colWidths=[logo_width, None]
        )
        head_tbl.setStyle(TableStyle([
            ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
            ("ALIGN",(1,0),(1,0),"CENTER"),
        ]))
        elems += [head_tbl, Spacer(1,8)]
    else:
        elems += [Paragraph(title, title_style), Paragraph(company_name, company_style), Spacer(1,8)]

    # Employee header table
    hdr_rows = [
        ["Employee Name", clean(std.get("Employee Name",""))],
        ["Employee Code", clean(std.get("Employee Code",""))],
        ["Pay Period",    clean(std.get("Pay Period",""))],
        ["Designation",   clean(std.get("Designation",""))],
        ["Absent Days",   fmt_amount(std.get("Absent Days",0), decimals=0)],
    ]
    hdr_tbl = Table(hdr_rows, colWidths=[2.2*inch, None])
    hdr_tbl.setStyle(TableStyle([
        ("FONTNAME",(0,0),(-1,-1),"Helvetica"),
        ("FONTSIZE",(0,0),(-1,-1),11),
        ("GRID",(0,0),(-1,-1),0.6,colors.black),
        ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
        ("FONTNAME",(0,0),(0,-1),"Helvetica-Bold"),
        ("FONTNAME",(1,0),(1,-1),"Helvetica-Bold"),
    ]))
    elems += [hdr_tbl, Spacer(1,10)]

    # ======= Two-column layout (landscape) =======
    earnings_colwidths = [3.6*inch, 1.4*inch]
    deductions_colwidths = [3.6*inch, 1.4*inch]

    # Earnings table
    earn_rows = [[f"Earnings ({currency_label})", "Amount"]]
    for lbl in ["Basic Pay","Other Allowance","Housing Allowance","Over time",
                "Reward for Full Day Attendance","Incentive"]:
        earn_rows.append([lbl, fmt_amount(std.get(lbl,0),2)])
    earn_rows.append(["Total Earnings", fmt_amount(std.get("Total Earnings (optional)",0),2)])
    earn_tbl = Table(earn_rows, colWidths=earnings_colwidths, repeatRows=1)
    earn_tbl.setStyle(TableStyle([
        ("FONTNAME",(0,0),(-1,-1),"Helvetica"),
        ("FONTSIZE",(0,0),(-1,-1),11),
        ("BACKGROUND",(0,0),(1,0),colors.lightgrey),
        ("GRID",(0,0),(-1,-1),0.6,colors.black),
        ("ALIGN",(1,1),(1,-1),"RIGHT"),
        ("FONTNAME",(0,0),(0,0),"Helvetica-Bold"),
        ("FONTNAME",(0,-1),(1,-1),"Helvetica-Bold"),
    ]))

    # Deductions table
    ded_rows = [[f"Deductions ({currency_label})", "Amount"]]
    for lbl in ["Absent Pay","Extra Leave Punishment Ded","Air Ticket Deduction","Other Fine Ded",
                "Medical Deduction","Mob Bill Deduction","I LOE Insurance Deduction","Sal Advance Deduction"]:
        ded_rows.append([lbl, fmt_amount(std.get(lbl,0),2)])
    ded_rows.append(["Total Deductions", fmt_amount(std.get("Total Deductions (optional)",0),2)])
    ded_tbl = Table(ded_rows, colWidths=deductions_colwidths, repeatRows=1)
    ded_tbl.setStyle(TableStyle([
        ("FONTNAME",(0,0),(-1,-1),"Helvetica"),
        ("FONTSIZE",(0,0),(-1,-1),11),
        ("BACKGROUND",(0,0),(1,0),colors.lightgrey),
        ("GRID",(0,0),(-1,-1),0.6,colors.black),
        ("ALIGN",(1,1),(1,-1),"RIGHT"),
        ("FONTNAME",(0,0),(0,0),"Helvetica-Bold"),
        ("FONTNAME",(0,-1),(1,-1),"Helvetica-Bold"),
    ]))

    # Wrapper with faint border
    two_col = Table([[earn_tbl, ded_tbl]], colWidths=[5.0*inch, 5.0*inch])
    two_col.setStyle(TableStyle([
        ("VALIGN",(0,0),(-1,-1),"TOP"),
        ("LEFTPADDING",(0,0),(-1,-1),0),
        ("RIGHTPADDING",(0,0),(-1,-1),0),
        ("TOPPADDING",(0,0),(-1,-1),0),
        ("BOTTOMPADDING",(0,0),(-1,-1),0),
        ("GRID",(0,0),(-1,-1),0.3,colors.grey),
    ]))
    elems += [two_col, Spacer(1,12)]

    # Centered Summary table
    sum_rows = [
        ["Total Earnings",   fmt_amount(std.get("Total Earnings (optional)",0),2)],
        ["Total Deductions", fmt_amount(std.get("Total Deductions (optional)",0),2)],
        ["Net Pay",          fmt_amount(std.get("Net Pay (optional)",0),2)],
    ]
    sum_tbl = Table(sum_rows, colWidths=[4.6*inch, 1.8*inch], hAlign="CENTER")
    sum_tbl.setStyle(TableStyle([
        ("FONTNAME",(0,0),(-1,-1),"Helvetica"),
        ("FONTSIZE",(0,0),(-1,-1),11),
        ("GRID",(0,0),(-1,-1),0.6,colors.black),
        ("ALIGN",(1,0),(1,-1),"RIGHT"),
        ("FONTNAME",(0,2),(1,2),"Helvetica-Bold"),
        ("BACKGROUND",(0,2),(1,2),colors.whitesmoke),
    ]))
    elems += [sum_tbl]

    # Net in words (optional)
    words = std.get("_net_words","")
    if words:
        elems += [Spacer(1,10), Paragraph(f"<b>Net to pay (in words):</b> {words}", label_style)]

    # Signature row
    elems += [Spacer(1, FOOTER_SPACER_PT)]
    foot = Table([["Accounts","Employee Signature"]], colWidths=[4.0*inch,4.0*inch])
    foot.setStyle(TableStyle([
        ("FONTNAME",(0,0),(-1,-1),"Helvetica"),
        ("FONTSIZE",(0,0),(-1,-1),11),
        ("ALIGN",(0,0),(0,0),"LEFT"),
        ("ALIGN",(1,0),(1,0),"RIGHT"),
    ]))
    elems.append(foot)

    doc.build(elems)
    buf.seek(0)
    return buf.read()

# -------------------------- Streamlit UI --------------------------
st.set_page_config(page_title="PAYSLIP (Landscape)", page_icon="🧾", layout="centered")
st.title("PAYSLIP (Landscape)")

with st.expander("Settings", expanded=True):
    colA, colB, colC = st.columns([2,1,1])
    company_name = colA.text_input("Company name", value=DEFAULT_COMPANY)
    page_size_label = colB.selectbox("Page size", list(PAGE_SIZES.keys()), index=0)
    title = colC.text_input("PDF heading text", value=DEFAULT_TITLE)

with st.expander("Branding & Format", expanded=False):
    c1, c2, c3 = st.columns([1,1,1])
    currency_label = c1.text_input("Currency label", value="AED")
    include_words = c2.checkbox("Show amount in words", value=True)
    logo_file = c3.file_uploader("Optional logo (PNG/JPG)", type=["png","jpg","jpeg"])

excel_file = st.file_uploader("Upload Payroll Excel (.xlsx)", type=["xlsx"])

if not excel_file:
    st.info("Upload your payroll file to generate payslips.")
else:
    try:
        df = pd.read_excel(excel_file)  # first sheet
    except Exception as e:
        st.error(f"Failed to read Excel file: {e}")
        st.stop()

    # Build a normalized header lookup once
    norm_map = build_lookup(df.columns)

    st.success(f"Loaded {len(df)} rows.")
    if st.button("Generate PDFs"):
        zbuf = io.BytesIO()
        page_size = PAGE_SIZES[page_size_label]
        prog = st.progress(0.0)
        logo_bytes = logo_file.read() if logo_file else None

        with zipfile.ZipFile(zbuf, "w", zipfile.ZIP_DEFLATED) as zf:
            total = len(df)
            values = df.values
            for i in range(total):
                row_series = df.iloc[i]
                row_vals = values[i]

                emp_name = clean(get_value(row_series, norm_map, EMP_NAME_CANDIDATES))
                emp_code = clean(get_value(row_series, norm_map, EMP_CODE_CANDIDATES))
                designation = clean(get_value(row_series, norm_map, DESIGNATION_CANDIDATES))
                absent_days = get_absent_days(row_series, norm_map)

                std = {
                    "Employee Name": emp_name,
                    "Employee Code": emp_code,
                    "Designation": designation,
                    "Pay Period": "",      # map another header if you have it
                    "Absent Days": absent_days,
                }

                # Earnings
                for lbl, L in EARNINGS_LETTERS.items():
                    std[lbl] = parse_number(letter_value(row_vals, L)) or 0.0

                # Deductions
                for lbl, Ls in DEDUCTIONS_LETTERS.items():
                    std[lbl] = sum_letters(row_vals, Ls)

                # Totals
                te   = parse_number(letter_value(row_vals, TOTAL_EARNINGS_LETTER))
                td   = parse_number(letter_value(row_vals, TOTAL_DEDUCTIONS_LETTER))
                npay = parse_number(letter_value(row_vals, NET_PAY_LETTER))

                if te is None:
                    te = sum(std.get(k,0.0) for k in EARNINGS_LETTERS.keys())
                if td is None:
                    td = sum(std.get(k,0.0) for k in DEDUCTIONS_LETTERS.keys())
                if npay is None:
                    npay = te - td

                std["Total Earnings (optional)"]   = te
                std["Total Deductions (optional)"] = td
                std["Net Pay (optional)"]          = npay
                std["_net_words"] = amount_in_words(npay) if include_words else ""

                try:
                    pdf_bytes = build_pdf_for_row(
                        std, company_name, title, page_size,
                        currency_label=currency_label, include_words=include_words,
                        logo_bytes=logo_bytes
                    )

                    # Build filename from non-empty parts
                    name_part = safe_filename(emp_name)
                    code_part = safe_filename(emp_code)
                    parts = [p for p in [code_part, name_part] if p]
                    base = " - ".join(parts) if parts else f"Payslip_{i+1}"
                    fname = base + ".pdf"

                    zf.writestr(fname, pdf_bytes)
                except Exception as e:
                    zf.writestr(f"row_{i+1}_ERROR.txt", f"Row {i+1}: {e}")

                prog.progress((i+1)/max(1,total))

        zbuf.seek(0)
        run_id = datetime.now().strftime("%Y%m%d-%H%M%S")
        st.download_button(
            "⬇️ Download ZIP of PDFs",
            data=zbuf.read(),
            file_name=f"Payslips_Landscape_{run_id}.zip",
            mime="application/zip",
        )

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
