import io
import re
import zipfile
import traceback
from datetime import datetime

import pandas as pd
import streamlit as st
from num2words import num2words

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, LETTER, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
)
from reportlab.platypus.flowables import KeepInFrame, CondPageBreak
from openpyxl.utils.cell import column_index_from_string

# ========================== SETTINGS ==========================
DEFAULT_COMPANY = "AL Glazo Interiors and Décor LLC"
DEFAULT_TITLE = "PAYSLIP"
PAGE_SIZES = {"A4 (Landscape)": landscape(A4), "Letter (Landscape)": landscape(LETTER)}
UI_POWERED_BY_TEXT = 'Powered By <b>Jaseer</b>'

TABLE_FONT_SIZE = 10
HEADER_TITLE_SIZE = 18
HEADER_COMPANY_SIZE = 13

EMP_NAME_CANDIDATES = ["employee name", "name", "emp name", "staff name", "worker name"]
EMP_CODE_CANDIDATES = ["employee code", "emp code", "emp id", "employee id", "code", "id", "staff id", "worker id"]
DESIGNATION_CANDIDATES = ["designation", "title", "position", "proffession", "profession", "job title"]
ABSENT_DAYS_CANDIDATES = ["leave/days", "leave days", "absent days", "absent", "leave"]
PAY_PERIOD_CANDIDATES = ["pay period", "period", "month", "pay month"]
RATE_CANDIDATES = [
    "rate","hourly rate","per hour","cost per hour","wage","salary/hour","hr rate","hour rate"
]

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
    "I LOE Insurance Deduction": ["W"],
    "Sal Advance Deduction": ["X"],
}
TOTAL_EARNINGS_LETTER   = "AG"
TOTAL_DEDUCTIONS_LETTER = "AH"
NET_PAY_LETTER          = "AL"

# ========================== GENERIC HELPERS ==========================
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
        # accept HH:MM[:SS] as hours
        if re.fullmatch(r"\d{1,2}:\d{2}(:\d{2})?", s):
            h, m2, *rest = s.split(":")
            sec = int(rest[0]) if rest else 0
            return int(h) + int(m2)/60 + sec/3600
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

def build_lookup(columns):
    return {str(c).strip().lower(): c for c in columns}

def get_value(row, norm_map, candidates):
    for key in candidates:
        if key in norm_map:
            return row[norm_map[key]]
    for token in candidates:
        for norm, orig in norm_map.items():
            if token in norm:
                return row[orig]
    return ""

# ----- payroll column helpers -----
def letter_value(row_values, letter, max_cols=None):
    try:
        idx = column_index_from_string(letter) - 1
        if max_cols is not None and idx >= max_cols:
            return None
        return row_values[idx]
    except Exception:
        return None

def sum_letters(row_values, letters, max_cols=None):
    total = 0.0
    used = False
    for L in letters:
        v = parse_number(letter_value(row_values, L, max_cols))
        if v is not None:
            total += v
            used = True
    return total if used else 0.0

def get_absent_days(row, norm_map):
    v = get_value(row, norm_map, ABSENT_DAYS_CANDIDATES)
    num = parse_number(v)
    return num if num is not None else 0

# ========================== PAYSLIP PDF ==========================
def build_pdf_for_row(std, company_name, title, page_size,
                      currency_label="AED", include_words=True, logo_bytes=None, logo_width=1.2*inch) -> bytes:
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=page_size,
        leftMargin=0.8*inch, rightMargin=0.8*inch, topMargin=0.6*inch, bottomMargin=0.6*inch
    )
    doc.allowSplitting = 1

    styles = getSampleStyleSheet()
    title_style   = ParagraphStyle("TitleBig", parent=styles["Title"],   fontSize=HEADER_TITLE_SIZE, leading=20, alignment=1)
    company_style = ParagraphStyle("Company",  parent=styles["Heading2"], fontSize=HEADER_COMPANY_SIZE, leading=16, alignment=1)
    label_style   = ParagraphStyle("Label",    parent=styles["Normal"],   fontSize=TABLE_FONT_SIZE,   leading=TABLE_FONT_SIZE+2)

    elems = []
    if logo_bytes:
        img = Image(io.BytesIO(logo_bytes)); img._restrictSize(logo_width, logo_width*1.2)
        head_tbl = Table([[img, Paragraph(f"<b>{title}</b><br/>{company_name}",
                             ParagraphStyle("hdr", parent=styles["Normal"], fontSize=HEADER_COMPANY_SIZE, leading=16, alignment=1))]],
                         colWidths=[logo_width, None])
        head_tbl.setStyle(TableStyle([
            ("VALIGN",(0,0),(-1,-1),"MIDDLE"), ("ALIGN",(1,0),(1,0),"CENTER"),
            ("LEFTPADDING",(0,0),(-1,-1),0), ("RIGHTPADDING",(0,0),(-1,-1),0),
            ("TOPPADDING",(0,0),(-1,-1),0),  ("BOTTOMPADDING",(0,0),(-1,-1),2),
        ]))
        elems += [head_tbl, Spacer(1,6)]
    else:
        elems += [Paragraph(title, title_style), Paragraph(company_name, company_style), Spacer(1,6)]

    hdr_rows = [
        ["Employee Name", clean(std.get("Employee Name",""))],
        ["Employee Code", clean(std.get("Employee Code",""))],
        ["Pay Period",    clean(std.get("Pay Period",""))],
        ["Designation",   clean(std.get("Designation",""))],
        ["Absent Days",   fmt_amount(std.get("Absent Days",0), 0)],
    ]
    hdr_tbl = Table(hdr_rows, colWidths=[2.3*inch, None])
    hdr_tbl.setStyle(TableStyle([
        ("FONTNAME",(0,0),(-1,-1),"Helvetica"), ("FONTSIZE",(0,0),(-1,-1),TABLE_FONT_SIZE),
        ("GRID",(0,0),(-1,-1),0.6,colors.black), ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
        ("LEFTPADDING",(0,0),(-1,-1),3), ("RIGHTPADDING",(0,0),(-1,-1),3),
        ("TOPPADDING",(0,0),(-1,-1),2), ("BOTTOMPADDING",(0,0),(-1,-1),2),
        ("FONTNAME",(0,0),(0,-1),"Helvetica-Bold"), ("FONTNAME",(1,0),(1,-1),"Helvetica-Bold"),
    ]))
    elems += [hdr_tbl, Spacer(1,6)]

    earn_rows = [[f"Earnings ({currency_label})","Amount"]]
    for lbl in ["Basic Pay","Other Allowance","Housing Allowance","Over time","Reward for Full Day Attendance","Incentive"]:
        earn_rows.append([lbl, fmt_amount(std.get(lbl,0),2)])
    earn_rows.append(["Total Earnings", fmt_amount(std.get("Total Earnings (optional)",0),2)])
    earn_tbl = Table(earn_rows, colWidths=[3.6*inch,1.4*inch], repeatRows=1)
    earn_tbl.setStyle(TableStyle([
        ("FONTNAME",(0,0),(-1,-1),"Helvetica"), ("FONTSIZE",(0,0),(-1,-1),TABLE_FONT_SIZE),
        ("BACKGROUND",(0,0),(1,0),colors.lightgrey), ("GRID",(0,0),(-1,-1),0.6,colors.black),
        ("ALIGN",(1,1),(1,-1),"RIGHT"),
        ("LEFTPADDING",(0,0),(-1,-1),3), ("RIGHTPADDING",(0,0),(-1,-1),3),
        ("TOPPADDING",(0,0),(-1,-1),2), ("BOTTOMPADDING",(0,0),(-1,-1),2),
        ("FONTNAME",(0,0),(0,0),"Helvetica-Bold"), ("FONTNAME",(0,-1),(1,-1),"Helvetica-Bold"),
    ]))

    ded_rows = [[f"Deductions ({currency_label})","Amount"]]
    for lbl in ["Absent Pay","Extra Leave Punishment Ded","Air Ticket Deduction","Other Fine Ded",
                "Medical Deduction","Mob Bill Deduction","I LOE Insurance Deduction","Sal Advance Deduction"]:
        ded_rows.append([lbl, fmt_amount(std.get(lbl,0),2)])
    ded_rows.append(["Total Deductions", fmt_amount(std.get("Total Deductions (optional)",0),2)])
    ded_tbl = Table(ded_rows, colWidths=[3.6*inch,1.4*inch], repeatRows=1)
    ded_tbl.setStyle(TableStyle([
        ("FONTNAME",(0,0),(-1,-1),"Helvetica"), ("FONTSIZE",(0,0),(-1,-1),TABLE_FONT_SIZE),
        ("BACKGROUND",(0,0),(1,0),colors.lightgrey), ("GRID",(0,0),(-1,-1),0.6,colors.black),
        ("ALIGN",(1,1),(1,-1),"RIGHT"),
        ("LEFTPADDING",(0,0),(-1,-1),3), ("RIGHTPADDING",(0,0),(-1,-1),3),
        ("TOPPADDING",(0,0),(-1,-1),2), ("BOTTOMPADDING",(0,0),(-1,-1),2),
        ("FONTNAME",(0,0),(0,0),"Helvetica-Bold"), ("FONTNAME",(0,-1),(1,-1),"Helvetica-Bold"),
    ]))

    two_col = Table([[earn_tbl, ded_tbl]], colWidths=[5.0*inch, 5.0*inch])
    two_col.setStyle(TableStyle([
        ("VALIGN",(0,0),(-1,-1),"TOP"),
        ("LEFTPADDING",(0,0),(-1,-1),0), ("RIGHTPADDING",(0,0),(-1,-1),0),
        ("TOPPADDING",(0,0),(-1,-1),0),  ("BOTTOMPADDING",(0,0),(-1,-1),0),
        ("GRID",(0,0),(-1,-1),0.3,colors.grey),
    ]))
    elems += [two_col, Spacer(1,5)]

    sum_rows = [
        ["Total Earnings",   fmt_amount(std.get("Total Earnings (optional)",0),2)],
        ["Total Deductions", fmt_amount(std.get("Total Deductions (optional)",0),2)],
        ["Net Pay",          fmt_amount(std.get("Net Pay (optional)",0),2)],
    ]
    sum_tbl = Table(sum_rows, colWidths=[4.6*inch, 1.8*inch], hAlign="CENTER")
    sum_tbl.setStyle(TableStyle([
        ("FONTNAME",(0,0),(-1,-1),"Helvetica"), ("FONTSIZE",(0,0),(-1,-1),TABLE_FONT_SIZE),
        ("GRID",(0,0),(-1,-1),0.6,colors.black), ("ALIGN",(1,0),(1,-1),"RIGHT"),
        ("LEFTPADDING",(0,0),(-1,-1),3), ("RIGHTPADDING",(0,0),(-1,-1),3),
        ("TOPPADDING",(0,0),(-1,-1),2), ("BOTTOMPADDING",(0,0),(-1,-1),2),
        ("FONTNAME",(0,2),(1,2),"Helvetica-Bold"),
        ("BACKGROUND",(0,2),(1,2),colors.whitesmoke),
    ]))

    summary_block = [sum_tbl]
    words = amount_in_words(std.get("Net Pay (optional)",0))
    if words:
        summary_block += [Spacer(1,4), Paragraph(f"<b>Net to pay (in words):</b> {words}", label_style)]

    elems.append(KeepInFrame(maxWidth=None, maxHeight=1.8*inch, content=summary_block, mergeSpace=1, mode="shrink"))
    elems.append(CondPageBreak(0.6*inch))

    elems.append(Spacer(1, 18))
    foot = Table([["Accounts","Employee Signature"]], colWidths=[4.0*inch, 4.0*inch])
    foot.setStyle(TableStyle([
        ("FONTNAME",(0,0),(-1,-1),"Helvetica"), ("FONTSIZE",(0,0),(-1,-1),TABLE_FONT_SIZE),
        ("ALIGN",(0,0),(0,0),"LEFT"), ("ALIGN",(1,0),(1,0),"RIGHT"),
    ]))
    elems.append(foot)

    doc.build(elems)
    buf.seek(0)
    return buf.read()

# ========================== STREAMLIT UI (PAYSLIP) ==========================
st.set_page_config(page_title="PAYSLIP", page_icon="🧾", layout="centered")
st.title("PAYSLIP")

with st.expander("Settings", expanded=True):
    colA, colB = st.columns([2,1])
    company_name = colA.text_input("Company name", value=DEFAULT_COMPANY)
    page_size_label = colB.selectbox("Page size", list(PAGE_SIZES.keys()), index=0)
    default_hourly_rate = colA.number_input("Default hourly rate (if sheet has no rate)", min_value=0.0, value=0.0, step=0.5)
    default_ot_multiplier = colB.number_input("OT multiplier (× rate)", min_value=1.0, value=1.25, step=0.05)

title = st.text_input("PDF heading text", value=DEFAULT_TITLE)

with st.expander("Branding (optional)", expanded=False):
    c1, c2 = st.columns([1,1])
    currency_label = c1.text_input("Currency label", value="AED")
    logo_file = c2.file_uploader("Logo (PNG/JPG)", type=["png","jpg","jpeg"])

excel_file = st.file_uploader("Upload Payroll Excel (.xlsx)", type=["xlsx"])


def build_std_for_row(row_series, row_vals, norm_map, max_cols, pay_period_text=""):
    emp_name = clean(get_value(row_series, norm_map, EMP_NAME_CANDIDATES))
    emp_code = clean(get_value(row_series, norm_map, EMP_CODE_CANDIDATES))
    designation = clean(get_value(row_series, norm_map, DESIGNATION_CANDIDATES))
    absent_days = get_absent_days(row_series, norm_map)
    pay_period = pay_period_text or clean(get_value(row_series, norm_map, PAY_PERIOD_CANDIDATES))

    std = {
        "Employee Name": emp_name,
        "Employee Code": emp_code,
        "Designation": designation,
        "Pay Period": pay_period,
        "Absent Days": absent_days,
    }
    for lbl, L in EARNINGS_LETTERS.items():
        std[lbl] = parse_number(letter_value(row_vals, L, max_cols)) or 0.0
    for lbl, Ls in DEDUCTIONS_LETTERS.items():
        std[lbl] = sum_letters(row_vals, Ls, max_cols)

    te = parse_number(letter_value(row_vals, TOTAL_EARNINGS_LETTER, max_cols))
    td = parse_number(letter_value(row_vals, TOTAL_DEDUCTIONS_LETTER, max_cols))
    npay = parse_number(letter_value(row_vals, NET_PAY_LETTER, max_cols))
    if te is None:
        te = sum(std.get(k, 0.0) for k in EARNINGS_LETTERS.keys())
    if td is None:
        td = sum(std.get(k, 0.0) for k in DEDUCTIONS_LETTERS.keys())
    if npay is None:
        npay = te - td

    std["Total Earnings (optional)"] = te
    std["Total Deductions (optional)"] = td
    std["Net Pay (optional)"] = npay
    return std

if excel_file:
    try:
        xls = pd.ExcelFile(excel_file)
        sheet_name = xls.sheet_names[0]
        df = pd.read_excel(xls, sheet_name=sheet_name, header=0)
        df.columns = [str(c).strip() for c in df.columns]
        st.success(f"Loaded {len(df)} rows from sheet '{sheet_name}'. Pay Period will be '{sheet_name}'.")
    except Exception as e:
        st.error(f"Failed to read Excel file: {e}")
        st.stop()

    if st.button("Generate Payslips (ZIP)"):
        zbuf = io.BytesIO()
        logo_bytes = logo_file.read() if logo_file else None
        page_size = PAGE_SIZES[page_size_label]

        with zipfile.ZipFile(zbuf, "w", zipfile.ZIP_DEFLATED) as zf:
            values = df.values
            for i in range(len(df)):
                row_series = df.iloc[i]
                row_vals = values[i]
                try:
                    norm_map = build_lookup(df.columns)
                    std = build_std_for_row(row_series, row_vals, norm_map, df.shape[1], pay_period_text=sheet_name)
                    pdf_bytes = build_pdf_for_row(
                        std, company_name, title, page_size,
                        currency_label=currency_label, include_words=True, logo_bytes=logo_bytes
                    )
                    parts = [safe_filename(std["Employee Code"]), safe_filename(std["Employee Name"])]
                    parts = [p for p in parts if p]
                    fname = (" - ".join(parts) if parts else f"Payslip_{i+1}") + ".pdf"
                    zf.writestr(fname, pdf_bytes)
                except Exception as e:
                    tb = traceback.format_exc()
                    zf.writestr(f"row_{i+1}_ERROR.txt", f"Row {i+1}: {e}\n\n{tb}")

        zbuf.seek(0)
        run_id = datetime.now().strftime("%Y%m%d-%H%M%S")
        st.download_button("⬇️ Download Payslips (ZIP)", data=zbuf.read(),
                           file_name=f"Payslips_{sheet_name}_{run_id}.zip", mime="application/zip")

# ==========================================================
#       OVERTIME REPORT (WIDE 1..31 FORMAT) + ABSENT COUNT
# ==========================================================
st.markdown("---")
st.subheader("Overtime Report (Daily > threshold)")

att_file = st.file_uploader("Upload Attendance Excel (.xlsx)", type=["xlsx"], key="attendance")
ot_threshold = st.number_input("Daily OT threshold (hours)", min_value=0.0, max_value=24.0, value=8.0, step=0.5)

# ---- helpers (attendance) ----
def norm_cols_map(columns):
    return {str(c).strip().lower(): c for c in columns}

def pick_col(norm_map, *candidates):
    for cand in candidates:
        if cand in norm_map:
            return norm_map[cand]
    for cand in candidates:
        for k, v in norm_map.items():
            if cand in k:
                return v
    return None

ABSENT_TOKENS = {"a", "absent"}
PRESENT_TOKENS = {"p", "present"}
OFF_TOKENS = {"off","leave","holiday"}

def is_absent_cell(x):
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return False
    s = str(x).strip().lower()
    return s in ABSENT_TOKENS

def is_present_token(x):
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return False
    return str(x).strip().lower() in PRESENT_TOKENS

def parse_hours_cell(x):
    """Return hours as float (None if not hours). A=0, P=8, OFF/LEAVE/HOLIDAY=None."""
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return None
    s = str(x).strip()
    sl = s.lower()
    if sl in OFF_TOKENS or sl in {"-","--",""}:
        return None
    if sl in ABSENT_TOKENS:
        return 0.0
    if sl in PRESENT_TOKENS:
        return 8.0
    # HH:MM[:SS]
    if re.fullmatch(r"\d{1,2}:\d{2}(:\d{2})?", s):
        h, m, *rest = s.split(":")
        sec = int(rest[0]) if rest else 0
        return int(h) + int(m)/60 + sec/3600
    # 8h 30m or 8 Hrs
    m = re.search(r"(\d+(\.\d+)?)\s*h", s, re.I)
    if m:
        hours = float(m.group(1))
        m2 = re.search(r"(\d+)\s*m", s, re.I)
        if m2:
            hours += int(m2.group(1))/60
        return hours
    try:
        return float(s)
    except:
        return None

def coerce_day_label(val):
    if pd.isna(val):
        return None
    s = str(val).strip()
    try:
        f = float(s); i = int(f)
        if 1 <= i <= 31:
            return i
    except:
        pass
    m = re.match(r"^\s*(\d{1,2})\b", s)
    if m:
        i = int(m.group(1))
        if 1 <= i <= 31:
            return i
    return None

def detect_day_columns_from_headers(df):
    return [c for c in df.columns if coerce_day_label(c) is not None]

def promote_day_header_if_needed(df, look_first_rows=6):
    if detect_day_columns_from_headers(df):
        return df
    best_row, best_count = None, 0
    for r in range(min(look_first_rows, len(df))):
        count = sum(1 for v in df.iloc[r].tolist() if coerce_day_label(v) is not None)
        if count > best_count:
            best_row, best_count = r, count
    if best_row is not None and best_count >= 5:
        header_vals = df.iloc[best_row].astype(str).tolist()
        new_df = df.iloc[best_row+1:].copy()
        new_df.columns = [str(c).strip() for c in header_vals]
        return new_df
    return df

# Month parser for sheet names like "SEP 2025", "SEPTEMBER-2025", "2025-09"
MONTH_MAP = {"JAN":1,"JANUARY":1,"FEB":2,"FEBRUARY":2,"MAR":3,"MARCH":3,"APR":4,"APRIL":4,"MAY":5,"JUN":6,"JUNE":6,"JUL":7,"JULY":7,"AUG":8,"AUGUST":8,"SEP":9,"SEPT":9,"SEPTEMBER":9,"OCT":10,"OCTOBER":10,"NOV":11,"NOVEMBER":11,"DEC":12,"DECEMBER":12}
def parse_month_year(label: str):
    s = str(label or "").strip().upper()
    month = next((v for k,v in MONTH_MAP.items() if k in s), None)
    m_year = re.search(r"(20\d{2}|19\d{2})", s)
    year = int(m_year.group(1)) if m_year else datetime.today().year
    if month is None:
        m = re.search(r"\b(\d{1,2})[/-](\d{4})\b", s)  # e.g., 09-2025
        if m:
            month = int(m.group(1)); year = int(m.group(2))
    if month is None:
        month = datetime.today().month
    return year, month

# --- Core builder (single sheet) now returns cost-aware daily_long ---
def build_ot_report_wide(df_att, month_label="", rate_col_name=None, default_rate=0.0, ot_threshold=8.0, ot_multiplier=1.25):
    df_att = promote_day_header_if_needed(df_att)
    df_att.columns = [str(c).strip() for c in df_att.columns]
    nm = norm_cols_map(df_att.columns)
    name_col = pick_col(nm, "employee name", "name", "emp name", "staff name") or "NAME"
    code_col = pick_col(nm, "employee code", "emp code", "code", "id") or "E CODE"

    day_cols = detect_day_columns_from_headers(df_att)
    if not day_cols:
        raise ValueError("Could not find day columns 1..31 in the attendance sheet.")

    # detect rate column if not given
    rate_col = rate_col_name
    if rate_col is None:
        for cand in RATE_CANDIDATES:
            col = pick_col(nm, cand)
            if col:
                rate_col = col
                break

    id_vars = [c for c in [name_col, code_col] if c in df_att.columns]
    keep_cols = id_vars + ([rate_col] if (rate_col and rate_col in df_att.columns) else [])

    long = df_att.melt(id_vars=id_vars, value_vars=day_cols, var_name="DayLabel", value_name="CellRaw")

    # bring rate to long (if present)
    if rate_col and rate_col in df_att.columns:
        long = long.merge(dfp_att := df_att[keep_cols].drop_duplicates(), on=id_vars, how="left")
        long.rename(columns={rate_col: "Rate"}, inplace=True)
    else:
        long["Rate"] = default_rate

    long["Day"] = long["DayLabel"].map(coerce_day_label)
    long.dropna(subset=["Day"], inplace=True)
    long["Day"] = long["Day"].astype(int)

    long["Is_Absent"] = long["CellRaw"].map(is_absent_cell)
    long["Is_Present_Tok"] = long["CellRaw"].map(is_present_token)
    long["Hours"] = long["CellRaw"].map(parse_hours_cell)

    # Fill P=8 if Hours empty/0
    mask_fill_8h = long["Is_Present_Tok"] & (long["Hours"].isna() | (long["Hours"] == 0))
    long.loc[mask_fill_8h, "Hours"] = 8.0

    # Base/OT hours
    long["OT_Hours"] = (long["Hours"].fillna(0) - ot_threshold).clip(lower=0)
    long["Worked_Flag"] = ((~long["Is_Absent"]) & (long["Hours"].fillna(0) > 0)) | long["Is_Present_Tok"]

    # Costing (rate column is optional)
    long["Base_Hours"] = long["Hours"].fillna(0).clip(upper=ot_threshold)
    long["Base_Cost"] = long["Base_Hours"] * long["Rate"]
    long["OT_Cost"] = long["OT_Hours"] * long["Rate"] * ot_multiplier
    long["Total_Cost"] = long["Base_Cost"] + long["OT_Cost"]

    if month_label:
        long["Month"] = month_label

    # Real Date from sheet month + Day
    yy, mm = parse_month_year(month_label or "")
    try:
        long["Date"] = pd.to_datetime(dict(year=yy, month=mm, day=long["Day"]), errors="coerce")
    except Exception:
        long["Date"] = pd.NaT

    group_cols = [c for c in [code_col, name_col] if c in long.columns]
    summary = (
        long.groupby(group_cols, dropna=False)
            .agg(
                Absent_Days=("Is_Absent", lambda s: int(s.sum())),
                Days_With_OT=("OT_Hours", lambda s: int((s>0).sum())),
                Total_OT_Hours=("OT_Hours", "sum"),
                Total_Work_Hours=("Hours", "sum"),
                Base_Cost=("Base_Cost", "sum"),
                OT_Cost=("OT_Cost", "sum"),
                Total_Cost=("Total_Cost", "sum"),
                Days_Recorded=("Hours", lambda s: int(s.notna().sum())),
                Avg_Rate=("Rate", "mean"),
            )
            .reset_index()
    )
    return summary, long, name_col, code_col

def monthly_totals(summary_raw: pd.DataFrame, daily_long: pd.DataFrame, month_label: str):
    if summary_raw.empty:
        return pd.DataFrame([{
            "Month": month_label, "Employees": 0, "Total_Absent_Days": 0,
            "Days_With_OT": 0, "Total_OT_Hours": 0.0, "Total_Work_Hours": 0.0,
            "Base_Cost": 0.0, "OT_Cost": 0.0, "Total_Cost": 0.0,
            "Days_Recorded": 0, "Avg_Work_Hours_per_Day": 0.0
        }])
    employees         = int(len(summary_raw))
    total_absent      = int(summary_raw["Absent_Days"].sum())
    days_with_ot      = int(summary_raw["Days_With_OT"].sum())
    total_ot_hours    = float(summary_raw["Total_OT_Hours"].sum())
    total_work_hours  = float(summary_raw["Total_Work_Hours"].sum())
    base_cost         = float(summary_raw["Base_Cost"].sum())
    ot_cost           = float(summary_raw["OT_Cost"].sum())
    total_cost        = float(summary_raw["Total_Cost"].sum())
    days_recorded     = int(summary_raw["Days_Recorded"].sum())
    avg_work_per_day  = float(daily_long["Hours"].mean()) if not daily_long.empty else 0.0
    return pd.DataFrame([{
        "Month": month_label,
        "Employees": employees,
        "Total_Absent_Days": total_absent,
        "Days_With_OT": days_with_ot,
        "Total_OT_Hours": round(total_ot_hours, 2),
        "Total_Work_Hours": round(total_work_hours, 2),
        "Base_Cost": round(base_cost, 2),
        "OT_Cost": round(ot_cost, 2),
        "Total_Cost": round(total_cost, 2),
        "Days_Recorded": days_recorded,
        "Avg_Work_Hours_per_Day": round(avg_work_per_day, 2),
    }])

if att_file:
    try:
        xls2 = pd.ExcelFile(att_file)
        att_sheet = xls2.sheet_names[0]          # SHEET = month label
        df_att = pd.read_excel(xls2, sheet_name=att_sheet, header=0)
        st.success(f"Attendance loaded from sheet '{att_sheet}' with {len(df_att)} rows.")

        summary_raw, daily_long, name_col, code_col = build_ot_report_wide(
            df_att, month_label=att_sheet,
            default_rate=default_hourly_rate, ot_threshold=ot_threshold, ot_multiplier=default_ot_multiplier
        )

        month_totals_df = monthly_totals(summary_raw, daily_long, att_sheet)
        st.write("**Monthly Totals (single file)**")
        st.dataframe(month_totals_df, use_container_width=True)

        # Downloads
        @st.cache_data
        def to_csv_bytes(df): return df.to_csv(index=False).encode("utf-8")
        @st.cache_data
        def to_xlsx_bytes(df, sheet="Sheet1"):
            bio = io.BytesIO()
            with pd.ExcelWriter(bio, engine="openpyxl") as writer:
                df.to_excel(writer, index=False, sheet_name=sheet)
            bio.seek(0)
            return bio.read()

        c1, c2, c3 = st.columns(3)
        c1.download_button("⬇️ Download OT Summary (CSV)", data=to_csv_bytes(summary_raw),
                           file_name=f"OT_Summary_{att_sheet}.csv", mime="text/csv")
        c2.download_button("⬇️ Download Monthly Totals (CSV)", data=to_csv_bytes(month_totals_df),
                           file_name=f"OT_Monthly_Totals_{att_sheet}.csv", mime="text/csv")
        c3.download_button("⬇️ Download Daily Long (CSV)", data=to_csv_bytes(daily_long),
                           file_name=f"OT_Daily_{att_sheet}.csv", mime="text/csv")

        with st.expander("Show daily long data (one row per employee-day)", expanded=False):
            st.dataframe(daily_long, use_container_width=True)

    except Exception as e:
        st.error(f"Failed to build OT report: {e}")
        st.code(traceback.format_exc(), language="python")

# ==========================================================
#   MULTI-PROJECT TIMESHEETS (CONSOLIDATE & DAILY COSTING)
#   • Month = SHEET NAME
#   • Real calendar Date = (sheet month + Day)
#   • OT rate = BASIC/30/8 × OT_MULTIPLIER
#   • Base daily pay = Salary/Day else Gross/30
# ==========================================================
st.markdown("---")
st.subheader("Multi-Project Timesheets — Daily Costing Dashboard")

multi_files = st.file_uploader(
    "Upload month’s project timesheets (select multiple .xlsx)",
    type=["xlsx"], accept_multiple_files=True, key="multi_att"
)

PAY_BASIC_CANDS  = ["basic", "basic salary", "basic pay"]
PAY_GROSS_CANDS  = ["gross salary", "gross", "total salary"]
PAY_SALDAY_CANDS = ["salary/day", "salary per day", "per day salary", "day salary"]

if multi_files:
    all_daily = []
    all_emp_summaries = []

    for f in multi_files:
        try:
            # Project name from file; MONTH from SHEET NAME
            fname = f.name
            project_from_file = re.sub(r"\.xlsx$", "", fname, flags=re.I)

            xls = pd.ExcelFile(f)
            sheet = xls.sheet_names[0]                      # SHEET = month label
            dfp = pd.read_excel(xls, sheet_name=sheet, header=0)

            # ----- normalize wide sheet and identify columns -----
            dfp2 = promote_day_header_if_needed(dfp)
            dfp2.columns = [str(c).strip() for c in dfp2.columns]
            nm = norm_cols_map(dfp2.columns)
            name_col = pick_col(nm, "employee name", "name", "emp name", "staff name") or "NAME"
            code_col = pick_col(nm, "employee code", "emp code", "code", "id") or "E CODE"
            day_cols = detect_day_columns_from_headers(dfp2)
            if not day_cols:
                raise ValueError("Could not find day columns 1..31 in a sheet.")

            basic_col  = pick_col(nm, *PAY_BASIC_CANDS)
            gross_col  = pick_col(nm, *PAY_GROSS_CANDS)
            salday_col = pick_col(nm, *PAY_SALDAY_CANDS)

            id_vars = [c for c in [name_col, code_col] if c in dfp2.columns]
            long = dfp2.melt(id_vars=id_vars, value_vars=day_cols, var_name="DayLabel", value_name="CellRaw")

            # Attach pay columns
            carry = id_vars.copy()
            if basic_col: carry.append(basic_col)
            if gross_col: carry.append(gross_col)
            if salday_col: carry.append(salday_col)
            long = long.merge(dfp2[carry].drop_duplicates(), on=id_vars, how="left")

            # Standardize identifiers
            long["Employee Name"] = long[name_col] if name_col in long.columns else ""
            long["Employee Code"] = long[code_col] if code_col in long.columns else ""

            # Day / Hours
            long["Day"] = long["DayLabel"].map(coerce_day_label)
            long.dropna(subset=["Day"], inplace=True)
            long["Day"] = long["Day"].astype(int)
            long["Is_Absent"] = long["CellRaw"].map(is_absent_cell)
            long["Is_Present_Tok"] = long["CellRaw"].map(is_present_token)
            long["Hours"] = long["CellRaw"].map(parse_hours_cell)

            # Treat bare 'P' as 8 hours if Hours empty/0
            mask_fill_8h = long["Is_Present_Tok"] & (long["Hours"].isna() | (long["Hours"] == 0))
            long.loc[mask_fill_8h, "Hours"] = 8.0

            # Insert project first, then de-dup per Project+Employee+Day
            long.insert(0, "Project", project_from_file)
            long = long.sort_values(["Project","Employee Code","Day","Hours"], ascending=[True,True,True,False])
            long = long.drop_duplicates(subset=["Project","Employee Code","Day"], keep="first")

            # Rates
            long["Basic"] = pd.to_numeric(long[basic_col], errors="coerce").fillna(0.0) if basic_col in long.columns else 0.0
            if salday_col and salday_col in long.columns:
                long["Salary_Day"] = pd.to_numeric(long[salday_col], errors="coerce")
            else:
                long["Salary_Day"] = (pd.to_numeric(long[gross_col], errors="coerce")/30.0) if gross_col in long.columns else 0.0
            long["OT_Rate"] = (long["Basic"]/30.0/8.0) * default_ot_multiplier

            # Daily split & costing
            long["OT_Hours"] = (long["Hours"].fillna(0) - ot_threshold).clip(lower=0)
            long["Worked_Flag"] = ((~long["Is_Absent"]) & (long["Hours"].fillna(0) > 0)) | long["Is_Present_Tok"]
            long["Base_Daily_Cost"] = long["Salary_Day"].where(long["Worked_Flag"], other=0.0)
            long["OT_Cost"] = long["OT_Hours"] * long["OT_Rate"]
            long["Total_Daily_Cost"] = long["Base_Daily_Cost"] + long["OT_Cost"]

            # Month & real calendar Date (from sheet name)
            long["Month"] = sheet
            yy, mm = parse_month_year(sheet)
            try:
                long["Date"] = pd.to_datetime(dict(year=yy, month=mm, day=long["Day"]), errors="coerce")
            except Exception:
                long["Date"] = pd.NaT

            # Per-employee rollup (per file)
            emp_sum = (
                long.groupby(["Project","Employee Code","Employee Name"], dropna=False)
                    .agg(
                        Present_Days=("Worked_Flag", lambda s: int(s.sum())),
                        Absent_Days=("Is_Absent", lambda s: int(s.sum())),
                        OT_Days=("OT_Hours", lambda s: int((s>0).sum())),
                        OT_Hours=("OT_Hours","sum"),
                        Total_Hours=("Hours","sum"),
                        Base_Cost=("Base_Daily_Cost","sum"),
                        OT_Cost=("OT_Cost","sum"),
                        Total_Cost=("Total_Daily_Cost","sum"),
                    ).reset_index()
            )

            all_daily.append(long)
            all_emp_summaries.append(emp_sum)

        except Exception as e:
            st.error(f"Failed to parse {f.name}: {e}")
            st.code(traceback.format_exc(), language="python")

    if all_daily:
        proj_daily = pd.concat(all_daily, ignore_index=True)
        proj_summary = pd.concat(all_emp_summaries, ignore_index=True)

        # ---------------- Dashboard Filters ----------------
        st.markdown("### Filters")
        projects = sorted(proj_daily["Project"].dropna().unique().tolist())
        sel_projects = st.multiselect("Projects", projects, default=projects)

        # Calendar From → To (fallback to day slider if Date missing)
        if "Date" in proj_daily.columns and not proj_daily["Date"].isna().all():
            min_date = pd.to_datetime(proj_daily["Date"].min()).date()
            max_date = pd.to_datetime(proj_daily["Date"].max()).date()
            date_from, date_to = st.date_input(
                "Date range", value=(min_date, max_date),
                min_value=min_date, max_value=max_date
            )
            mask = proj_daily["Project"].isin(sel_projects) & proj_daily["Date"].between(
                pd.to_datetime(date_from), pd.to_datetime(date_to)
            )
        else:
            min_day, max_day = int(proj_daily["Day"].min()), int(proj_daily["Day"].max())
            day_from, day_to = st.slider("Day range", min_value=min_day, max_value=max_day, value=(min_day, max_day))
            mask = proj_daily["Project"].isin(sel_projects) & (proj_daily["Day"].between(int(day_from), int(day_to)))

        filt_daily = proj_daily.loc[mask].copy()

        # ---------- Ensure required columns exist to avoid KeyError ----------
        required_num_cols = ["Hours","OT_Hours","Base_Daily_Cost","OT_Cost","Total_Daily_Cost","Day"]
        for c in required_num_cols:
            if c not in filt_daily.columns:
                filt_daily[c] = 0.0
            filt_daily[c] = pd.to_numeric(filt_daily[c], errors="coerce").fillna(0.0)
        for c in ["Employee Name","Worked_Flag","Is_Absent"]:
            if c not in filt_daily.columns:
                filt_daily[c] = "" if c == "Employee Name" else False

        # ---------------- Attendance Summary (present/absent per employee & project) ----------------
        attendance_summary = (
            filt_daily.groupby(["Project","Employee Code","Employee Name"], dropna=False)
                .agg(
                    Present_Days=("Worked_Flag", lambda s: int(s.sum())),
                    Absent_Days=("Is_Absent", lambda s: int(s.sum())),
                    Total_Hours=("Hours","sum"),
                    OT_Days=("OT_Hours", lambda s: int((s>0).sum())),
                    OT_Hours=("OT_Hours","sum"),
                    Base_Cost=("Base_Daily_Cost","sum"),
                    OT_Cost=("OT_Cost","sum"),
                    Total_Cost=("Total_Daily_Cost","sum"),
                ).reset_index().sort_values(["Project","Employee Name"])
        )

        # ---------------- Project × Day totals ----------------
        if len(filt_daily) == 0:
            st.info("No rows for the current filters.")
            by_proj_day = pd.DataFrame(columns=["Project","Day","Employees","Hours","OT_Hours","Base_Cost","OT_Cost","Total_Cost"])
            emp_daily = pd.DataFrame(columns=["Project","Employee Code","Employee Name","Day","Hours","OT_Hours","Salary_Day","OT_Rate","Base_Daily_Cost","OT_Cost","Total_Daily_Cost"])
            project_totals = pd.DataFrame(columns=["Project","Employees","Total_Work_Hours","Total_OT_Hours","Base_Cost","OT_Cost","Total_Cost"])
        else:
            by_proj_day = (
                filt_daily.groupby(["Project","Day"], dropna=False)
                    .agg(
                        Employees=("Employee Name", lambda s: s.nunique()),
                        Hours=("Hours","sum"),
                        OT_Hours=("OT_Hours","sum"),
                        Base_Cost=("Base_Daily_Cost","sum"),
                        OT_Cost=("OT_Cost","sum"),
                        Total_Cost=("Total_Daily_Cost","sum"),
                    ).reset_index().sort_values(["Project","Day"])
            )

            # ---------------- Employee daily detail ----------------
            emp_daily = (
                filt_daily[[
                    "Project","Employee Code","Employee Name","Day","Hours","OT_Hours","Salary_Day","OT_Rate","Base_Daily_Cost","OT_Cost","Total_Daily_Cost"
                ]].sort_values(["Project","Employee Name","Day"])
            )

            # ---------------- Project-wise totals (filtered) ----------------
            project_totals = (
                filt_daily.groupby(["Project"], dropna=False)
                    .agg(
                        Employees=("Employee Name", lambda s: s.nunique()),
                        Total_Work_Hours=("Hours", "sum"),
                        Total_OT_Hours=("OT_Hours", "sum"),
                        Base_Cost=("Base_Daily_Cost", "sum"),
                        OT_Cost=("OT_Cost", "sum"),
                        Total_Cost=("Total_Daily_Cost", "sum"),
                    ).reset_index().sort_values("Project")
            )

        st.markdown("#### Attendance Summary (Filtered)")
        st.dataframe(attendance_summary, use_container_width=True)

        st.markdown("#### Project × Day — Daily Costing")
        st.dataframe(by_proj_day, use_container_width=True)

        st.markdown("#### Employee Daily Detail")
        st.dataframe(emp_daily, use_container_width=True)

        st.markdown("#### Project-wise Totals (Filtered)")
        st.dataframe(project_totals, use_container_width=True)

        # ---------------- Downloads ----------------
        @st.cache_data
        def to_csv_bytes(df):
            return df.to_csv(index=False).encode("utf-8")

        @st.cache_data
        def to_xlsx_bytes(dfs: dict):
            bio = io.BytesIO()
            with pd.ExcelWriter(bio, engine="openpyxl") as writer:
                for sheet, df in dfs.items():
                    df.to_excel(writer, index=False, sheet_name=sheet[:31])
            bio.seek(0)
            return bio.read()

        c1, c2, c3, c4, c5 = st.columns(5)
        c1.download_button("⬇️ Project × Day (CSV)", data=to_csv_bytes(by_proj_day), file_name="Project_Day_Costs.csv")
        c2.download_button("⬇️ Employee Daily (CSV)", data=to_csv_bytes(emp_daily), file_name="Employee_Daily_Costs.csv")
        c3.download_button("⬇️ Project Totals (CSV)", data=to_csv_bytes(project_totals), file_name="Project_Totals_Filtered.csv")
        c4.download_button("⬇️ Attendance Summary (CSV)", data=to_csv_bytes(attendance_summary), file_name="Attendance_Summary_Filtered.csv")
        c5.download_button(
            "⬇️ Excel Pack (All Tabs)",
            data=to_xlsx_bytes({
                "Project_Day_Costs": by_proj_day,
                "Employee_Daily": emp_daily,
                "Project_Totals": project_totals,
                "Attendance_Summary": attendance_summary,
                "Daily_Long_All": filt_daily,
            }),
            file_name="Timesheet_DailyCost_Filtered.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

# --- UI footer: Powered By Jaseer (NOT in PDFs) ---
st.markdown(
    f"""
    <style>
      .stApp {{ padding-bottom: 60px; }}
      .custom-footer {{
        position: fixed; left: 0; right: 0; bottom: 0;
        text-align: center; padding: 10px 0;
        color: #6b7280; font-size: 12px; background: rgba(255,255,255,0.7);
      }}
    </style>
    <div class="custom-footer">{UI_POWERED_BY_TEXT}</div>
    """,
    unsafe_allow_html=True,
)
