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

EARNINGS_LETTERS = {
    "Basic Pay": "F",
    "Other Allowance": "G",
    "Housing Allowance": "H",
    "Repay Extra Leave Punishment / Other Ded": "AC",
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
        if re.fullmatch(r"\d{1,2}:\d{2}(:\d{2})?", s):
            h, m2, *rest = s.split(":")
            sec = int(rest[0]) if rest else 0
            return int(h) + int(m2)/60 + sec/3600
        return None

def fmt_amount(x, decimals=2):
    v = float(x) if isinstance(x, (int, float)) else parse_number(x)
    if v is None or pd.isna(v): v = 0.0
    return f"{v:,.{decimals}f}"

def amount_in_words(x):
    v = float(x) if isinstance(x, (int, float)) else parse_number(x)
    if v is None or pd.isna(v): return ""
    sign = "minus " if v < 0 else ""
    v = abs(v); whole = int(v); frac = int(round((v - whole) * 100))
    words = num2words(whole, lang="en").replace("-", " ")
    if frac: words = f"{words} and {frac:02d}/100"
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

def letter_value(row_values, letter, max_cols=None):
    try:
        idx = column_index_from_string(letter) - 1
        if max_cols is not None and idx >= max_cols:
            return None
        return row_values[idx]
    except Exception:
        return None

def sum_letters(row_values, letters, max_cols=None):
    total = 0.0; used = False
    for L in letters:
        v = parse_number(letter_value(row_values, L, max_cols))
        if v is not None:
            total += v; used = True
    return total if used else 0.0

def get_absent_days(row, norm_map):
    v = get_value(row, norm_map, ABSENT_DAYS_CANDIDATES)
    num = parse_number(v)
    return num if num is not None else 0

# ========================== PAYSLIP PDF ==========================
def build_pdf_for_row(std, company_name, title, page_size,
                      currency_label="AED", include_words=True, logo_bytes=None, logo_width=1.2*inch) -> bytes:
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=page_size,
                            leftMargin=0.8*inch, rightMargin=0.8*inch,
                            topMargin=0.6*inch, bottomMargin=0.6*inch)
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
        ("FONTNAME",(0,0),(0,0),"Helvetica-Bold"), ("FONTNAME",(1,0),(1,0),"Helvetica-Bold"),
    ]))
    elems += [hdr_tbl, Spacer(1,6)]

    earn_rows = [[f"Earnings ({currency_label})","Amount"]]
    for lbl in ["Basic Pay","Other Allowance","Housing Allowance","Repay Extra Leave Punishment / Other Ded","Over time","Reward for Full Day Attendance","Incentive"]:
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

excel_file = st.file_uploader("Upload Payroll Excel (.xlsx)", type=["xlsx"], key="excel_payroll_file")

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
        st.success(f"Loaded {len(df)} rows from sheet '{sheet_name}' for Payslips. Pay Period will be '{sheet_name}'.")
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
                    fname = (" ".join([p for p in parts if p]) or f"Payslip_{i+1}") + ".pdf"
                    zf.writestr(fname, pdf_bytes)
                except Exception as e:
                    tb = traceback.format_exc()
                    zf.writestr(f"row_{i+1}_ERROR.txt", f"Row {i+1}: {e}\n\n{tb}")

        zbuf.seek(0)
        run_id = datetime.now().strftime("%Y%m%d-%H%M%S")
        st.download_button("⬇️ Download Payslips (ZIP)", data=zbuf.read(),
                           file_name=f"Payslips_{sheet_name}_{run_id}.zip", mime="application/zip")

# ==========================================================
#   MULTI-PROJECT TIMESHEETS (ALL WORKSHEETS)
# ==========================================================
st.markdown("---")
st.subheader("Multi-Project Timesheets — Daily Costing Dashboard")

normalize_30_days = st.checkbox(
    "Normalize attendance to 30 days (pad missing days as OFF)",
    value=True,
    help="If a month has <30 day columns (e.g., February), synthesize extra OFF days so summaries treat the month as 30 days."
)

multi_files = st.file_uploader(
    "Upload project timesheets (select multiple .xlsx; reads ALL worksheets in each)",
    type=["xlsx"], accept_multiple_files=True, key="multi_att_all_sheets"
)

PAY_BASIC_CANDS  = ["basic", "basic salary", "basic pay"]
PAY_GROSS_CANDS  = ["gross salary", "gross", "total salary"]
PAY_SALDAY_CANDS = ["salary/day", "salary per day", "per day salary", "day salary"]

ABSENT_TOKENS = {"a", "absent"}
PRESENT_TOKENS = {"p", "present"}
OFF_TOKENS = {"off","leave","holiday"}

def norm_cols_map(columns):
    return {str(c).strip().lower(): c for c in columns}

def pick_col(norm_map, *candidates):
    for cand in candidates:
        if cand in norm_map: return norm_map[cand]
    for cand in candidates:
        for k, v in norm_map.items():
            if cand in k: return v
    return None

def is_absent_cell(x):
    if x is None or (isinstance(x, float) and pd.isna(x)): return False
    return str(x).strip().lower() in ABSENT_TOKENS

def is_present_token(x):
    if x is None or (isinstance(x, float) and pd.isna(x)): return False
    return str(x).strip().lower() in PRESENT_TOKENS

def is_off_cell(x):
    if x is None or (isinstance(x, float) and pd.isna(x)): return False
    return str(x).strip().lower() in OFF_TOKENS

def parse_hours_cell(x):
    if x is None or (isinstance(x, float) and pd.isna(x)): return None
    s = str(x).strip().lower()
    if s in OFF_TOKENS or s in {"-","--",""}: return None
    if s in ABSENT_TOKENS: return 0.0
    if s in PRESENT_TOKENS: return None
    s_orig = str(x).strip()
    if re.fullmatch(r"\d{1,2}:\d{2}(:\d{2})?", s_orig):
        h, m, *rest = s_orig.split(":"); sec = int(rest[0]) if rest else 0
        return int(h) + int(m)/60 + sec/3600
    m = re.search(r"(\d+(\.\d+)?)\s*h", s_orig, re.I)
    if m:
        hours = float(m.group(1))
        m2 = re.search(r"(\d+)\s*m", s_orig, re.I)
        if m2: hours += int(m2.group(1))/60
        return hours
    try:
        return float(s_orig)
    except:
        return None

def coerce_day_label(val):
    if pd.isna(val): return None
    s = str(val).strip()
    try:
        f = float(s); i = int(f)
        if 1 <= i <= 31: return i
    except: pass
    m = re.match(r"^\s*(\d{1,2})\b", s)
    if m:
        i = int(m.group(1))
        if 1 <= i <= 31: return i
    return None

def detect_day_columns_from_headers(df):
    return [c for c in df.columns if coerce_day_label(c) is not None]

MONTH_MAP = {"JAN":1,"JANUARY":1,"FEB":2,"FEBRUARY":2,"MAR":3,"MARCH":3,"APR":4,"APRIL":4,"MAY":5,
             "JUN":6,"JUNE":6,"JUL":7,"JULY":7,"AUG":8,"AUGUST":8,"SEP":9,"SEPT":9,"SEPTEMBER":9,
             "OCT":10,"OCTOBER":10,"NOV":11,"NOVEMBER":11,"DEC":12,"DECEMBER":12}

def parse_month_year(label: str):
    s = str(label or "").strip().upper()
    month = next((v for k,v in MONTH_MAP.items() if k in s), None)
    m_year = re.search(r"(20\d{2}|19\d{2})", s)
    year = int(m_year.group(1)) if m_year else datetime.today().year
    if month is None:
        m = re.search(r"\b(\d{1,2})[/-](\d{4})\b", s)
        if m: month = int(m.group(1)); year = int(m.group(2))
    if month is None: month = datetime.today().month
    return year, month

def promote_day_header_if_needed(df, look_first_rows=6):
    if detect_day_columns_from_headers(df): return df
    best_row, best_count = None, 0
    for r in range(min(look_first_rows, len(df))):
        count = sum(1 for v in df.iloc[r].tolist() if coerce_day_label(v) is not None)
        if count > best_count: best_row, best_count = r, count
    if best_row is not None and best_count >= 5:
        header_vals = df.iloc[best_row].astype(str).tolist()
        new_df = df.iloc[best_row+1:].copy()
        new_df.columns = [str(c).strip() for c in header_vals]
        return new_df
    return df

def fmt_commas(df, money_cols=None, hour_cols=None, int_cols=None, decimals_money=2, decimals_hours=2):
    if df is None or len(df) == 0: return df
    money_cols = [c for c in (money_cols or []) if c in df.columns]
    hour_cols  = [c for c in (hour_cols  or []) if c in df.columns]
    int_cols   = [c for c in (int_cols   or []) if c in df.columns]
    fmts = {}
    fmts.update({c: f"{{:,.{decimals_money}f}}" for c in money_cols})
    fmts.update({c: f"{{:,.{decimals_hours}f}}" for c in hour_cols})
    fmts.update({c: "{:,.0f}" for c in int_cols})
    return df.style.format(fmts)

def tidy_one_sheet(df_sheet: pd.DataFrame, sheet_name: str, project_from_file: str,
                   default_ot_multiplier: float, normalize_30_days: bool = False):
    dfp2 = promote_day_header_if_needed(df_sheet)
    dfp2.columns = [str(c).strip() for c in dfp2.columns]
    nm = norm_cols_map(dfp2.columns)
    name_col = pick_col(nm, "employee name", "name", "emp name", "staff name") or "NAME"
    code_col = pick_col(nm, "employee code", "emp code", "code", "id") or "E CODE"
    day_cols = detect_day_columns_from_headers(dfp2)
    if not day_cols:
        raise ValueError("Could not find day columns 1..31 in a sheet.")

    basic_col  = pick_col(nm, "basic", "basic salary", "basic pay")
    gross_col  = pick_col(nm, "gross salary", "gross", "total salary")
    salday_col = pick_col(nm, "salary/day", "salary per day", "per day salary", "day salary")

    id_vars = [c for c in [name_col, code_col] if c in dfp2.columns]
    long = dfp2.melt(id_vars=id_vars, value_vars=day_cols, var_name="DayLabel", value_name="CellRaw")

    carry = id_vars.copy()
    if basic_col: carry.append(basic_col)
    if gross_col: carry.append(gross_col)
    if salday_col: carry.append(salday_col)
    long = long.merge(dfp2[carry].drop_duplicates(), on=id_vars, how="left")

    long["Employee Name"] = long[name_col] if name_col in long.columns else ""
    long["Employee Code"] = long[code_col] if code_col in long.columns else ""

    long["Day"] = long["DayLabel"].map(coerce_day_label)
    long.dropna(subset=["Day"], inplace=True)
    long["Day"] = long["Day"].astype(int)

    long["Is_Absent"] = long["CellRaw"].map(is_absent_cell)
    long["Is_Present_Tok"] = long["CellRaw"].map(is_present_token)
    long["Is_Off"] = long["CellRaw"].map(is_off_cell)
    long["Hours"] = long["CellRaw"].map(parse_hours_cell)

    long.insert(0, "Project", project_from_file)
    long = long.sort_values(["Project","Employee Code","Day","Hours"], ascending=[True,True,True,False])
    long = long.drop_duplicates(subset=["Project","Employee Code","Day"], keep="first")

    long["Basic"] = pd.to_numeric(long[basic_col], errors="coerce").fillna(0.0) if basic_col in long.columns else 0.0
    if salday_col and salday_col in long.columns:
        long["Salary_Day"] = pd.to_numeric(long[salday_col], errors="coerce")
    else:
        long["Salary_Day"] = (pd.to_numeric(long[gross_col], errors="coerce")/30.0) if gross_col in long.columns else 0.0
    long["OT_Rate"] = (long["Basic"]/30.0/8.0) * default_ot_multiplier

    ot_threshold = 8.0
    long["OT_Hours"] = (long["Hours"].fillna(0) - ot_threshold).clip(lower=0)

    long["Worked_Flag"] = (long["Is_Present_Tok"] | (long["Hours"].fillna(0) > 0)) & (~long["Is_Absent"]) & (~long["Is_Off"])

    long["Base_Daily_Cost"] = long["Salary_Day"].where(long["Worked_Flag"], other=0.0)
    long["OT_Cost"] = long["OT_Hours"] * long["OT_Rate"]
    long["Total_Daily_Cost"] = long["Base_Daily_Cost"] + long["OT_Cost"]

    long["Month"] = sheet_name
    yy, mm = parse_month_year(sheet_name)
    try:
        long["Date"] = pd.to_datetime(dict(year=yy, month=mm, day=long["Day"]), errors="coerce")
    except Exception:
        long["Date"] = pd.NaT

    # ===== normalize to 30 attendance days (pad OFF) =====
    if normalize_30_days:
        base_keys = long[["Project","Employee Code","Employee Name"]].drop_duplicates()
        all_days = pd.DataFrame({"Day": list(range(1, 31))})
        base_keys["_tmp"] = 1
        all_days["_tmp"] = 1
        wanted = base_keys.merge(all_days, on="_tmp").drop(columns="_tmp")

        have = long[["Project","Employee Code","Day"]].drop_duplicates()
        missing = wanted.merge(have, on=["Project","Employee Code","Day"], how="left", indicator=True)
        missing = missing.loc[missing["_merge"] == "left_only"].drop(columns="_merge")

        if not missing.empty:
            synth = missing.copy()
            synth["CellRaw"] = "OFF"
            synth["Is_Absent"] = False
            synth["Is_Present_Tok"] = False
            synth["Is_Off"] = True
            synth["Hours"] = 0.0
            synth["OT_Hours"] = 0.0
            synth["Worked_Flag"] = False

            carry_cols = ["Salary_Day","Basic","OT_Rate"]
            carry_src = (
                long[["Project","Employee Code"] + carry_cols]
                .drop_duplicates(subset=["Project","Employee Code"])
            )
            synth = synth.merge(carry_src, on=["Project","Employee Code"], how="left")
            for c in carry_cols:
                if c not in synth.columns:
                    synth[c] = 0.0
                synth[c] = pd.to_numeric(synth[c], errors="coerce").fillna(0.0)

            synth["Base_Daily_Cost"] = 0.0
            synth["OT_Cost"] = 0.0
            synth["Total_Daily_Cost"] = 0.0

            try:
                synth["Date"] = pd.to_datetime(dict(year=yy, month=mm, day=synth["Day"]), errors="coerce")
            except Exception:
                synth["Date"] = pd.NaT

            keep_cols = list(long.columns)
            for c in keep_cols:
                if c not in synth.columns:
                    synth[c] = pd.NA

            long = pd.concat([long, synth[keep_cols]], ignore_index=True)

    return long

if multi_files:
    all_daily = []
    for f in multi_files:
        try:
            project_from_file = re.sub(r"\.xlsx$", "", f.name, flags=re.I)
            xls = pd.ExcelFile(f)
            for sheet in xls.sheet_names:
                try:
                    df_sheet = pd.read_excel(xls, sheet_name=sheet, header=0)
                    if df_sheet.shape[0] == 0 or df_sheet.shape[1] == 0:
                        continue
                    long = tidy_one_sheet(df_sheet, sheet, project_from_file, default_ot_multiplier, normalize_30_days)
                    all_daily.append(long)
                except Exception as inner_e:
                    st.warning(f"[{project_from_file}] Skipped sheet '{sheet}': {inner_e}")
        except Exception as e:
            st.error(f"Failed to parse file {f.name}: {e}")
            st.code(traceback.format_exc(), language="python")

    if all_daily:
        proj_daily = pd.concat(all_daily, ignore_index=True)

        # Filters
        st.markdown("### Filters")
        projects = sorted(proj_daily["Project"].dropna().unique().tolist())
        sel_projects = st.multiselect("Projects", projects, default=projects)

        if "Date" in proj_daily.columns and not proj_daily["Date"].isna().all():
            min_date = pd.to_datetime(proj_daily["Date"].min()).date()
            max_date = pd.to_datetime(proj_daily["Date"].max()).date()
            c_from, c_to = st.columns(2)
            date_from = c_from.date_input("From date", value=min_date, min_value=min_date, max_value=max_date, key="from_date")
            date_to   = c_to.date_input("To date",   value=max_date, min_value=min_date, max_value=max_date, key="to_date")
            if date_from > date_to: date_from, date_to = date_to, date_from
            mask = proj_daily["Project"].isin(sel_projects) & proj_daily["Date"].between(
                pd.to_datetime(date_from), pd.to_datetime(date_to)
            )
        else:
            min_day, max_day = int(proj_daily["Day"].min()), int(proj_daily["Day"].max())
            day_from, day_to = st.slider("Day range", min_value=min_day, max_value=max_day, value=(min_day, max_day))
            mask = proj_daily["Project"].isin(sel_projects) & proj_daily["Day"].between(int(day_from), int(day_to))

        filt_daily = proj_daily.loc[mask].copy()

        # Remove employees with no work at all in the range
        if not filt_daily.empty:
            emp_presence = (
                filt_daily.groupby(["Project","Employee Code","Employee Name"], dropna=False)
                .agg(Total_Work_Hours=("Hours","sum"), Worked_Days=("Worked_Flag","sum"))
                .reset_index()
            )
            keep_keys = emp_presence.loc[
                (emp_presence["Total_Work_Hours"] > 0) | (emp_presence["Worked_Days"] > 0),
                ["Project","Employee Code"]
            ].drop_duplicates()
            if not keep_keys.empty:
                filt_daily = filt_daily.merge(keep_keys, on=["Project","Employee Code"], how="inner")
            else:
                filt_daily = filt_daily.iloc[0:0]

        # Ensure columns
        for c in ["Hours","OT_Hours","Base_Daily_Cost","OT_Cost","Total_Daily_Cost","Day"]:
            if c not in filt_daily.columns: filt_daily[c] = 0.0
            filt_daily[c] = pd.to_numeric(filt_daily[c], errors="coerce").fillna(0.0)
        for c in ["Employee Name","Worked_Flag","Is_Absent"]:
            if c not in filt_daily.columns:
                filt_daily[c] = "" if c == "Employee Name" else False

        # ===== Attendance summary =====
        grp = (
            filt_daily.groupby(["Project","Employee Code","Employee Name"], dropna=False)
            .agg(
                Present_Days=("Worked_Flag", lambda s: int(s.sum())),
                Absent_Marked=("Is_Absent",  lambda s: int(s.sum())),
                Total_Hours=("Hours","sum"),
                OT_Days=("OT_Hours", lambda s: int((s>0).sum())),
                OT_Hours=("OT_Hours","sum"),
                Base_Cost=("Base_Daily_Cost","sum"),
                OT_Cost=("OT_Cost","sum"),
                Total_Cost=("Total_Daily_Cost","sum"),
                Days_Seen=("Day", lambda s: int(pd.Series(s).nunique())),
            )
            .reset_index()
        )

        if normalize_30_days:
            expected_days = 30
            grp["Absent_Days"] = (grp["Absent_Marked"] +
                                  (expected_days - (grp["Present_Days"] + grp["Absent_Marked"])).clip(lower=0))
        else:
            grp["Absent_Days"] = grp["Absent_Marked"]

        attendance_summary = grp[[
            "Project","Employee Code","Employee Name",
            "Present_Days","Absent_Days",
            "Total_Hours","OT_Days","OT_Hours",
            "Base_Cost","OT_Cost","Total_Cost"
        ]].sort_values(["Project","Employee Name"])

        # Worked days only
        work_daily = filt_daily.loc[filt_daily["Worked_Flag"] == True].copy()

        if len(work_daily) == 0:
            st.info("No worked-day rows for the current filters.")
            by_proj_day = pd.DataFrame(columns=[
                "Project","Day","Employees","Hours","OT_Hours","Base_Cost","OT_Cost","Total_Cost",
                "Cum_Hours","Cum_OT_Hours","Cum_Base_Cost","Cum_OT_Cost","Cum_Total_Cost","Accumulated"
            ])
            emp_daily = pd.DataFrame(columns=[
                "Project","Employee Code","Employee Name","Month","Day","Hours","OT_Hours","Total_Hours",
                "Salary_Day","OT_Rate","Base_Daily_Cost","OT_Cost","Total_Daily_Cost","Accumulated"
            ])
            project_totals = pd.DataFrame(columns=["Project","Employees","Total_Work_Hours","Total_OT_Hours","Base_Cost","OT_Cost","Total_Cost"])
        else:
            by_proj_day = (
                work_daily.groupby(["Project","Day"], dropna=False)
                .agg(
                    Employees=("Employee Name", lambda s: s.nunique()),
                    Hours=("Hours","sum"),
                    OT_Hours=("OT_Hours","sum"),
                    Base_Cost=("Base_Daily_Cost","sum"),
                    OT_Cost=("OT_Cost","sum"),
                    Total_Cost=("Total_Daily_Cost","sum"),
                ).reset_index().sort_values(["Project","Day"])
            )
            by_proj_day["Cum_Hours"]      = by_proj_day.groupby("Project")["Hours"].cumsum()
            by_proj_day["Cum_OT_Hours"]   = by_proj_day.groupby("Project")["OT_Hours"].cumsum()
            by_proj_day["Cum_Base_Cost"]  = by_proj_day.groupby("Project")["Base_Cost"].cumsum()
            by_proj_day["Cum_OT_Cost"]    = by_proj_day.groupby("Project")["OT_Cost"].cumsum()
            by_proj_day["Cum_Total_Cost"] = by_proj_day.groupby("Project")["Total_Cost"].cumsum()
            by_proj_day["Accumulated"]    = by_proj_day["Cum_Total_Cost"]

            # Employee Daily (Total_Hours placed right after OT_Hours)
            emp_daily = (
                work_daily[[
                    "Project","Employee Code","Employee Name","Month","Day","Hours","OT_Hours",
                    "Salary_Day","OT_Rate","Base_Daily_Cost","OT_Cost","Total_Daily_Cost"
                ]].sort_values(["Project","Employee Code","Day"])
            )
            emp_daily["Total_Hours"] = emp_daily["Hours"].fillna(0) + emp_daily["OT_Hours"].fillna(0)
            emp_daily["Accumulated"] = emp_daily.groupby(["Project","Employee Code"])["Total_Daily_Cost"].cumsum()
            # Reorder columns so Total_Hours is right after OT_Hours
            emp_daily = emp_daily[[
                "Project","Employee Code","Employee Name","Month","Day",
                "Hours","OT_Hours","Total_Hours",
                "Salary_Day","OT_Rate","Base_Daily_Cost","OT_Cost","Total_Daily_Cost","Accumulated"
            ]]

            project_totals = (
                work_daily.groupby(["Project"], dropna=False)
                .agg(
                    Employees=("Employee Name", lambda s: s.nunique()),
                    Total_Work_Hours=("Hours", "sum"),
                    Total_OT_Hours=("OT_Hours", "sum"),
                    Base_Cost=("Base_Daily_Cost", "sum"),
                    OT_Cost=("OT_Cost", "sum"),
                    Total_Cost=("Total_Daily_Cost", "sum"),
                ).reset_index().sort_values("Project")
            )

        # Tables
        st.markdown("#### Attendance Summary (Filtered)")
        st.dataframe(
            fmt_commas(attendance_summary,
                       money_cols=["Base_Cost","OT_Cost","Total_Cost"],
                       hour_cols=["Total_Hours","OT_Hours"],
                       int_cols=["Present_Days","Absent_Days","OT_Days"]),
            use_container_width=True
        )

        st.markdown("#### Project × Day — Daily Costing (worked days only)")
        st.dataframe(
            fmt_commas(by_proj_day,
                       money_cols=["Base_Cost","OT_Cost","Total_Cost","Cum_Base_Cost","Cum_OT_Cost","Cum_Total_Cost","Accumulated"],
                       hour_cols=["Hours","OT_Hours","Cum_Hours","Cum_OT_Hours"],
                       int_cols=["Employees","Day"]),
            use_container_width=True
        )

        st.markdown("#### Employee Daily Detail (worked days only)")
        st.dataframe(
            fmt_commas(emp_daily,
                       money_cols=["Salary_Day","OT_Rate","Base_Daily_Cost","OT_Cost","Total_Daily_Cost","Accumulated"],
                       hour_cols=["Hours","OT_Hours","Total_Hours"],
                       int_cols=["Day"]),
            use_container_width=True
        )

        st.markdown("#### Project-wise Totals (worked days only)")
        st.dataframe(
            fmt_commas(project_totals,
                       money_cols=["Base_Cost","OT_Cost","Total_Cost"],
                       hour_cols=["Total_Work_Hours","Total_OT_Hours"],
                       int_cols=["Employees"]),
            use_container_width=True
        )

        # Downloads
        @st.cache_data
        def to_csv_bytes(df): return df.to_csv(index=False).encode("utf-8")

        if len(work_daily) == 0:
            by_proj_day_export = pd.DataFrame(columns=[
                "Project","Day","Employees","Hours","OT_Hours",
                "Base_Cost","OT_Cost","Total_Cost","Accumulated"
            ])
        else:
            by_proj_day_export = by_proj_day.copy()
            by_proj_day_export["Accumulated"] = by_proj_day_export.groupby("Project")["Total_Cost"].cumsum()
            by_proj_day_export = by_proj_day_export[
                ["Project","Day","Employees","Hours","OT_Hours","Base_Cost","OT_Cost","Total_Cost","Accumulated"]
            ]

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
                "Project_Day_Costs": by_proj_day_export,
                "Employee_Daily": emp_daily,  # Total_Hours after OT_Hours
                "Project_Totals": project_totals,
                "Attendance_Summary": attendance_summary,
                "Daily_Long_All": filt_daily.loc[filt_daily["Worked_Flag"] == True].copy(),
            }),
            file_name="Timesheet_DailyCost_Filtered.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

# --- UI footer ---
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

