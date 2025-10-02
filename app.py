import io
import re
import zipfile
import traceback
from datetime import datetime, timedelta

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

# ---- Fuzzy header candidates (case-insensitive) ----
EMP_NAME_CANDIDATES = ["employee name", "name", "emp name", "staff name", "worker name"]
EMP_CODE_CANDIDATES = ["employee code", "emp code", "emp id", "employee id", "code", "id", "staff id", "worker id"]
DESIGNATION_CANDIDATES = ["designation", "title", "position", "proffession", "profession", "job title"]
ABSENT_DAYS_CANDIDATES = ["leave/days", "leave days", "absent days", "absent", "leave"]
PAY_PERIOD_CANDIDATES = ["pay period", "period", "month", "pay month"]  # fallback only

# ---- Amount mapping by Excel LETTER (adjust if your file changes) ----
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
        if re.fullmatch(r"\d{1,2}:\d{2}(:\d{2})?", s):
            parts = s.split(":")
            h = int(parts[0]); m = int(parts[1]); sec = int(parts[2]) if len(parts)==3 else 0
            return h + m/60 + sec/3600
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

# ========================== PAYSLIP HELPERS ==========================
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
    doc.allowSplitting = 1

    styles = getSampleStyleSheet()
    title_style   = ParagraphStyle("TitleBig",   parent=styles["Title"],   fontSize=HEADER_TITLE_SIZE, leading=20, alignment=1)
    company_style = ParagraphStyle("Company",    parent=styles["Heading2"], fontSize=HEADER_COMPANY_SIZE, leading=16, alignment=1)
    label_style   = ParagraphStyle("Label",      parent=styles["Normal"],   fontSize=TABLE_FONT_SIZE,   leading=TABLE_FONT_SIZE+2)

    elems = []

    if logo_bytes:
        img = Image(io.BytesIO(logo_bytes)); img._restrictSize(logo_width, logo_width*1.2)
        head_tbl = Table([[img, Paragraph(f"<b>{title}</b><br/>{company_name}",
                             ParagraphStyle("hdr", parent=styles["Normal"], fontSize=HEADER_COMPANY_SIZE, leading=16, alignment=1))]],
                         colWidths=[logo_width, None])
        head_tbl.setStyle(TableStyle([
            ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
            ("ALIGN",(1,0),(1,0),"CENTER"),
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
        sheet_name = xls.sheet_names[0]  # first sheet used automatically
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
            total = len(df)
            for i in range(total):
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
        st.download_button(
            "⬇️ Download Payslips (ZIP)",
            data=zbuf.read(),
            file_name=f"Payslips_{sheet_name}_{run_id}.zip",
            mime="application/zip",
        )

# ==========================================================
#                OVERTIME REPORT (WIDE 1..31 FORMAT)
# ==========================================================
st.markdown("---")
st.subheader("Overtime Report (Daily > 8 hours)")

att_file = st.file_uploader("Upload Attendance Excel (.xlsx)", type=["xlsx"], key="attendance")
ot_threshold = st.number_input("Daily threshold (hours)", min_value=0.0, max_value=24.0, value=8.0, step=0.5)

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

def parse_hours_cell(x):
    if pd.isna(x):
        return None
    if isinstance(x, (int, float)):
        return float(x)
    s = str(x).strip()
    if s == "" or s.lower() in {"off","a","absent","leave","holiday","-","--"}:
        return None
    if re.fullmatch(r"\d{1,2}:\d{2}(:\d{2})?", s):
        parts = s.split(":")
        h = int(parts[0]); m = int(parts[1]); sec = int(parts[2]) if len(parts)==3 else 0
        return h + m/60 + sec/3600
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

def build_ot_report_wide(df_att, month_label=""):
    nm = norm_cols_map(df_att.columns)
    name_col = pick_col(nm, "employee name", "name", "emp name", "staff name") or "NAME"
    code_col = pick_col(nm, "employee code", "emp code", "code", "id") or "E CODE"

    day_cols = []
    for c in df_att.columns:
        cs = str(c).strip()
        if re.fullmatch(r"(?:0?[1-9]|[12][0-9]|3[01])", cs):
            day_cols.append(c)
    if not day_cols:
        raise ValueError("Could not find day columns 1..31 in the attendance sheet.")

    long = df_att.melt(
        id_vars=[col for col in [name_col, code_col] if col in df_att.columns],
        value_vars=day_cols,
        var_name="Day",
        value_name="HoursRaw"
    )
    long["Hours"] = long["HoursRaw"].map(parse_hours_cell)
    long["OT_Hours"] = (long["Hours"] - ot_threshold).clip(lower=0)
    long["Has_OT"] = long["OT_Hours"] > 0
    if month_label:
        long["Month"] = month_label

    group_cols = [code_col, name_col]
    summary = (
        long.groupby(group_cols, dropna=False)
            .agg(
                Days_With_OT=("Has_OT", lambda s: int(s.sum())),
                Total_OT_Hours=("OT_Hours", "sum"),
                Total_Work_Hours=("Hours", "sum"),
                Days_Recorded=("Hours", lambda s: int(s.notna().sum())),
            )
            .reset_index()
    )

    disp = summary.copy()
    for c in ["Total_OT_Hours","Total_Work_Hours"]:
        disp[c] = disp[c].map(lambda v: f"{v:.2f}")
    disp = disp.sort_values("Total_OT_Hours", ascending=False)
    return summary, disp, long, name_col, code_col

def monthly_totals(summary_raw: pd.DataFrame, daily_long: pd.DataFrame, month_label: str):
    if summary_raw.empty:
        return pd.DataFrame([{
            "Month": month_label, "Employees": 0, "Days_With_OT": 0,
            "Total_OT_Hours": 0.0, "Total_Work_Hours": 0.0,
            "Days_Recorded": 0, "Avg_Work_Hours_per_Day": 0.0
        }])

    employees         = int(len(summary_raw))
    days_with_ot      = int(summary_raw["Days_With_OT"].sum())
    total_ot_hours    = float(summary_raw["Total_OT_Hours"].sum())
    total_work_hours  = float(summary_raw["Total_Work_Hours"].sum())
    days_recorded     = int(summary_raw["Days_Recorded"].sum())
    avg_work_per_day  = float(daily_long["Hours"].mean()) if not daily_long.empty else 0.0

    return pd.DataFrame([{
        "Month": month_label,
        "Employees": employees,
        "Days_With_OT": days_with_ot,
        "Total_OT_Hours": round(total_ot_hours, 2),
        "Total_Work_Hours": round(total_work_hours, 2),
        "Days_Recorded": days_recorded,
        "Avg_Work_Hours_per_Day": round(avg_work_per_day, 2),
    }])

if att_file:
    try:
        xls2 = pd.ExcelFile(att_file)
        att_sheet = xls2.sheet_names[0]
        df_att = pd.read_excel(xls2, sheet_name=att_sheet, header=0)
        df_att.columns = [str(c).strip() for c in df_att.columns]
        st.success(f"Attendance loaded from sheet '{att_sheet}' with {len(df_att)} rows.")

        summary_raw, summary_disp, daily_long, name_col, code_col = build_ot_report_wide(df_att, month_label=att_sheet)

        st.write("**Overtime Summary (per employee)**")
        st.dataframe(summary_disp.rename(columns={
            code_col: "Employee Code", name_col: "Employee Name"
        }), use_container_width=True)

        month_totals_df = monthly_totals(summary_raw, daily_long, att_sheet)
        st.write("**Monthly Totals (all employees)**")
        st.dataframe(month_totals_df, use_container_width=True)

        # Downloads
        @st.cache_data
        def to_csv_bytes(df):
            return df.to_csv(index=False).encode("utf-8")

        @st.cache_data
        def to_xlsx_bytes(df, sheet="Sheet1"):
            bio = io.BytesIO()
            with pd.ExcelWriter(bio, engine="xlsxwriter") as writer:
                df.to_excel(writer, index=False, sheet_name=sheet)
            bio.seek(0)
            return bio.read()

        c1, c2, c3 = st.columns(3)
        c1.download_button(
            "⬇️ Download OT Summary (CSV)",
            data=to_csv_bytes(summary_raw),
            file_name=f"OT_Summary_{att_sheet}.csv",
            mime="text/csv",
        )
        c2.download_button(
            "⬇️ Download Monthly Totals (CSV)",
            data=to_csv_bytes(month_totals_df),
            file_name=f"OT_Monthly_Totals_{att_sheet}.csv",
            mime="text/csv",
        )
        c3.download_button(
            "⬇️ Download Daily Long (CSV)",
            data=to_csv_bytes(daily_long),
            file_name=f"OT_Daily_{att_sheet}.csv",
            mime="text/csv",
        )

        with st.expander("Show daily long data (one row per employee-day)", expanded=False):
            st.dataframe(daily_long, use_container_width=True)

    except Exception as e:
        st.error(f"Failed to build OT report: {e}")
        st.code(traceback.format_exc(), language="python")

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
