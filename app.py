import io
import re
import zipfile
import traceback
from datetime import datetime

import pandas as pd
import streamlit as st
from num2words import num2words
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots

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

# Helper functions (same as original)
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

def build_pdf_for_row(std, company_name, title, page_size,
                      currency_label="AED", include_words=True, logo_bytes=None, logo_width=1.2*inch) -> bytes:
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=page_size,
                            leftMargin=0.8*inch, rightMargin=0.8*inch,
                            topMargin=0.6*inch, bottomMargin=0.6*inch)
    doc.allowSplitting = 1
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle("TitleBig", parent=styles["Title"], fontSize=HEADER_TITLE_SIZE, leading=20, alignment=1)
    company_style = ParagraphStyle("Company", parent=styles["Heading2"], fontSize=HEADER_COMPANY_SIZE, leading=16, alignment=1)
    label_style = ParagraphStyle("Label", parent=styles["Normal"], fontSize=TABLE_FONT_SIZE, leading=TABLE_FONT_SIZE+2)
    elems = []
    if logo_bytes:
        img = Image(io.BytesIO(logo_bytes)); img._restrictSize(logo_width, logo_width*1.2)
        head_tbl = Table([[img, Paragraph(f"<b>{title}</b><br/>{company_name}",
                             ParagraphStyle("hdr", parent=styles["Normal"], fontSize=HEADER_COMPANY_SIZE, leading=16, alignment=1))]],
                         colWidths=[logo_width, None])
        head_tbl.setStyle(TableStyle([("VALIGN",(0,0),(-1,-1),"MIDDLE"), ("ALIGN",(1,0),(1,0),"CENTER"),
            ("LEFTPADDING",(0,0),(-1,-1),0), ("RIGHTPADDING",(0,0),(-1,-1),0),
            ("TOPPADDING",(0,0),(-1,-1),0), ("BOTTOMPADDING",(0,0),(-1,-1),2)]))
        elems += [head_tbl, Spacer(1,6)]
    else:
        elems += [Paragraph(title, title_style), Paragraph(company_name, company_style), Spacer(1,6)]
    hdr_rows = [["Employee Name", clean(std.get("Employee Name",""))], ["Employee Code", clean(std.get("Employee Code",""))],
        ["Pay Period", clean(std.get("Pay Period",""))], ["Designation", clean(std.get("Designation",""))],
        ["Absent Days", fmt_amount(std.get("Absent Days",0), 0)]]
    hdr_tbl = Table(hdr_rows, colWidths=[2.3*inch, None])
    hdr_tbl.setStyle(TableStyle([("FONTNAME",(0,0),(-1,-1),"Helvetica"), ("FONTSIZE",(0,0),(-1,-1),TABLE_FONT_SIZE),
        ("GRID",(0,0),(-1,-1),0.6,colors.black), ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
        ("LEFTPADDING",(0,0),(-1,-1),3), ("RIGHTPADDING",(0,0),(-1,-1),3),
        ("TOPPADDING",(0,0),(-1,-1),2), ("BOTTOMPADDING",(0,0),(-1,-1),2),
        ("FONTNAME",(0,0),(0,0),"Helvetica-Bold"), ("FONTNAME",(1,0),(1,0),"Helvetica-Bold")]))
    elems += [hdr_tbl, Spacer(1,6)]
    earn_rows = [[f"Earnings ({currency_label})","Amount"]]
    for lbl in ["Basic Pay","Other Allowance","Housing Allowance","Over time","Reward for Full Day Attendance","Incentive"]:
        earn_rows.append([lbl, fmt_amount(std.get(lbl,0),2)])
    earn_rows.append(["Total Earnings", fmt_amount(std.get("Total Earnings (optional)",0),2)])
    earn_tbl = Table(earn_rows, colWidths=[3.6*inch,1.4*inch], repeatRows=1)
    earn_tbl.setStyle(TableStyle([("FONTNAME",(0,0),(-1,-1),"Helvetica"), ("FONTSIZE",(0,0),(-1,-1),TABLE_FONT_SIZE),
        ("BACKGROUND",(0,0),(1,0),colors.lightgrey), ("GRID",(0,0),(-1,-1),0.6,colors.black), ("ALIGN",(1,1),(1,-1),"RIGHT"),
        ("LEFTPADDING",(0,0),(-1,-1),3), ("RIGHTPADDING",(0,0),(-1,-1),3), ("TOPPADDING",(0,0),(-1,-1),2), ("BOTTOMPADDING",(0,0),(-1,-1),2),
        ("FONTNAME",(0,0),(0,0),"Helvetica-Bold"), ("FONTNAME",(0,-1),(1,-1),"Helvetica-Bold")]))
    ded_rows = [[f"Deductions ({currency_label})","Amount"]]
    for lbl in ["Absent Pay","Extra Leave Punishment Ded","Air Ticket Deduction","Other Fine Ded",
                "Medical Deduction","Mob Bill Deduction","I LOE Insurance Deduction","Sal Advance Deduction"]:
        ded_rows.append([lbl, fmt_amount(std.get(lbl,0),2)])
    ded_rows.append(["Total Deductions", fmt_amount(std.get("Total Deductions (optional)",0),2)])
    ded_tbl = Table(ded_rows, colWidths=[3.6*inch,1.4*inch], repeatRows=1)
    ded_tbl.setStyle(TableStyle([("FONTNAME",(0,0),(-1,-1),"Helvetica"), ("FONTSIZE",(0,0),(-1,-1),TABLE_FONT_SIZE),
        ("BACKGROUND",(0,0),(1,0),colors.lightgrey), ("GRID",(0,0),(-1,-1),0.6,colors.black), ("ALIGN",(1,1),(1,-1),"RIGHT"),
        ("LEFTPADDING",(0,0),(-1,-1),3), ("RIGHTPADDING",(0,0),(-1,-1),3), ("TOPPADDING",(0,0),(-1,-1),2), ("BOTTOMPADDING",(0,0),(-1,-1),2),
        ("FONTNAME",(0,0),(0,0),"Helvetica-Bold"), ("FONTNAME",(0,-1),(1,-1),"Helvetica-Bold")]))
    two_col = Table([[earn_tbl, ded_tbl]], colWidths=[5.0*inch, 5.0*inch])
    two_col.setStyle(TableStyle([("VALIGN",(0,0),(-1,-1),"TOP"), ("LEFTPADDING",(0,0),(-1,-1),0), ("RIGHTPADDING",(0,0),(-1,-1),0),
        ("TOPPADDING",(0,0),(-1,-1),0), ("BOTTOMPADDING",(0,0),(-1,-1),0), ("GRID",(0,0),(-1,-1),0.3,colors.grey)]))
    elems += [two_col, Spacer(1,5)]
    sum_rows = [["Total Earnings", fmt_amount(std.get("Total Earnings (optional)",0),2)],
        ["Total Deductions", fmt_amount(std.get("Total Deductions (optional)",0),2)],
        ["Net Pay", fmt_amount(std.get("Net Pay (optional)",0),2)]]
    sum_tbl = Table(sum_rows, colWidths=[4.6*inch, 1.8*inch], hAlign="CENTER")
    sum_tbl.setStyle(TableStyle([("FONTNAME",(0,0),(-1,-1),"Helvetica"), ("FONTSIZE",(0,0),(-1,-1),TABLE_FONT_SIZE),
        ("GRID",(0,0),(-1,-1),0.6,colors.black), ("ALIGN",(1,0),(1,-1),"RIGHT"),
        ("LEFTPADDING",(0,0),(-1,-1),3), ("RIGHTPADDING",(0,0),(-1,-1),3), ("TOPPADDING",(0,0),(-1,-1),2), ("BOTTOMPADDING",(0,0),(-1,-1),2),
        ("FONTNAME",(0,2),(1,2),"Helvetica-Bold"), ("BACKGROUND",(0,2),(1,2),colors.whitesmoke)]))
    summary_block = [sum_tbl]
    words = amount_in_words(std.get("Net Pay (optional)",0))
    if words:
        summary_block += [Spacer(1,4), Paragraph(f"<b>Net to pay (in words):</b> {words}", label_style)]
    elems.append(KeepInFrame(maxWidth=None, maxHeight=1.8*inch, content=summary_block, mergeSpace=1, mode="shrink"))
    elems.append(CondPageBreak(0.6*inch))
    elems.append(Spacer(1, 18))
    foot = Table([["Accounts","Employee Signature"]], colWidths=[4.0*inch, 4.0*inch])
    foot.setStyle(TableStyle([("FONTNAME",(0,0),(-1,-1),"Helvetica"), ("FONTSIZE",(0,0),(-1,-1),TABLE_FONT_SIZE),
        ("ALIGN",(0,0),(0,0),"LEFT"), ("ALIGN",(1,0),(1,0),"RIGHT")]))
    elems.append(foot)
    doc.build(elems)
    buf.seek(0)
    return buf.read()

def build_std_for_row(row_series, row_vals, norm_map, max_cols, pay_period_text=""):
    emp_name = clean(get_value(row_series, norm_map, EMP_NAME_CANDIDATES))
    emp_code = clean(get_value(row_series, norm_map, EMP_CODE_CANDIDATES))
    designation = clean(get_value(row_series, norm_map, DESIGNATION_CANDIDATES))
    absent_days = get_absent_days(row_series, norm_map)
    pay_period = pay_period_text or clean(get_value(row_series, norm_map, PAY_PERIOD_CANDIDATES))
    std = {"Employee Name": emp_name, "Employee Code": emp_code, "Designation": designation, "Pay Period": pay_period, "Absent Days": absent_days}
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

st.set_page_config(page_title="Payroll Analytics", page_icon="🧾", layout="wide")
st.title("💼 Payroll & Timesheet Analytics Dashboard")

with st.expander("⚙️ Settings", expanded=True):
    colA, colB = st.columns([2,1])
    company_name = colA.text_input("Company name", value=DEFAULT_COMPANY)
    page_size_label = colB.selectbox("Page size", list(PAGE_SIZES.keys()), index=0)
    default_hourly_rate = colA.number_input("Default hourly rate", min_value=0.0, value=0.0, step=0.5)
    default_ot_multiplier = colB.number_input("OT multiplier", min_value=1.0, value=1.25, step=0.05)

title = st.text_input("PDF heading text", value=DEFAULT_TITLE)

with st.expander("🎨 Branding", expanded=False):
    c1, c2 = st.columns([1,1])
    currency_label = c1.text_input("Currency label", value="AED")
    logo_file = c2.file_uploader("Logo", type=["png","jpg","jpeg"])

excel_file = st.file_uploader("📊 Upload Payroll Excel", type=["xlsx"], key="payroll_file")

if excel_file:
    try:
        xls = pd.ExcelFile(excel_file)
        sheet_name = xls.sheet_names[0]
        df = pd.read_excel(xls, sheet_name=sheet_name, header=0)
        df.columns = [str(c).strip() for c in df.columns]
        st.success(f"✅ Loaded {len(df)} rows")
        payroll_data = []
        values = df.values
        for i in range(len(df)):
            row_series = df.iloc[i]
            row_vals = values[i]
            try:
                norm_map = build_lookup(df.columns)
                std = build_std_for_row(row_series, row_vals, norm_map, df.shape[1], pay_period_text=sheet_name)
                payroll_data.append(std)
            except:
                pass
        payroll_df = pd.DataFrame(payroll_data)
        st.markdown("---")
        st.subheader("📊 Payroll Analytics Dashboard")
        tab1, tab2, tab3, tab4 = st.tabs(["💰 Overview", "📈 Earnings & Deductions", "👥 Employees", "📉 Details"])
        with tab1:
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("Employees", len(payroll_df))
            with col2:
                total_earnings = payroll_df["Total Earnings (optional)"].sum()
                st.metric("Total Earnings", f"{currency_label} {total_earnings:,.2f}")
            with col3:
                total_deductions = payroll_df["Total Deductions (optional)"].sum()
                st.metric("Total Deductions", f"{currency_label} {total_deductions:,.2f}")
            with col4:
                net_payroll = payroll_df["Net Pay (optional)"].sum()
                st.metric("Net Payroll", f"{currency_label} {net_payroll:,.2f}")
            fig_dist = px.histogram(payroll_df, x="Net Pay (optional)", title="Net Pay Distribution",
                labels={"Net Pay (optional)": f"Net Pay ({currency_label})"}, color_discrete_sequence=["#3b82f6"])
            fig_dist.update_layout(showlegend=False, height=400)
            st.plotly_chart(fig_dist, use_container_width=True)
        with tab2:
            fig_compare = go.Figure()
            fig_compare.add_trace(go.Bar(name='Earnings', x=payroll_df["Employee Name"],
                y=payroll_df["Total Earnings (optional)"], marker_color='#10b981'))
            fig_compare.add_trace(go.Bar(name='Deductions', x=payroll_df["Employee Name"],
                y=payroll_df["Total Deductions (optional)"], marker_color='#ef4444'))
            fig_compare.update_layout(title="Earnings vs Deductions", xaxis_title="Employee",
                yaxis_title=f"Amount ({currency_label})", barmode='group', height=500)
            st.plotly_chart(fig_compare, use_container_width=True)
            earnings_cols = ["Basic Pay", "Other Allowance", "Housing Allowance", "Over time", "Reward for Full Day Attendance", "Incentive"]
            earnings_totals = payroll_df[earnings_cols].sum()
            fig_earnings = px.pie(values=earnings_totals.values, names=earnings_totals.index, title="Earnings Breakdown",
                color_discrete_sequence=px.colors.sequential.Blues_r)
            fig_earnings.update_traces(textposition='inside', textinfo='percent+label')
            st.plotly_chart(fig_earnings, use_container_width=True)
        with tab3:
            top_earners = payroll_df.nlargest(10, "Net Pay (optional)")
            fig_top = px.bar(top_earners, y="Employee Name", x="Net Pay (optional)", orientation='h', title="Top 10 Earners",
                labels={"Net Pay (optional)": f"Net Pay ({currency_label})", "Employee Name": "Employee"},
                color="Net Pay (optional)", color_continuous_scale="Viridis")
            fig_top.update_layout(height=500, showlegend=False)
            st.plotly_chart(fig_top, use_container_width=True)
            fig_absent = px.scatter(payroll_df, x="Absent Days", y="Net Pay (optional)", size="Total Deductions (optional)",
                hover_data=["Employee Name", "Designation"], title="Absent Days vs Net Pay",
                labels={"Net Pay (optional)": f"Net Pay ({currency_label})", "Absent Days": "Days Absent"},
                color="Designation", color_discrete_sequence=px.colors.qualitative.Set2)
            fig_absent.update_layout(height=500)
            st.plotly_chart(fig_absent, use_container_width=True)
        with tab4:
            if "Designation" in payroll_df.columns:
                des_summary = payroll_df.groupby("Designation").agg({
                    "Net Pay (optional)": ["mean", "sum", "count"], "Absent Days": "mean"}).round(2)
                des_summary.columns = ['Avg Net Pay', 'Total Net Pay', 'Count', 'Avg Absent Days']
                des_summary = des_summary.reset_index()
                fig_des = px.bar(des_summary, x="Designation", y="Total Net Pay", title="Total Payroll by Designation",
                    labels={"Total Net Pay": f"Total Net Pay ({currency_label})"}, color="Count", color_continuous_scale="Sunset")
                fig_des.update_layout(height=400)
                st.plotly_chart(fig_des, use_container_width=True)
                st.dataframe(des_summary, use_container_width=True)
            deductions_cols = ["Absent Pay", "Extra Leave Punishment Ded", "Air Ticket Deduction", "Other Fine Ded",
                "Medical Deduction", "Mob Bill Deduction", "I LOE Insurance Deduction", "Sal Advance Deduction"]
            deductions_totals = payroll_df[deductions_cols].sum()
            deductions_totals = deductions_totals[deductions_totals > 0]
            if len(deductions_totals) > 0:
                fig_ded = px.bar(x=deductions_totals.index, y=deductions_totals.values, title="Deductions Breakdown",
                    labels={"x": "Type", "y": f"Amount ({currency_label})"}, color=deductions_totals.values, color_continuous_scale="Reds")
                fig_ded.update_layout(height=400, showlegend=False)
                st.plotly_chart(fig_ded, use_container_width=True)
    except Exception as e:
        st.error(f"Error: {e}")
        st.stop()
    if st.button("🎫 Generate Payslips ZIP", type="primary"):
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
                    pdf_bytes = build_pdf_for_row(std, company_name, title, page_size,
                        currency_label=currency_label, include_words=True, logo_bytes=logo_bytes)
                    parts = [safe_filename(std["Employee Code"]), safe_filename(std["Employee Name"])]
                    parts = [p for p in parts if p]
                    fname = (" - ".join(parts) if parts else f"Payslip_{i+1}") + ".pdf"
                    zf.writestr(fname, pdf_bytes)
                except Exception as e:
                    tb = traceback.format_exc()
                    zf.writestr(f"row_{i+1}_ERROR.txt", f"Row {i+1}: {e}\n\n{tb}")
        zbuf.seek(0)
        run_id = datetime.now().strftime("%Y%m%d-%H%M%S")
        st.download_button("⬇️ Download ZIP", data=zbuf.read(),
            file_name=f"Payslips_{sheet_name}_{run_id}.zip", mime="application/zip")

st.markdown("---")
st.subheader("📅 Multi-Project Timesheets Dashboard")

PAY_BASIC_CANDS = ["basic", "basic salary", "basic pay"]
PAY_GROSS_CANDS = ["gross salary", "gross", "total salary"]
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
    hour_cols = [c for c in (hour_cols or []) if c in df.columns]
    int_cols = [c for c in (int_cols or []) if c in df.columns]
    fmts = {}
    fmts.update({c: f"{{:,.{decimals_money}f}}" for c in money_cols})
    fmts.update({c: f"{{:,.{decimals_hours}f}}" for c in hour_cols})
    fmts.update({c: "{:,.0f}" for c in int_cols})
    return df.style.format(fmts)

def tidy_one_sheet(df_sheet: pd.DataFrame, sheet_name: str, project_from_file: str, default_ot_multiplier: float):
    dfp2 = promote_day_header_if_needed(df_sheet)
    dfp2.columns = [str(c).strip() for c in dfp2.columns]
    nm = norm_cols_map(dfp2.columns)
    name_col = pick_col(nm, "employee name", "name", "emp name", "staff name") or "NAME"
    code_col = pick_col(nm, "employee code", "emp code", "code", "id") or "E CODE"
    day_cols = detect_day_columns_from_headers(dfp2)
    if not day_cols:
        raise ValueError("Could not find day columns 1..31 in a sheet.")
    basic_col = pick_col(nm, "basic", "basic salary", "basic pay")
    gross_col = pick_col(nm, "gross salary", "gross", "total salary")
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
    return long

multi_files = st.file_uploader("Upload project timesheets (multiple .xlsx)", type=["xlsx"], accept_multiple_files=True, key="multi_sheets")

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
                    long = tidy_one_sheet(df_sheet, sheet, project_from_file, default_ot_multiplier)
                    all_daily.append(long)
                except Exception as inner_e:
                    st.warning(f"[{project_from_file}] Skipped sheet '{sheet}': {inner_e}")
        except Exception as e:
            st.error(f"Failed to parse file {f.name}: {e}")
    
    if all_daily:
        proj_daily = pd.concat(all_daily, ignore_index=True)
        st.markdown("### 🔍 Filters")
        projects = sorted(proj_daily["Project"].dropna().unique().tolist())
        sel_projects = st.multiselect("Projects", projects, default=projects)
        
        if "Date" in proj_daily.columns and not proj_daily["Date"].isna().all():
            min_date = pd.to_datetime(proj_daily["Date"].min()).date()
            max_date = pd.to_datetime(proj_daily["Date"].max()).date()
            c_from, c_to = st.columns(2)
            date_from = c_from.date_input("From", value=min_date, min_value=min_date, max_value=max_date, key="from_d")
            date_to = c_to.date_input("To", value=max_date, min_value=min_date, max_value=max_date, key="to_d")
            if date_from > date_to: date_from, date_to = date_to, date_from
            mask = proj_daily["Project"].isin(sel_projects) & proj_daily["Date"].between(pd.to_datetime(date_from), pd.to_datetime(date_to))
        else:
            min_day, max_day = int(proj_daily["Day"].min()), int(proj_daily["Day"].max())
            day_from, day_to = st.slider("Day range", min_value=min_day, max_value=max_day, value=(min_day, max_day))
            mask = proj_daily["Project"].isin(sel_projects) & proj_daily["Day"].between(int(day_from), int(day_to))
        
        filt_daily = proj_daily.loc[mask].copy()
        
        if not filt_daily.empty:
            emp_presence = (filt_daily.groupby(["Project","Employee Code","Employee Name"], dropna=False)
                .agg(Total_Work_Hours=("Hours","sum"), Worked_Days=("Worked_Flag","sum")).reset_index())
            keep_keys = emp_presence.loc[(emp_presence["Total_Work_Hours"] > 0) | (emp_presence["Worked_Days"] > 0),
                ["Project","Employee Code"]].drop_duplicates()
            if not keep_keys.empty:
                filt_daily = filt_daily.merge(keep_keys, on=["Project","Employee Code"], how="inner")
            else:
                filt_daily = filt_daily.iloc[0:0]
        
        for c in ["Hours","OT_Hours","Base_Daily_Cost","OT_Cost","Total_Daily_Cost","Day"]:
            if c not in filt_daily.columns: filt_daily[c] = 0.0
            filt_daily[c] = pd.to_numeric(filt_daily[c], errors="coerce").fillna(0.0)
        for c in ["Employee Name","Worked_Flag","Is_Absent"]:
            if c not in filt_daily.columns:
                filt_daily[c] = "" if c == "Employee Name" else False
        
        attendance_summary = (filt_daily.groupby(["Project","Employee Code","Employee Name"], dropna=False)
            .agg(Present_Days=("Worked_Flag", lambda s: int(s.sum())), Absent_Days=("Is_Absent", lambda s: int(s.sum())),
                Total_Hours=("Hours","sum"), OT_Days=("OT_Hours", lambda s: int((s>0).sum())), OT_Hours=("OT_Hours","sum"),
                Base_Cost=("Base_Daily_Cost","sum"), OT_Cost=("OT_Cost","sum"), Total_Cost=("Total_Daily_Cost","sum"))
            .reset_index().sort_values(["Project","Employee Name"]))
        
        work_daily = filt_daily.loc[filt_daily["Worked_Flag"] == True].copy()
        
        if len(work_daily) == 0:
            st.info("No worked days found")
            by_proj_day = pd.DataFrame()
            emp_daily = pd.DataFrame()
            project_totals = pd.DataFrame()
        else:
            by_proj_day = (work_daily.groupby(["Project","Day"], dropna=False)
                .agg(Employees=("Employee Name", lambda s: s.nunique()), Hours=("Hours","sum"), OT_Hours=("OT_Hours","sum"),
                    Base_Cost=("Base_Daily_Cost","sum"), OT_Cost=("OT_Cost","sum"), Total_Cost=("Total_Daily_Cost","sum"))
                .reset_index().sort_values(["Project","Day"]))
            by_proj_day["Cum_Hours"] = by_proj_day.groupby("Project")["Hours"].cumsum()
            by_proj_day["Cum_OT_Hours"] = by_proj_day.groupby("Project")["OT_Hours"].cumsum()
            by_proj_day["Cum_Base_Cost"] = by_proj_day.groupby("Project")["Base_Cost"].cumsum()
            by_proj_day["Cum_OT_Cost"] = by_proj_day.groupby("Project")["OT_Cost"].cumsum()
            by_proj_day["Cum_Total_Cost"] = by_proj_day.groupby("Project")["Total_Cost"].cumsum()
            by_proj_day["Accumulated"] = by_proj_day["Cum_Total_Cost"]
            
            emp_daily = (work_daily[["Project","Employee Code","Employee Name","Day","Hours","OT_Hours",
                "Salary_Day","OT_Rate","Base_Daily_Cost","OT_Cost","Total_Daily_Cost"]].sort_values(["Project","Employee Code","Day"]))
            emp_daily["Total_Hours"] = emp_daily["Hours"].fillna(0) + emp_daily["OT_Hours"].fillna(0)
            emp_daily["Accumulated"] = emp_daily.groupby(["Project","Employee Code"])["Total_Daily_Cost"].cumsum()
            
            project_totals = (work_daily.groupby(["Project"], dropna=False)
                .agg(Employees=("Employee Name", lambda s: s.nunique()), Total_Work_Hours=("Hours", "sum"),
                    Total_OT_Hours=("OT_Hours", "sum"), Base_Cost=("Base_Daily_Cost", "sum"),
                    OT_Cost=("OT_Cost", "sum"), Total_Cost=("Total_Daily_Cost", "sum"))
                .reset_index().sort_values("Project"))
        
        st.markdown("---")
        st.subheader("📊 Timesheet Analytics")
        
        tab1, tab2, tab3, tab4, tab5 = st.tabs(["📊 Overview", "📈 Trends", "👥 Employees", "⏰ Hours", "📋 Tables"])
        
        with tab1:
            if not project_totals.empty:
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.metric("Projects", len(project_totals))
                with col2:
                    st.metric("Employees", int(project_totals["Employees"].sum()))
                with col3:
                    st.metric("Total Hours", f"{project_totals['Total_Work_Hours'].sum():,.1f}")
                with col4:
                    st.metric("Total Cost", f"{currency_label} {project_totals['Total_Cost'].sum():,.2f}")
                
                fig_proj = go.Figure()
                fig_proj.add_trace(go.Bar(name='Base', x=project_totals["Project"], y=project_totals["Base_Cost"], marker_color='#3b82f6'))
                fig_proj.add_trace(go.Bar(name='OT', x=project_totals["Project"], y=project_totals["OT_Cost"], marker_color='#f59e0b'))
                fig_proj.update_layout(title="Project Costs", xaxis_title="Project", yaxis_title=f"Cost ({currency_label})",
                    barmode='stack', height=500)
                st.plotly_chart(fig_proj, use_container_width=True)
                
                fig_pie = px.pie(project_totals, values="Total_Cost", names="Project", title="Cost by Project",
                    color_discrete_sequence=px.colors.qualitative.Set3)
                st.plotly_chart(fig_pie, use_container_width=True)
        
        with tab2:
            if not by_proj_day.empty:
                fig_cumul = px.line(by_proj_day, x="Day", y="Accumulated", color="Project", title="Cumulative Cost",
                    labels={"Accumulated": f"Cost ({currency_label})"}, markers=True)
                fig_cumul.update_layout(height=500)
                st.plotly_chart(fig_cumul, use_container_width=True)
                
                fig_daily = px.bar(by_proj_day, x="Day", y="Total_Cost", color="Project", title="Daily Cost",
                    labels={"Total_Cost": f"Cost ({currency_label})"}, barmode='group')
                st.plotly_chart(fig_daily, use_container_width=True)
                
                fig_scatter = px.scatter(by_proj_day, x="Hours", y="Total_Cost", size="Employees", color="Project",
                    title="Hours vs Cost", hover_data=["Day", "OT_Hours"])
                st.plotly_chart(fig_scatter, use_container_width=True)
        
        with tab3:
            if not attendance_summary.empty:
                top_hrs = attendance_summary.nlargest(15, "Total_Hours")
                fig_top = px.bar(top_hrs, y="Employee Name", x="Total_Hours", color="Project", orientation='h',
                    title="Top 15 by Hours", labels={"Total_Hours": "Hours"})
                fig_top.update_layout(height=600)
                st.plotly_chart(fig_top, use_container_width=True)
                
                top_cost = attendance_summary.nlargest(15, "Total_Cost")
                fig_cost = px.bar(top_cost, y="Employee Name", x="Total_Cost", color="OT_Cost", orientation='h',
                    title="Top 15 by Cost", labels={"Total_Cost": f"Cost ({currency_label})"}, color_continuous_scale="Reds")
                fig_cost.update_layout(height=600)
                st.plotly_chart(fig_cost, use_container_width=True)
                
                fig_att = px.scatter(attendance_summary, x="Present_Days", y="Absent_Days", size="Total_Cost",
                    color="Project", hover_data=["Employee Name", "Total_Hours"], title="Attendance Pattern")
                st.plotly_chart(fig_att, use_container_width=True)
        
        with tab4:
            if not work_daily.empty:
                fig_dist = px.histogram(work_daily[work_daily["Hours"] > 0], x="Hours", color="Project",
                    title="Hours Distribution", nbins=20, barmode='overlay', opacity=0.7)
                st.plotly_chart(fig_dist, use_container_width=True)
                
                ot_summary = work_daily[work_daily["OT_Hours"] > 0].groupby("Project").agg({
                    "OT_Hours": "sum", "OT_Cost": "sum", "Employee Name": "nunique"}).reset_index()
                ot_summary.columns = ["Project", "Total OT Hours", "Total OT Cost", "Employees with OT"]
                
                if not ot_summary.empty:
                    fig_ot = make_subplots(rows=1, cols=2, subplot_titles=("OT Hours", "OT Cost"))
                    fig_ot.add_trace(go.Bar(x=ot_summary["Project"], y=ot_summary["Total OT Hours"],
                        name="Hours", marker_color="#f59e0b"), row=1, col=1)
                    fig_ot.add_trace(go.Bar(x=ot_summary["Project"], y=ot_summary["Total OT Cost"],
                        name="Cost", marker_color="#ef4444"), row=1, col=2)
                    fig_ot.update_layout(height=500, showlegend=False)
                    st.plotly_chart(fig_ot, use_container_width=True)
                    st.dataframe(ot_summary, use_container_width=True)
        
        with tab5:
            st.markdown("#### Attendance Summary")
            st.dataframe(fmt_commas(attendance_summary, money_cols=["Base_Cost","OT_Cost","Total_Cost"],
                hour_cols=["Total_Hours","OT_Hours"], int_cols=["Present_Days","Absent_Days","OT_Days"]),
                use_container_width=True)
            
            if not by_proj_day.empty:
                st.markdown("#### Project × Day")
                st.dataframe(fmt_commas(by_proj_day, money_cols=["Base_Cost","OT_Cost","Total_Cost","Cum_Base_Cost","Cum_OT_Cost","Cum_Total_Cost","Accumulated"],
                    hour_cols=["Hours","OT_Hours","Cum_Hours","Cum_OT_Hours"], int_cols=["Employees","Day"]),
                    use_container_width=True)
            
            if not emp_daily.empty:
                st.markdown("#### Employee Daily")
                st.dataframe(fmt_commas(emp_daily, money_cols=["Salary_Day","OT_Rate","Base_Daily_Cost","OT_Cost","Total_Daily_Cost","Accumulated"],
                    hour_cols=["Hours","OT_Hours","Total_Hours"], int_cols=["Day"]),
                    use_container_width=True)
            
            if not project_totals.empty:
                st.markdown("#### Project Totals")
                st.dataframe(fmt_commas(project_totals, money_cols=["Base_Cost","OT_Cost","Total_Cost"],
                    hour_cols=["Total_Work_Hours","Total_OT_Hours"], int_cols=["Employees"]),
                    use_container_width=True)
        
        @st.cache_data
        def to_csv_bytes(df): return df.to_csv(index=False).encode("utf-8")
        
        @st.cache_data
        def to_xlsx_bytes(dfs: dict):
            bio = io.BytesIO()
            with pd.ExcelWriter(bio, engine="openpyxl") as writer:
                for sheet, df in dfs.items():
                    if not df.empty:
                        df.to_excel(writer, index=False, sheet_name=sheet[:31])
            bio.seek(0)
            return bio.read()
        
        st.markdown("---")
        st.markdown("### 📥 Export")
        c1, c2, c3, c4, c5 = st.columns(5)
        if not by_proj_day.empty:
            c1.download_button("⬇️ Project×Day", data=to_csv_bytes(by_proj_day), file_name="Project_Day.csv")
        if not emp_daily.empty:
            c2.download_button("⬇️ Employee Daily", data=to_csv_bytes(emp_daily), file_name="Employee_Daily.csv")
        if not project_totals.empty:
            c3.download_button("⬇️ Totals", data=to_csv_bytes(project_totals), file_name="Totals.csv")
        if not attendance_summary.empty:
            c4.download_button("⬇️ Attendance", data=to_csv_bytes(attendance_summary), file_name="Attendance.csv")
        
        excel_data = {"Project_Day": by_proj_day, "Employee_Daily": emp_daily, "Totals": project_totals,
            "Attendance": attendance_summary, "All_Data": work_daily}
        c5.download_button("⬇️ Excel Pack", data=to_xlsx_bytes(excel_data), file_name="Timesheet_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

st.markdown("---")
st.markdown(f'<div style="text-align:center;color:#6b7280;font-size:12px;">{UI_POWERED_BY_TEXT}</div>', unsafe_allow_html=True)
