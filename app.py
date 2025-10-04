import io
import re
import zipfile
import traceback
import logging
from datetime import datetime
from typing import Optional, Dict, List, Tuple, Any
from dataclasses import dataclass

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

# ========================== CONFIGURATION ==========================
@dataclass
class AppConfig:
    DEFAULT_COMPANY: str = "AL Glazo Interiors and Décor LLC"
    DEFAULT_TITLE: str = "PAYSLIP"
    MAX_FILE_SIZE_MB: int = 50
    MAX_FILENAME_LENGTH: int = 200
    TABLE_FONT_SIZE: int = 10
    HEADER_TITLE_SIZE: int = 18
    HEADER_COMPANY_SIZE: int = 13
    MAX_DAILY_HOURS: float = 24.0
    MIN_DAILY_HOURS: float = 0.0
    DEFAULT_MONTH_DAYS: int = 30
    
config = AppConfig()

# Setup logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Column name candidates
COLUMN_MAPPINGS = {
    'employee_name': ["employee name", "name", "emp name", "staff name", "worker name"],
    'employee_code': ["employee code", "emp code", "emp id", "employee id", "code", "id", "staff id", "worker id"],
    'designation': ["designation", "title", "position", "profession", "job title"],
    'absent_days': ["leave/days", "leave days", "absent days", "absent", "leave"],
    'pay_period': ["pay period", "period", "month", "pay month"],
    'rate': ["rate", "hourly rate", "per hour", "cost per hour", "wage", "salary/hour", "hr rate", "hour rate"],
    'basic_pay': ["basic", "basic salary", "basic pay"],
    'gross_pay': ["gross salary", "gross", "total salary"],
    'salary_per_day': ["salary/day", "salary per day", "per day salary", "day salary"],
}

ABSENT_TOKENS = {"a", "absent"}
PRESENT_TOKENS = {"p", "present"}
OFF_TOKENS = {"off", "leave", "holiday"}

PAGE_SIZES = {
    "A4 (Landscape)": landscape(A4),
    "Letter (Landscape)": landscape(LETTER)
}

# ========================== VALIDATION ==========================
class ValidationError(Exception):
    """Custom exception for validation errors"""
    pass

def validate_file_size(file) -> None:
    """Validate uploaded file size"""
    if hasattr(file, 'size'):
        size_mb = file.size / (1024 * 1024)
        if size_mb > config.MAX_FILE_SIZE_MB:
            raise ValidationError(f"File too large: {size_mb:.1f}MB. Max: {config.MAX_FILE_SIZE_MB}MB")

def validate_hours(hours: float, context: str = "") -> float:
    """Validate hours are within reasonable range"""
    if pd.isna(hours):
        return 0.0
    if hours < config.MIN_DAILY_HOURS or hours > config.MAX_DAILY_HOURS:
        logger.warning(f"Suspicious hours value: {hours} {context}")
        return max(config.MIN_DAILY_HOURS, min(hours, config.MAX_DAILY_HOURS))
    return hours

def validate_dataframe(df: pd.DataFrame, required_cols: List[str]) -> Tuple[bool, str]:
    """Validate DataFrame has minimum required structure"""
    if df.empty:
        return False, "The uploaded file contains no data"
    
    missing = []
    for col_type in required_cols:
        if col_type in COLUMN_MAPPINGS:
            found = find_column_by_candidates(df, COLUMN_MAPPINGS[col_type])
            if not found:
                missing.append(col_type)
    
    if missing:
        return False, f"Missing required columns: {', '.join(missing)}"
    
    return True, "Validation passed"

# ========================== DATA PARSING UTILITIES ==========================
def clean_string(s: Any) -> str:
    """Clean and normalize string input"""
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return ""
    return re.sub(r"\s+", " ", str(s).strip())

def parse_number(x: Any) -> Optional[float]:
    """Parse various number formats including HH:MM time"""
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return None
    
    s = str(x).strip()
    if s in ["", "-", "–", "nan", "NaN", "None"]:
        return None
    
    # Remove commas
    s = s.replace(",", "")
    
    # Handle parentheses as negative
    m = re.fullmatch(r"\((\d+(?:\.\d+)?)\)", s)
    if m:
        s = "-" + m.group(1)
    
    # Try direct float conversion
    try:
        return float(s)
    except ValueError:
        pass
    
    # Handle HH:MM[:SS] format
    if re.fullmatch(r"\d{1,2}:\d{2}(?::\d{2})?", s):
        parts = s.split(":")
        hours = int(parts[0])
        minutes = int(parts[1])
        seconds = int(parts[2]) if len(parts) > 2 else 0
        return hours + minutes/60 + seconds/3600
    
    # Handle "8h 30m" format
    hours_match = re.search(r"(\d+(?:\.\d+)?)\s*h", s, re.I)
    if hours_match:
        hours = float(hours_match.group(1))
        mins_match = re.search(r"(\d+)\s*m", s, re.I)
        if mins_match:
            hours += int(mins_match.group(1)) / 60
        return hours
    
    return None

def format_amount(x: Any, decimals: int = 2) -> str:
    """Format number as currency amount"""
    value = parse_number(x)
    if value is None:
        value = 0.0
    return f"{value:,.{decimals}f}"

def amount_to_words(x: Any) -> str:
    """Convert amount to words"""
    value = parse_number(x)
    if value is None or pd.isna(value):
        return ""
    
    sign = "minus " if value < 0 else ""
    value = abs(value)
    whole = int(value)
    fraction = int(round((value - whole) * 100))
    
    words = num2words(whole, lang="en").replace("-", " ")
    if fraction:
        words = f"{words} and {fraction:02d}/100"
    
    return (sign + words).strip().capitalize() + " only"

def sanitize_filename(s: Any) -> str:
    """Create safe filename from string"""
    if s is None or (isinstance(s, float) and pd.isna(s)):
        s = ""
    else:
        s = str(s)
    
    # Remove invalid characters
    s = re.sub(r"[\\/:*?\"<>|]+", " ", s).strip()
    s = re.sub(r"\s+", " ", s)
    
    # Limit length
    if len(s) > config.MAX_FILENAME_LENGTH:
        s = s[:config.MAX_FILENAME_LENGTH]
    
    return s or "untitled"

def build_column_lookup(columns: List[str]) -> Dict[str, str]:
    """Create normalized column name lookup"""
    return {str(c).strip().lower(): c for c in columns}

def find_column_by_candidates(df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
    """Find first matching column from candidate list"""
    norm_map = build_column_lookup(df.columns)
    
    # Exact match first
    for candidate in candidates:
        if candidate in norm_map:
            return norm_map[candidate]
    
    # Partial match
    for candidate in candidates:
        for norm_col, orig_col in norm_map.items():
            if candidate in norm_col:
                return orig_col
    
    return None

def get_column_value(row: pd.Series, df: pd.DataFrame, col_type: str) -> Any:
    """Get value from row using column type mapping"""
    if col_type not in COLUMN_MAPPINGS:
        return None
    
    col_name = find_column_by_candidates(df, COLUMN_MAPPINGS[col_type])
    if col_name and col_name in row.index:
        return row[col_name]
    return None

# ========================== DATE PARSING ==========================
MONTH_MAP = {
    "JAN": 1, "JANUARY": 1, "FEB": 2, "FEBRUARY": 2, "MAR": 3, "MARCH": 3,
    "APR": 4, "APRIL": 4, "MAY": 5, "JUN": 6, "JUNE": 6, "JUL": 7, "JULY": 7,
    "AUG": 8, "AUGUST": 8, "SEP": 9, "SEPT": 9, "SEPTEMBER": 9,
    "OCT": 10, "OCTOBER": 10, "NOV": 11, "NOVEMBER": 11, "DEC": 12, "DECEMBER": 12
}

def parse_month_year(label: str) -> Tuple[int, int]:
    """Parse month and year from label like 'SEP 2025' or '09-2025'"""
    s = str(label or "").strip().upper()
    
    # Try to find month name
    month = next((v for k, v in MONTH_MAP.items() if k in s), None)
    
    # Try to find year
    year_match = re.search(r"(20\d{2}|19\d{2})", s)
    year = int(year_match.group(1)) if year_match else datetime.today().year
    
    # Try numeric format MM-YYYY or MM/YYYY
    if month is None:
        numeric_match = re.search(r"\b(\d{1,2})[/-](\d{4})\b", s)
        if numeric_match:
            month = int(numeric_match.group(1))
            year = int(numeric_match.group(2))
    
    # Default to current month if not found
    if month is None:
        month = datetime.today().month
        logger.warning(f"Could not parse month from '{label}', using current month: {month}")
    
    return year, month

def get_days_in_month(year: int, month: int) -> int:
    """Get actual number of days in a specific month"""
    try:
        if month == 12:
            next_month = datetime(year + 1, 1, 1)
        else:
            next_month = datetime(year, month + 1, 1)
        last_day = next_month - pd.Timedelta(days=1)
        return last_day.day
    except Exception:
        return config.DEFAULT_MONTH_DAYS

# ========================== ATTENDANCE PARSING ==========================
def is_absent_cell(x: Any) -> bool:
    """Check if cell indicates absence"""
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return False
    return str(x).strip().lower() in ABSENT_TOKENS

def is_present_token(x: Any) -> bool:
    """Check if cell is 'P' for present"""
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return False
    return str(x).strip().lower() in PRESENT_TOKENS

def parse_hours_cell(x: Any) -> Optional[float]:
    """Parse hours from attendance cell (handles A, P, OFF, hours)"""
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return None
    
    s = str(x).strip().lower()
    
    # Check for off/leave/holiday
    if s in OFF_TOKENS or s in {"-", "--", ""}:
        return None
    
    # Absent = 0 hours
    if s in ABSENT_TOKENS:
        return 0.0
    
    # Present = 8 hours (default)
    if s in PRESENT_TOKENS:
        return 8.0
    
    # Try to parse as number
    hours = parse_number(x)
    if hours is not None:
        return validate_hours(hours, f"from cell '{x}'")
    
    return None

def coerce_day_number(val: Any) -> Optional[int]:
    """Extract day number (1-31) from cell value"""
    if pd.isna(val):
        return None
    
    s = str(val).strip()
    
    # Try direct integer conversion
    try:
        day = int(float(s))
        if 1 <= day <= 31:
            return day
    except ValueError:
        pass
    
    # Extract leading digits
    match = re.match(r"^\s*(\d{1,2})\b", s)
    if match:
        day = int(match.group(1))
        if 1 <= day <= 31:
            return day
    
    return None

def detect_day_columns(df: pd.DataFrame) -> List[str]:
    """Detect columns representing days (1-31)"""
    return [col for col in df.columns if coerce_day_number(col) is not None]

def promote_day_header(df: pd.DataFrame, look_ahead: int = 6) -> pd.DataFrame:
    """Auto-detect and promote day header row if needed"""
    # Check if current headers already have day columns
    if detect_day_columns(df):
        return df
    
    # Search first few rows for day headers
    best_row = None
    best_count = 0
    
    for row_idx in range(min(look_ahead, len(df))):
        row_vals = df.iloc[row_idx].tolist()
        day_count = sum(1 for v in row_vals if coerce_day_number(v) is not None)
        
        if day_count > best_count:
            best_row = row_idx
            best_count = day_count
    
    # Promote row if we found good day columns
    if best_row is not None and best_count >= 5:
        new_headers = [str(v).strip() for v in df.iloc[best_row]]
        new_df = df.iloc[best_row + 1:].copy()
        new_df.columns = new_headers
        new_df.reset_index(drop=True, inplace=True)
        logger.info(f"Promoted row {best_row} as header with {best_count} day columns")
        return new_df
    
    return df

# ========================== PDF GENERATION ==========================
def build_payslip_pdf(
    employee_data: Dict[str, Any],
    company_name: str,
    title: str,
    page_size: tuple,
    currency: str = "AED",
    logo_bytes: Optional[bytes] = None,
    logo_width: float = 1.2 * inch
) -> bytes:
    """Generate PDF payslip for single employee"""
    
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer, pagesize=page_size,
        leftMargin=0.8*inch, rightMargin=0.8*inch,
        topMargin=0.6*inch, bottomMargin=0.6*inch
    )
    
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "TitleBig", parent=styles["Title"],
        fontSize=config.HEADER_TITLE_SIZE, leading=20, alignment=1
    )
    company_style = ParagraphStyle(
        "Company", parent=styles["Heading2"],
        fontSize=config.HEADER_COMPANY_SIZE, leading=16, alignment=1
    )
    
    elements = []
    
    # Header with logo
    if logo_bytes:
        img = Image(io.BytesIO(logo_bytes))
        img._restrictSize(logo_width, logo_width * 1.2)
        header_text = Paragraph(
            f"<b>{title}</b><br/>{company_name}",
            ParagraphStyle(
                "hdr", parent=styles["Normal"],
                fontSize=config.HEADER_COMPANY_SIZE, leading=16, alignment=1
            )
        )
        header_table = Table([[img, header_text]], colWidths=[logo_width, None])
        header_table.setStyle(TableStyle([
            ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
            ("ALIGN", (1,0), (1,0), "CENTER"),
            ("LEFTPADDING", (0,0), (-1,-1), 0),
            ("RIGHTPADDING", (0,0), (-1,-1), 0),
            ("TOPPADDING", (0,0), (-1,-1), 0),
            ("BOTTOMPADDING", (0,0), (-1,-1), 2),
        ]))
        elements.extend([header_table, Spacer(1, 6)])
    else:
        elements.extend([
            Paragraph(title, title_style),
            Paragraph(company_name, company_style),
            Spacer(1, 6)
        ])
    
    # Employee info header
    header_rows = [
        ["Employee Name", clean_string(employee_data.get("Employee Name", ""))],
        ["Employee Code", clean_string(employee_data.get("Employee Code", ""))],
        ["Pay Period", clean_string(employee_data.get("Pay Period", ""))],
        ["Designation", clean_string(employee_data.get("Designation", ""))],
        ["Absent Days", format_amount(employee_data.get("Absent Days", 0), 0)],
    ]
    
    header_table = Table(header_rows, colWidths=[2.3*inch, None])
    header_table.setStyle(TableStyle([
        ("FONTNAME", (0,0), (-1,-1), "Helvetica"),
        ("FONTSIZE", (0,0), (-1,-1), config.TABLE_FONT_SIZE),
        ("GRID", (0,0), (-1,-1), 0.6, colors.black),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("LEFTPADDING", (0,0), (-1,-1), 3),
        ("RIGHTPADDING", (0,0), (-1,-1), 3),
        ("TOPPADDING", (0,0), (-1,-1), 2),
        ("BOTTOMPADDING", (0,0), (-1,-1), 2),
        ("FONTNAME", (0,0), (0,-1), "Helvetica-Bold"),
        ("FONTNAME", (1,0), (1,-1), "Helvetica-Bold"),
    ]))
    elements.extend([header_table, Spacer(1, 6)])
    
    # Earnings table
    earnings_items = [
        "Basic Pay", "Other Allowance", "Housing Allowance",
        "Over time", "Reward for Full Day Attendance", "Incentive"
    ]
    earnings_rows = [[f"Earnings ({currency})", "Amount"]]
    for item in earnings_items:
        earnings_rows.append([item, format_amount(employee_data.get(item, 0), 2)])
    earnings_rows.append(["Total Earnings", format_amount(employee_data.get("Total Earnings", 0), 2)])
    
    earnings_table = Table(earnings_rows, colWidths=[3.6*inch, 1.4*inch], repeatRows=1)
    earnings_table.setStyle(TableStyle([
        ("FONTNAME", (0,0), (-1,-1), "Helvetica"),
        ("FONTSIZE", (0,0), (-1,-1), config.TABLE_FONT_SIZE),
        ("BACKGROUND", (0,0), (1,0), colors.lightgrey),
        ("GRID", (0,0), (-1,-1), 0.6, colors.black),
        ("ALIGN", (1,1), (1,-1), "RIGHT"),
        ("LEFTPADDING", (0,0), (-1,-1), 3),
        ("RIGHTPADDING", (0,0), (-1,-1), 3),
        ("TOPPADDING", (0,0), (-1,-1), 2),
        ("BOTTOMPADDING", (0,0), (-1,-1), 2),
        ("FONTNAME", (0,0), (0,0), "Helvetica-Bold"),
        ("FONTNAME", (0,-1), (1,-1), "Helvetica-Bold"),
    ]))
    
    # Deductions table
    deduction_items = [
        "Absent Pay", "Extra Leave Punishment Ded", "Air Ticket Deduction",
        "Other Fine Ded", "Medical Deduction", "Mob Bill Deduction",
        "I LOE Insurance Deduction", "Sal Advance Deduction"
    ]
    deductions_rows = [[f"Deductions ({currency})", "Amount"]]
    for item in deduction_items:
        deductions_rows.append([item, format_amount(employee_data.get(item, 0), 2)])
    deductions_rows.append(["Total Deductions", format_amount(employee_data.get("Total Deductions", 0), 2)])
    
    deductions_table = Table(deductions_rows, colWidths=[3.6*inch, 1.4*inch], repeatRows=1)
    deductions_table.setStyle(TableStyle([
        ("FONTNAME", (0,0), (-1,-1), "Helvetica"),
        ("FONTSIZE", (0,0), (-1,-1), config.TABLE_FONT_SIZE),
        ("BACKGROUND", (0,0), (1,0), colors.lightgrey),
        ("GRID", (0,0), (-1,-1), 0.6, colors.black),
        ("ALIGN", (1,1), (1,-1), "RIGHT"),
        ("LEFTPADDING", (0,0), (-1,-1), 3),
        ("RIGHTPADDING", (0,0), (-1,-1), 3),
        ("TOPPADDING", (0,0), (-1,-1), 2),
        ("BOTTOMPADDING", (0,0), (-1,-1), 2),
        ("FONTNAME", (0,0), (0,0), "Helvetica-Bold"),
        ("FONTNAME", (0,-1), (1,-1), "Helvetica-Bold"),
    ]))
    
    # Two-column layout for earnings and deductions
    two_col_table = Table(
        [[earnings_table, deductions_table]],
        colWidths=[5.0*inch, 5.0*inch]
    )
    two_col_table.setStyle(TableStyle([
        ("VALIGN", (0,0), (-1,-1), "TOP"),
        ("LEFTPADDING", (0,0), (-1,-1), 0),
        ("RIGHTPADDING", (0,0), (-1,-1), 0),
        ("TOPPADDING", (0,0), (-1,-1), 0),
        ("BOTTOMPADDING", (0,0), (-1,-1), 0),
        ("GRID", (0,0), (-1,-1), 0.3, colors.grey),
    ]))
    elements.extend([two_col_table, Spacer(1, 5)])
    
    # Summary table
    summary_rows = [
        ["Total Earnings", format_amount(employee_data.get("Total Earnings", 0), 2)],
        ["Total Deductions", format_amount(employee_data.get("Total Deductions", 0), 2)],
        ["Net Pay", format_amount(employee_data.get("Net Pay", 0), 2)],
    ]
    
    summary_table = Table(summary_rows, colWidths=[4.6*inch, 1.8*inch], hAlign="CENTER")
    summary_table.setStyle(TableStyle([
        ("FONTNAME", (0,0), (-1,-1), "Helvetica"),
        ("FONTSIZE", (0,0), (-1,-1), config.TABLE_FONT_SIZE),
        ("GRID", (0,0), (-1,-1), 0.6, colors.black),
        ("ALIGN", (1,0), (1,-1), "RIGHT"),
        ("LEFTPADDING", (0,0), (-1,-1), 3),
        ("RIGHTPADDING", (0,0), (-1,-1), 3),
        ("TOPPADDING", (0,0), (-1,-1), 2),
        ("BOTTOMPADDING", (0,0), (-1,-1), 2),
        ("FONTNAME", (0,2), (1,2), "Helvetica-Bold"),
        ("BACKGROUND", (0,2), (1,2), colors.whitesmoke),
    ]))
    
    summary_content = [summary_table]
    
    # Amount in words
    words = amount_to_words(employee_data.get("Net Pay", 0))
    if words:
        label_style = ParagraphStyle(
            "Label", parent=styles["Normal"],
            fontSize=config.TABLE_FONT_SIZE, leading=config.TABLE_FONT_SIZE + 2
        )
        summary_content.extend([
            Spacer(1, 4),
            Paragraph(f"<b>Net to pay (in words):</b> {words}", label_style)
        ])
    
    elements.append(KeepInFrame(
        maxWidth=None, maxHeight=1.8*inch,
        content=summary_content, mergeSpace=1, mode="shrink"
    ))
    elements.append(CondPageBreak(0.6*inch))
    
    # Signature footer
    elements.append(Spacer(1, 18))
    footer_table = Table(
        [["Accounts", "Employee Signature"]],
        colWidths=[4.0*inch, 4.0*inch]
    )
    footer_table.setStyle(TableStyle([
        ("FONTNAME", (0,0), (-1,-1), "Helvetica"),
        ("FONTSIZE", (0,0), (-1,-1), config.TABLE_FONT_SIZE),
        ("ALIGN", (0,0), (0,0), "LEFT"),
        ("ALIGN", (1,0), (1,0), "RIGHT"),
    ]))
    elements.append(footer_table)
    
    # Build PDF
    doc.build(elements)
    buffer.seek(0)
    return buffer.read()

# ========================== STREAMLIT UI ==========================
def main():
    st.set_page_config(
        page_title="Payroll Management System",
        page_icon="💼",
        layout="wide"
    )
    
    st.title("💼 Payroll Management System")
    st.markdown("**Professional payroll processing with overtime tracking and multi-project timesheets**")
    
    # Sidebar configuration
    with st.sidebar:
        st.header("⚙️ Settings")
        company_name = st.text_input(
            "Company Name",
            value=config.DEFAULT_COMPANY
        )
        page_size_label = st.selectbox(
            "PDF Page Size",
            list(PAGE_SIZES.keys()),
            index=0
        )
        currency_label = st.text_input("Currency", value="AED")
        
        st.divider()
        st.header("💰 Rate Settings")
        default_hourly_rate = st.number_input(
            "Default Hourly Rate",
            min_value=0.0,
            value=0.0,
            step=0.5,
            help="Used when rate column is not found"
        )
        ot_multiplier = st.number_input(
            "OT Multiplier",
            min_value=1.0,
            value=1.25,
            step=0.05,
            help="Overtime pay multiplier (e.g., 1.25 = time and a quarter)"
        )
        ot_threshold = st.number_input(
            "Daily OT Threshold (hours)",
            min_value=0.0,
            max_value=24.0,
            value=8.0,
            step=0.5,
            help="Hours worked beyond this count as overtime"
        )
    
    # Main tabs
    tab1, tab2, tab3 = st.tabs([
        "📄 Payslip Generation",
        "⏰ Overtime Report",
        "📊 Multi-Project Dashboard"
    ])
    
    # TAB 1: PAYSLIP GENERATION
    with tab1:
        st.header("Generate Employee Payslips")
        
        col1, col2 = st.columns([2, 1])
        with col1:
            title = st.text_input("PDF Heading", value=config.DEFAULT_TITLE)
        with col2:
            logo_file = st.file_uploader(
                "Company Logo (optional)",
                type=["png", "jpg", "jpeg"],
                help="Logo will appear on payslips"
            )
        
        excel_file = st.file_uploader(
            "Upload Payroll Excel File",
            type=["xlsx"],
            help="Excel file with employee payroll data"
        )
        
        if excel_file:
            try:
                validate_file_size(excel_file)
                
                with st.spinner("Loading payroll data..."):
                    xls = pd.ExcelFile(excel_file)
                    sheet_name = xls.sheet_names[0]
                    df = pd.read_excel(xls, sheet_name=sheet_name, header=0)
                    df.columns = [str(c).strip() for c in df.columns]
                
                st.success(f"✓ Loaded {len(df)} employees from sheet '{sheet_name}'")
                
                # Data preview
                with st.expander("📋 Preview Data (first 10 rows)", expanded=False):
                    st.dataframe(df.head(10), use_container_width=True)
                
                # Validation
                is_valid, msg = validate_dataframe(df, ['employee_name', 'employee_code'])
                if not is_valid:
                    st.error(f"⚠️ {msg}")
                    st.info("💡 The system will attempt to find similar column names")
                
                # Generate button
                if st.button("🚀 Generate All Payslips", type="primary", use_container_width=True):
                    logo_bytes = logo_file.read() if logo_file else None
                    page_size = PAGE_SIZES[page_size_label]
                    
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    
                    zip_buffer = io.BytesIO()
                    success_count = 0
                    error_count = 0
                    
                    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                        for idx in range(len(df)):
                            try:
                                row = df.iloc[idx]
                                
                                # Extract employee data
                                emp_data = {
                                    "Employee Name": clean_string(get_column_value(row, df, 'employee_name')),
                                    "Employee Code": clean_string(get_column_value(row, df, 'employee_code')),
                                    "Designation": clean_string(get_column_value(row, df, 'designation')),
                                    "Pay Period": sheet_name,
                                    "Absent Days": parse_number(get_column_value(row, df, 'absent_days')) or 0,
                                }
                                
                                # Add earnings and deductions
                                earnings_cols = ['Basic Pay', 'Other Allowance', 'Housing Allowance', 
                                               'Over time', 'Reward for Full Day Attendance', 'Incentive']
                                deduction_cols = ['Absent Pay', 'Extra Leave Punishment Ded', 
                                                'Air Ticket Deduction', 'Other Fine Ded', 
                                                'Medical Deduction', 'Mob Bill Deduction',
                                                'I LOE Insurance Deduction', 'Sal Advance Deduction']
                                
                                total_earnings = 0.0
                                for col in earnings_cols:
                                    val = parse_number(row.get(col, 0)) or 0.0
                                    emp_data[col] = val
                                    total_earnings += val
                                
                                total_deductions = 0.0
                                for col in deduction_cols:
                                    val = parse_number(row.get(col, 0)) or 0.0
                                    emp_data[col] = val
                                    total_deductions += val
                                
                                emp_data["Total Earnings"] = total_earnings
                                emp_data["Total Deductions"] = total_deductions
                                emp_data["Net Pay"] = total_earnings - total_deductions
                                
                                # Generate PDF
                                pdf_bytes = build_payslip_pdf(
                                    emp_data, company_name, title, page_size,
                                    currency_label, logo_bytes
                                )
                                
                                # Create filename
                                code_part = sanitize_filename(emp_data["Employee Code"])
                                name_part = sanitize_filename(emp_data["Employee Name"])
                                filename = f"{code_part} - {name_part}.pdf" if code_part else f"Payslip_{idx+1}.pdf"
                                
                                zf.writestr(filename, pdf_bytes)
                                success_count += 1
                                
                            except Exception as e:
                                error_count += 1
                                error_msg = f"Row {idx+1} Error: {str(e)}\n\n{traceback.format_exc()}"
                                zf.writestr(f"ERROR_row_{idx+1}.txt", error_msg)
                                logger.error(f"Failed to process row {idx+1}: {e}")
                            
                            # Update progress
                            progress = (idx + 1) / len(df)
                            progress_bar.progress(progress)
                            status_text.text(f"Processing: {idx+1}/{len(df)} ({success_count} successful, {error_count} errors)")
                    
                    zip_buffer.seek(0)
                    
                    # Show results
                    if error_count > 0:
                        st.warning(f"Completed with {error_count} errors. Check ERROR files in ZIP.")
                    else:
                        st.success(f"Successfully generated {success_count} payslips!")
                    
                    # Download button
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    st.download_button(
                        "⬇️ Download Payslips ZIP",
                        data=zip_buffer.read(),
                        file_name=f"Payslips_{sheet_name}_{timestamp}.zip",
                        mime="application/zip",
                        type="primary"
                    )
