>>> import io
... import re
... import hashlib
... from datetime import datetime
... 
... import streamlit as st
... import pandas as pd
... import pdfplumber
... from openpyxl import load_workbook
... from openpyxl.styles import Font, Alignment
... from openpyxl.utils import get_column_letter
... 
... # MySQL
... try:
...     import mysql.connector
...     MYSQL_AVAILABLE = True
... except ImportError:
...     MYSQL_AVAILABLE = False
... 
... 
... # ============================================================
... # STREAMLIT APP
... # PDF -> EXCEL + MYSQL
... # ============================================================
... 
... st.set_page_config(
...     page_title="Quotation PDF Extractor",
...     page_icon="📄",
...     layout="wide"
... )
... 
... st.title("📄 Quotation PDF Extractor")
... st.caption("Extract quotation PDFs → Excel → MySQL")
... 
... # ============================================================
... # CONFIG
... # ============================================================

DEFAULT_HOST = "127.0.0.1"
DEFAULT_PORT = 3306
DEFAULT_USER = "root"
DEFAULT_PASSWORD = "dhinesh"

FINAL_COLUMNS = [
    "Source PDF",
    "Quotation No.",
    "Quotation Date",
    "To",
    "Quotation Validity(in days)",
    "Inquiry Ref",
    "Inquiry Date",
    "Kind Attention",
    "SUB :-",
    "Sr. No.",
    "Item Code & Description",
    "HSN CODE",
    "Qty",
    "SUOM",
    "Unit Sales Price (INR)",
    "Total (INR)",
    "Packing & Forward %",
    "Packing & Forward",
    "Freight on Sales",
    "SGST %",
    "SGST",
    "CGST %",
    "CGST",
    "IGST %",
    "IGST",
    "Total Tax Amount",
    "Grand Total",
    "Amount (in Words)",
    "Prepared By",
    "Extraction_Check",
]


# ============================================================
# PDF FUNCTIONS
# ============================================================

def clean_number(value):
    if value is None:
        return ""
    value = str(value).strip()
    if not value:
        return ""
    value = value.replace(",", "")
    try:
        return float(value)
    except Exception:
        return value


def normalize_text(text):
    if not text:
        return ""
    text = text.replace("\xa0", " ")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n\s*\n+", "\n", text)
    return text.strip()


def find_value(pattern, text, flags=re.IGNORECASE | re.DOTALL):
    match = re.search(pattern, text, flags)
    if match:
        value = match.group(1).strip()
        value = re.sub(r"\s+", " ", value)
        return value
    return ""


def extract_pdf_text(file_bytes):
    pages_text = []
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page_number, page in enumerate(pdf.pages, start=1):
            text = page.extract_text()
            if text:
                pages_text.append(
                    f"\n--- PAGE {page_number} ---\n{text}"
                )
    return "\n".join(pages_text)


def extract_header(text):
    data = {}

    data["Quotation No."] = find_value(
        r"Quotation\s*No\.?\s*:\s*(.*?)"
        r"\s+Quotation\s*Date",
        text
    )

    data["Quotation Date"] = find_value(
        r"Quotation\s*Date\s*:\s*"
        r"([0-9]{1,2}[-/][0-9]{1,2}[-/][0-9]{2,4})",
        text
    )

    data["Quotation Validity(in days)"] = find_value(
        r"Quotation\s*Validity\s*\(\s*in\s*days\s*\)"
        r"\s*:\s*(\d+)",
        text
    )

    data["Inquiry Ref"] = find_value(
        r"Inquiry\s*Ref\s*:?\s*"
        r"([A-Za-z0-9\/_.\-]+)",
        text
    )

    data["Inquiry Date"] = find_value(
        r"Inquiry\s*Date\s*:?\s*"
        r"([0-9]{1,2}[-/][0-9]{1,2}[-/][0-9]{2,4})",
        text
    )

    data["Kind Attention"] = find_value(
        r"Kind\s*Attention\s*:?\s*"
        r"(.*?)(?:\(\)|\n|SUB\s*:)",
        text
    )

    data["SUB :-"] = find_value(
        r"SUB\s*:?\s*-?\s*"
        r"(.*?)(?:\n|Dear\s+Sir)",
        text
    )

    return data


def extract_to_address(text):
    lines = text.splitlines()
    to_lines = []
    capturing = False

    for line in lines:
        clean_line = line.strip()

        if re.match(r"^To\s*,?\s*$", clean_line, re.IGNORECASE):
            capturing = True
            continue

        same_line = re.match(
            r"^To\s*,\s*(.+)$",
            clean_line,
            re.IGNORECASE
        )

        if same_line:
            capturing = True
            value = same_line.group(1).strip()
            if value:
                to_lines.append(value)
            continue

        if capturing:
            if re.match(r"^From\s*,?\s*$", clean_line, re.IGNORECASE):
                break

            if re.match(
                r"^(GSTIN|PAN\s*NO|Phone\s*No|Phone|Fax\s*No|Email|"
                r"Inquiry\s*Ref|Inquiry\s*Date|Kind\s*Attention|SUB\s*:)",
                clean_line,
                re.IGNORECASE
            ):
                break

            if clean_line:
                to_lines.append(clean_line)

    cleaned_lines = []
    for line in to_lines:
        line = re.sub(r"\s+", " ", line).strip()
        if line:
            cleaned_lines.append(line)

    return "\n".join(cleaned_lines)


def extract_tax(text):
    data = {
        "Packing & Forward %": "",
        "Packing & Forward": "",
        "Freight on Sales": "",
        "SGST %": "",
        "SGST": "",
        "CGST %": "",
        "CGST": "",
        "IGST %": "",
        "IGST": "",
        "Total Tax Amount": "",
        "Grand Total": "",
        "Amount (in Words)": "",
    }

    patterns = [
        (
            "Packing & Forward",
            r"Packing\s*&\s*Forward\s*(\d+(?:\.\d+)?)\s*%\s+([\d,]+(?:\.\d+)?)"
        ),
        (
            "Freight on Sales",
            r"Freight\s+on\s+Sales\s+([\d,]+(?:\.\d+)?)"
        ),
        (
            "SGST",
            r"SGST\s*(\d+(?:\.\d+)?)\s*%\s+([\d,]+(?:\.\d+)?)"
        ),
        (
            "CGST",
            r"CGST\s*(\d+(?:\.\d+)?)\s*%\s+([\d,]+(?:\.\d+)?)"
        ),
        (
            "IGST",
            r"IGST\s*(\d+(?:\.\d+)?)\s*%\s+([\d,]+(?:\.\d+)?)"
        ),
    ]

    for key, pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            if key == "Packing & Forward":
                data["Packing & Forward %"] = float(match.group(1))
                data["Packing & Forward"] = clean_number(match.group(2))
            elif key in ("SGST", "CGST", "IGST"):
                data[f"{key} %"] = float(match.group(1))
                data[key] = clean_number(match.group(2))
            else:
                data[key] = clean_number(match.group(1))

    match = re.search(
        r"Total\s*Tax\s*Amount\s*:\s*([\d,]+(?:\.\d+)?)",
        text,
        re.IGNORECASE
    )
    if match:
        data["Total Tax Amount"] = clean_number(match.group(1))

    match = re.search(
        r"Grand\s*Total\s*:\s*([\d,]+(?:\.\d+)?)",
        text,
        re.IGNORECASE
    )
    if match:
        data["Grand Total"] = clean_number(match.group(1))

    match = re.search(
        r"Amount\s*\(\s*in\s*Words\s*\)\s*:\s*"
        r"(.*?)(?:\n\s*Page\s+\d+\s+of\s+\d+|\Z)",
        text,
        re.IGNORECASE | re.DOTALL
    )
    if match:
        data["Amount (in Words)"] = re.sub(
            r"\s+", " ", match.group(1)
        ).strip()

    return data


def extract_prepared_by(text):
    lines = text.splitlines()

    for i, line in enumerate(lines):
        if re.search(r"Prepared\s+By", line, re.IGNORECASE):

            same_line = re.search(
                r"Prepared\s+By\s*:\s*(.*?)(?:Authorized\s+By|$)",
                line,
                re.IGNORECASE
            )

            if same_line:
                value = same_line.group(1).strip()
                if value:
                    return value

            for j in range(i + 1, min(i + 5, len(lines))):
                candidate = lines[j].strip()

                if not candidate:
                    continue

                if re.search(r"Authorized\s+By", candidate, re.IGNORECASE):
                    continue

                if re.search(r"Page\s+\d+", candidate, re.IGNORECASE):
                    continue

                return candidate

    return ""


ITEM_START_PATTERN = re.compile(
    r"^\s*"
    r"(\d+)\s+"
    r"(.+?)\s+"
    r"(\d{6,10})\s+"
    r"([\d,]+(?:\.\d+)?)\s+"
    r"([\d,]+(?:\.\d+)?)\s+"
    r"([\d,]+(?:\.\d+)?)"
    r"\s*$"
)


def extract_items(text):
    lines = text.splitlines()
    items = []
    current_item = None
    in_item_section = False

    for line in lines:
        clean_line = line.strip()

        if not clean_line:
            continue

        if (
            "Item Code & Description" in clean_line
            or "Your Item Code & Description" in clean_line
        ):
            in_item_section = True
            continue

        if re.search(r"Total\s+Value\s*:", clean_line, re.IGNORECASE):
            if current_item:
                items.append(current_item)
                current_item = None
            in_item_section = False
            continue

        if re.search(r"Tax\s+Details", clean_line, re.IGNORECASE):
            if current_item:
                items.append(current_item)
                current_item = None
            in_item_section = False
            continue

        if in_item_section:
            match = ITEM_START_PATTERN.match(clean_line)

            if match:
                if current_item:
                    items.append(current_item)

                current_item = {
                    "Sr. No.": match.group(1),
                    "Item Code & Description": match.group(2).strip(),
                    "HSN CODE": match.group(3),
                    "Qty": clean_number(match.group(4)),
                    "SUOM": "",
                    "Unit Sales Price (INR)": clean_number(match.group(5)),
                    "Total (INR)": clean_number(match.group(6)),
                }
                continue

            if current_item:
                if re.search(
                    r"^(Sr\.?\s*No|Item Code|Your Item|HSN|CODE|Qty|SUOM|"
                    r"Unit Sales Price|Total)",
                    clean_line,
                    re.IGNORECASE
                ):
                    continue

                suom_match = re.search(
                    r"\b(NOS|PCS|PC|SET|SETS|KG|KGS|GRAM|GM|LTR|LITRE|"
                    r"LITER|MTR|METER|METRE|BOX|UNIT|EA|EACH)\b",
                    clean_line,
                    re.IGNORECASE
                )

                if suom_match:
                    if not current_item["SUOM"]:
                        current_item["SUOM"] = suom_match.group(1).upper()

                    remaining = re.sub(
                        r"\b(NOS|PCS|PC|SET|SETS|KG|KGS|GRAM|GM|LTR|LITRE|"
                        r"LITER|MTR|METER|METRE|BOX|UNIT|EA|EACH)\b",
                        "",
                        clean_line,
                        flags=re.IGNORECASE
                    ).strip()

                    if remaining:
                        current_item["Item Code & Description"] += "\n" + remaining
                else:
                    current_item["Item Code & Description"] += "\n" + clean_line

    if current_item:
        items.append(current_item)

    for item in items:
        item["Item Code & Description"] = re.sub(
            r"\n+", "\n", item["Item Code & Description"]
        ).strip()

    return items


def extraction_check(row):
    problems = []

    checks = {
        "Quotation No.": row.get("Quotation No.", ""),
        "To": row.get("To", ""),
        "Inquiry Ref": row.get("Inquiry Ref", ""),
        "Item Code & Description": row.get("Item Code & Description", ""),
        "HSN CODE": row.get("HSN CODE", ""),
        "Qty": row.get("Qty", ""),
        "Grand Total": row.get("Grand Total", ""),
        "Prepared By": row.get("Prepared By", ""),
    }

    if not str(row.get("Quotation Date", "")).strip():
        problems.append("Quotation Date")

    for field, value in checks.items():
        if not str(value).strip() or str(value).lower() == "nan":
            problems.append(field)

    return "CHECK: " + ", ".join(problems) if problems else "OK"


def extract_quotation(pdf_name, file_bytes):
    text = normalize_text(extract_pdf_text(file_bytes))

    header = extract_header(text)
    to_address = extract_to_address(text)
    tax = extract_tax(text)
    prepared_by = extract_prepared_by(text)
    items = extract_items(text)

    result = []

    if items:
        for item in items:
            row = {"Source PDF": pdf_name}
            row.update(header)
            row["To"] = to_address
            row.update(item)
            row.update(tax)
            row["Prepared By"] = prepared_by
            result.append(row)
    else:
        row = {"Source PDF": pdf_name}
        row.update(header)
        row["To"] = to_address

        for col in [
            "Sr. No.", "Item Code & Description", "HSN CODE",
            "Qty", "SUOM", "Unit Sales Price (INR)", "Total (INR)"
        ]:
            row[col] = ""

        row.update(tax)
        row["Prepared By"] = prepared_by
        result.append(row)

    return result


# ============================================================
# EXCEL CREATION
# ============================================================

def make_excel(df):
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(
            writer,
            sheet_name="Quotation Data",
            index=False
        )

        error_df = df[df["Extraction_Check"] != "OK"][
            ["Source PDF", "Extraction_Check"]
        ].copy()

        error_df.to_excel(
            writer,
            sheet_name="Errors",
            index=False
        )

    output.seek(0)

    wb = load_workbook(output)
    ws = wb["Quotation Data"]

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True
        )

    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(
                vertical="top",
                wrap_text=True
            )

    header_map = {
        cell.value: cell.column
        for cell in ws[1]
    }

    for col_name in ["Quotation Date", "Inquiry Date"]:
        col_num = header_map.get(col_name)
        if col_num:
            for row in ws.iter_rows(
                min_col=col_num,
                max_col=col_num,
                min_row=2
            ):
                for cell in row:
                    cell.number_format = "dd-mm-yyyy"

    numeric_formats = {
        "Qty": "0.0000",
        "Unit Sales Price (INR)": "#,##0.0000",
        "Total (INR)": "#,##0.00",
        "Packing & Forward %": "0.00",
        "Packing & Forward": "#,##0.0000",
        "Freight on Sales": "#,##0.0000",
        "SGST %": "0.00",
        "SGST": "#,##0.0000",
        "CGST %": "0.00",
        "CGST": "#,##0.0000",
        "IGST %": "0.00",
        "IGST": "#,##0.0000",
        "Total Tax Amount": "#,##0.00",
        "Grand Total": "#,##0.00",
    }

    for col_name, fmt in numeric_formats.items():
        col_num = header_map.get(col_name)
        if col_num:
            for row in ws.iter_rows(
                min_col=col_num,
                max_col=col_num,
                min_row=2
            ):
                for cell in row:
                    cell.number_format = fmt

    for column_cells in ws.columns:
        column_letter = get_column_letter(column_cells[0].column)
        max_length = 0

        for cell in column_cells:
            try:
                max_length = max(max_length, len(str(cell.value)))
            except Exception:
                pass

        ws.column_dimensions[column_letter].width = min(
            max(max_length + 2, 12), 50
        )

    if "To" in header_map:
        ws.column_dimensions[
            get_column_letter(header_map["To"])
        ].width = 45

    if "Item Code & Description" in header_map:
        ws.column_dimensions[
            get_column_letter(header_map["Item Code & Description"])
        ].width = 50

    if "Amount (in Words)" in header_map:
        ws.column_dimensions[
            get_column_letter(header_map["Amount (in Words)"])
        ].width = 50

    for row_number in range(2, ws.max_row + 1):
        ws.row_dimensions[row_number].height = 45

    final = io.BytesIO()
    wb.save(final)
    final.seek(0)
    return final.getvalue()


# ============================================================
# MYSQL FUNCTIONS
# ============================================================

def mysql_connect(host, port, user, password, database=None):
    if not MYSQL_AVAILABLE:
        raise RuntimeError(
            "mysql-connector-python is not installed. "
            "Run: pip install mysql-connector-python"
        )

    config = {
        "host": host,
        "port": int(port),
        "user": user,
        "password": password,
    }

    if database:
        config["database"] = database

    return mysql.connector.connect(**config)


def quote_identifier(name):
    return "`" + name.replace("`", "``") + "`"


def create_mysql_table(host, port, user, password, database, table_name):
    conn = mysql_connect(host, port, user, password)
    cur = conn.cursor()

    cur.execute(
        f"CREATE DATABASE IF NOT EXISTS {quote_identifier(database)}"
        f" CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci"
    )

    cur.close()
    conn.close()

    conn = mysql_connect(
        host, port, user, password, database
    )
    cur = conn.cursor()

    columns = """
        id BIGINT AUTO_INCREMENT PRIMARY KEY,
        source_pdf VARCHAR(500),
        quotation_no VARCHAR(255),
        quotation_date DATE NULL,
        customer_to TEXT,
        quotation_validity_days INT NULL,
        inquiry_ref VARCHAR(255),
        inquiry_date DATE NULL,
        kind_attention VARCHAR(500),
        subject_text TEXT,
        sr_no VARCHAR(50),
        item_code_description TEXT,
        hsn_code VARCHAR(50),
        qty DECIMAL(18,4) NULL,
        suom VARCHAR(50),
        unit_sales_price_inr DECIMAL(18,4) NULL,
        total_inr DECIMAL(18,2) NULL,
        packing_forward_percent DECIMAL(10,2) NULL,
        packing_forward DECIMAL(18,4) NULL,
        freight_on_sales DECIMAL(18,4) NULL,
        sgst_percent DECIMAL(10,2) NULL,
        sgst DECIMAL(18,4) NULL,
        cgst_percent DECIMAL(10,2) NULL,
        cgst DECIMAL(18,4) NULL,
        igst_percent DECIMAL(10,2) NULL,
        igst DECIMAL(18,4) NULL,
        total_tax_amount DECIMAL(18,2) NULL,
        grand_total DECIMAL(18,2) NULL,
        amount_in_words TEXT,
        prepared_by VARCHAR(255),
        extraction_check VARCHAR(1000),
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        UNIQUE KEY uq_source_item (
            source_pdf(191),
            quotation_no(191),
            sr_no(50)
        )
    """

    cur.execute(
        f"CREATE TABLE IF NOT EXISTS {quote_identifier(table_name)} ({columns})"
    )

    conn.commit()
    cur.close()
    conn.close()


def to_date(value):
    if pd.isna(value) or str(value).strip() == "":
        return None

    parsed = pd.to_datetime(value, dayfirst=True, errors="coerce")
    if pd.isna(parsed):
        return None

    return parsed.date()


def numeric_or_none(value):
    if pd.isna(value) or str(value).strip() == "":
        return None
    try:
        return float(value)
    except Exception:
        return None


def insert_dataframe_mysql(
    df,
    host,
    port,
    user,
    password,
    database,
    table_name
):
    conn = mysql_connect(
        host, port, user, password, database
    )
    cur = conn.cursor()

    sql = f"""
        INSERT INTO {quote_identifier(table_name)} (
            source_pdf,
            quotation_no,
            quotation_date,
            customer_to,
            quotation_validity_days,
            inquiry_ref,
            inquiry_date,
            kind_attention,
            subject_text,
            sr_no,
            item_code_description,
            hsn_code,
            qty,
            suom,
            unit_sales_price_inr,
            total_inr,
            packing_forward_percent,
            packing_forward,
            freight_on_sales,
            sgst_percent,
            sgst,
            cgst_percent,
            cgst,
            igst_percent,
            igst,
            total_tax_amount,
            grand_total,
            amount_in_words,
            prepared_by,
            extraction_check
        )
        VALUES (
            %s,%s,%s,%s,%s,%s,%s,%s,%s,%s,
            %s,%s,%s,%s,%s,%s,%s,%s,%s,%s,
            %s,%s,%s,%s,%s,%s,%s,%s,%s,%s
        )
        ON DUPLICATE KEY UPDATE
            quotation_date = VALUES(quotation_date),
            customer_to = VALUES(customer_to),
            quotation_validity_days = VALUES(quotation_validity_days),
            inquiry_ref = VALUES(inquiry_ref),
            inquiry_date = VALUES(inquiry_date),
            kind_attention = VALUES(kind_attention),
            subject_text = VALUES(subject_text),
            item_code_description = VALUES(item_code_description),
            hsn_code = VALUES(hsn_code),
            qty = VALUES(qty),
            suom = VALUES(suom),
            unit_sales_price_inr = VALUES(unit_sales_price_inr),
            total_inr = VALUES(total_inr),
            packing_forward_percent = VALUES(packing_forward_percent),
            packing_forward = VALUES(packing_forward),
            freight_on_sales = VALUES(freight_on_sales),
            sgst_percent = VALUES(sgst_percent),
            sgst = VALUES(sgst),
            cgst_percent = VALUES(cgst_percent),
            cgst = VALUES(cgst),
            igst_percent = VALUES(igst_percent),
            igst = VALUES(igst),
            total_tax_amount = VALUES(total_tax_amount),
            grand_total = VALUES(grand_total),
            amount_in_words = VALUES(amount_in_words),
            prepared_by = VALUES(prepared_by),
            extraction_check = VALUES(extraction_check)
    """

    inserted = 0

    for _, row in df.iterrows():
        values = (
            str(row.get("Source PDF", "")),
            str(row.get("Quotation No.", "")),
            to_date(row.get("Quotation Date")),
            str(row.get("To", "")),
            int(float(row["Quotation Validity(in days)"]))
            if str(row.get("Quotation Validity(in days)", "")).strip()
            and str(row.get("Quotation Validity(in days)", "")).lower() != "nan"
            else None,
            str(row.get("Inquiry Ref", "")),
            to_date(row.get("Inquiry Date")),
            str(row.get("Kind Attention", "")),
            str(row.get("SUB :-", "")),
            str(row.get("Sr. No.", "")),
            str(row.get("Item Code & Description", "")),
            str(row.get("HSN CODE", "")),
            numeric_or_none(row.get("Qty")),
            str(row.get("SUOM", "")),
            numeric_or_none(row.get("Unit Sales Price (INR)")),
            numeric_or_none(row.get("Total (INR)")),
            numeric_or_none(row.get("Packing & Forward %")),
            numeric_or_none(row.get("Packing & Forward")),
            numeric_or_none(row.get("Freight on Sales")),
            numeric_or_none(row.get("SGST %")),
            numeric_or_none(row.get("SGST")),
            numeric_or_none(row.get("CGST %")),
            numeric_or_none(row.get("CGST")),
            numeric_or_none(row.get("IGST %")),
            numeric_or_none(row.get("IGST")),
            numeric_or_none(row.get("Total Tax Amount")),
            numeric_or_none(row.get("Grand Total")),
            str(row.get("Amount (in Words)", "")),
            str(row.get("Prepared By", "")),
            str(row.get("Extraction_Check", "")),
        )

        cur.execute(sql, values)
        inserted += 1

    conn.commit()
    cur.close()
    conn.close()

    return inserted


def get_mysql_preview(
    host, port, user, password, database, table_name, limit=100
):
    conn = mysql_connect(
        host, port, user, password, database
    )

    query = (
        f"SELECT * FROM {quote_identifier(table_name)} "
        f"ORDER BY id DESC LIMIT {int(limit)}"
    )

    df = pd.read_sql(query, conn)
    conn.close()
    return df


# ============================================================
# MYSQL CONFIGURATION
# ============================================================
# MySQL settings are intentionally NOT shown in the Streamlit UI.
# Change these values here if your MySQL connection changes.

mysql_host = DEFAULT_HOST
mysql_port = DEFAULT_PORT
mysql_user = DEFAULT_USER
mysql_password = DEFAULT_PASSWORD

# Database and Table are shown on the frontend.
# Host, Port, Username and Password remain hidden.
with st.sidebar:
    st.header("🗄️ MySQL")
    mysql_database = st.text_input(
        "Database",
        value="quotation_db",
        key="mysql_database"
    )
    mysql_table = st.text_input(
        "Table",
        value="quotation_data",
        key="mysql_table"
    )


# ============================================================
# PDF UPLOAD
# ============================================================

st.subheader("1️⃣ Upload PDF Files")

# IMPORTANT:
# Do NOT assign to st.session_state["pdf_uploader"] after the
# file_uploader widget has been created. Streamlit raises
# StreamlitWidgetAlreadyInstantiatedError in that case.
#
# We use a versioned widget key. Incrementing the version creates
# a fresh uploader and safely clears the selected files.

if "uploader_version" not in st.session_state:
    st.session_state["uploader_version"] = 0

uploader_key = f"pdf_uploader_{st.session_state['uploader_version']}"

uploaded_files = st.file_uploader(
    "Select one or more PDFs",
    type=["pdf"],
    accept_multiple_files=True,
    key=uploader_key
)

upload_col1, upload_col2 = st.columns(2)

with upload_col1:
    extract_clicked = st.button(
        "🚀 Extract PDFs",
        type="primary",
        use_container_width=True,
        disabled=not bool(uploaded_files),
        key="extract_pdfs"
    )

with upload_col2:
    clear_upload_clicked = st.button(
        "🗑️ Clear Uploaded PDFs",
        use_container_width=True,
        key="clear_uploaded_pdfs"
    )

# Safely clear the uploader by changing its widget key.
# Never do: st.session_state["pdf_uploader"] = []
if clear_upload_clicked:
    st.session_state.pop("df", None)
    st.session_state.pop("errors", None)
    st.session_state.pop("mysql_df", None)
    st.session_state["uploader_version"] += 1
    st.rerun()

if uploaded_files:
    st.success(
        f"{len(uploaded_files)} PDF file(s) selected for bulk processing."
    )

    # Duplicate-safe based on file CONTENT, not filename.
    unique_files = []
    seen_hashes = set()

    for uploaded in uploaded_files:
        file_bytes = uploaded.getvalue()
        file_hash = hashlib.sha256(file_bytes).hexdigest()

        if file_hash not in seen_hashes:
            seen_hashes.add(file_hash)
            unique_files.append((uploaded.name, file_bytes))

    duplicate_count = len(uploaded_files) - len(unique_files)

    if duplicate_count:
        st.warning(
            f"{duplicate_count} duplicate PDF(s) were ignored automatically "
            "based on file content."
        )

    st.write(f"Unique PDFs to process: **{len(unique_files)}**")

    if extract_clicked:
        all_records = []
        errors = []

        progress = st.progress(0)
        status = st.empty()

        total_files = len(unique_files)

        for index, (file_name, file_bytes) in enumerate(
            unique_files, start=1
        ):
            status.write(
                f"Processing {index}/{total_files}: **{file_name}**"
            )

            try:
                records = extract_quotation(file_name, file_bytes)
                all_records.extend(records)
            except Exception as exc:
                errors.append({
                    "Source PDF": file_name,
                    "Error": str(exc)
                })

            progress.progress(index / total_files)

        status.success("Bulk PDF extraction completed.")

        df = pd.DataFrame(all_records)

        # Guarantee the exact output schema even when extraction
        # returns no records or a field is missing.
        for column in FINAL_COLUMNS:
            if column not in df.columns:
                df[column] = ""

        if not df.empty:
            df["Quotation Date"] = pd.to_datetime(
                df["Quotation Date"],
                dayfirst=True,
                errors="coerce"
            )

            df["Inquiry Date"] = pd.to_datetime(
                df["Inquiry Date"],
                dayfirst=True,
                errors="coerce"
            )

            df["Extraction_Check"] = df.apply(
                extraction_check,
                axis=1
            )

            df = df[FINAL_COLUMNS]
        else:
            # Keep the expected columns even when every PDF failed.
            df = df[FINAL_COLUMNS]

        st.session_state["df"] = df
        st.session_state["errors"] = errors

        st.success(
            f"Extraction completed. {len(unique_files)} unique PDF(s) "
            f"processed and {len(df)} Excel row(s) created."
        )


# ============================================================
# RESULTS
# ============================================================

if "df" in st.session_state:
    df = st.session_state["df"]
    errors = st.session_state.get("errors", [])

    st.divider()
    st.subheader("2️⃣ Extraction Result")

    result_col1, result_col2 = st.columns([1, 1])
    with result_col1:
        st.caption("Review the extracted data below.")
    with result_col2:
        clear_results_clicked = st.button(
            "🧹 Clear Results Table",
            use_container_width=True,
            key="clear_results"
        )

    if clear_results_clicked:
        st.session_state.pop("df", None)
        st.session_state.pop("errors", None)
        st.session_state.pop("mysql_df", None)
        st.rerun()

    c1, c2, c3, c4 = st.columns(4)

    c1.metric(
        "PDFs Processed",
        df["Source PDF"].nunique() if not df.empty else 0
    )

    c2.metric(
        "Excel Rows",
        len(df)
    )

    c3.metric(
        "OK Rows",
        int((df["Extraction_Check"] == "OK").sum())
        if not df.empty else 0
    )

    c4.metric(
        "Check Rows",
        int((df["Extraction_Check"] != "OK").sum())
        if not df.empty else 0
    )

    if not df.empty:
        st.dataframe(
            df,
            use_container_width=True,
            height=500
        )

        excel_bytes = make_excel(df)

        st.download_button(
            label="📥 Download Excel",
            data=excel_bytes,
            file_name="quotation_extracted.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    if errors:
        st.warning("Some PDFs could not be processed.")

        st.dataframe(
            pd.DataFrame(errors),
            use_container_width=True
        )

    # ========================================================
    # MYSQL
    # ========================================================

    st.divider()
    st.subheader("3️⃣ Create MySQL Table & Insert Data")

    if not MYSQL_AVAILABLE:
        st.error(
            "mysql-connector-python is not installed. "
            "Install it with: pip install mysql-connector-python"
        )

    if st.button(
        "🗄️ Create Table + Insert into MySQL",
        type="primary",
        use_container_width=True,
        key="mysql_insert"
    ):
        if not MYSQL_AVAILABLE:
            st.error(
                "mysql-connector-python is not installed. "
                "Run: pip install mysql-connector-python"
            )
        elif df.empty:
            st.error("There is no extracted data to insert.")
        else:
            try:
                with st.spinner("Connecting to MySQL..."):
                    create_mysql_table(
                        mysql_host,
                        mysql_port,
                        mysql_user,
                        mysql_password,
                        mysql_database,
                        mysql_table
                    )

                    inserted = insert_dataframe_mysql(
                        df,
                        mysql_host,
                        mysql_port,
                        mysql_user,
                        mysql_password,
                        mysql_database,
                        mysql_table
                    )

                st.success(
                    f"MySQL operation completed successfully. "
                    f"{inserted} row(s) inserted/updated in "
                    f"{mysql_database}.{mysql_table}"
                )

            except Exception as exc:
                st.error(
                    "MySQL operation failed."
                )
                st.exception(exc)

    # ========================================================
    # MYSQL PREVIEW
    # ========================================================

    if st.button(
        "🔄 Load MySQL Table Preview",
        use_container_width=True,
        key="mysql_preview"
    ):
        if not MYSQL_AVAILABLE:
            st.error(
                "mysql-connector-python is not installed. "
                "Run: pip install mysql-connector-python"
            )
        else:
            try:
                mysql_df = get_mysql_preview(
                    mysql_host,
                    mysql_port,
                    mysql_user,
                    mysql_password,
                    mysql_database,
                    mysql_table
                )

                st.session_state["mysql_df"] = mysql_df
                st.success(
                    f"Loaded {len(mysql_df)} row(s) from MySQL."
                )

            except Exception as exc:
                st.error(
                    "Could not read the MySQL table."
                )
                st.exception(exc)

    if "mysql_df" in st.session_state:
        st.subheader("MySQL Table Preview")
        st.dataframe(
            st.session_state["mysql_df"],
            use_container_width=True,
            height=400
        )


# ============================================================
# FOOTER
# ============================================================

st.divider()
st.caption(
    "PDF extraction logic is based on the quotation extraction "
    "structure provided by the user. One PDF can produce multiple "
    "rows when it contains multiple quotation items."
