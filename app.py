import streamlit as st
import pdfplumber
import pandas as pd
import re
from datetime import datetime
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment

# ==========================================
# PAGE SETTINGS
# ==========================================

st.set_page_config(
    page_title="PDF Extractor",
    layout="wide"
)

st.title("📄 GST Shipping Bill PDF Extractor")

st.write("Upload PDF files and extract invoice data into Excel.")

# ==========================================
# EXTRACT TEXT FROM PDF
# ==========================================

def extract_pdf_text(uploaded_file):

    full_text = ""

    try:

        with pdfplumber.open(uploaded_file) as pdf:

            for page in pdf.pages:

                text = page.extract_text()

                if text:
                    full_text += text + "\n"

    except Exception as e:

        st.error(f"PDF Error: {e}")

    return full_text

# ==========================================
# EXTRACT REQUIRED DATA
# ==========================================

def extract_data(text):

    data = {
        "Invoice No. (from Shipping Bill)": "",
        "Port Code (from Shipping Bill)": "",
        "Shipping Bill No. (from Shipping Bill)": "",
        "Shipping Bill Date (from Shipping Bill)": "",
        "Invoice Date (from Shipping Bill)": "",
        "TAXABLE VALUE": 0.0,
        "IGST TAX AMOUNT": 0.0,
        "TAXABLE VALUE (FOB+FREIGHT+INSURANCE)": 0.0,
        "LUT": ""
    }

    clean_text = re.sub(r'\s+', ' ', text)

    # ======================================
    # PORT CODE / SB NO / SB DATE
    # ======================================

    sb_match = re.search(
        r'INDIAN CUSTOMS EDI SYSTEM\s+([A-Z0-9]+)\s+(\d+)\s+([0-9A-Z\-]+)',
        clean_text
    )

    if sb_match:

        data["Port Code (from Shipping Bill)"] = sb_match.group(1)

        data["Shipping Bill No. (from Shipping Bill)"] = sb_match.group(2)

        data["Shipping Bill Date (from Shipping Bill)"] = sb_match.group(3)

    # ======================================
    # INVOICE NUMBER + DATE
    # ======================================

    inv_match = re.search(
        r'(JTIPL/\d{4}/\d{3})\s+(\d{2}/\d{2}/\d{4})',
        clean_text
    )

    if inv_match:

        data["Invoice No. (from Shipping Bill)"] = inv_match.group(1)

        original_date = inv_match.group(2)

        converted_date = datetime.strptime(
            original_date,
            "%d/%m/%Y"
        ).strftime("%d-%b-%y").upper()

        data["Invoice Date (from Shipping Bill)"] = converted_date

    # ======================================
    # FOB VALUE
    # ======================================

    fob_match = re.search(
        r'LM\s+([0-9.]+)',
        clean_text
    )

    fob_value = 0.0

    if fob_match:

        fob_value = float(fob_match.group(1))

        data["TAXABLE VALUE"] = round(
            fob_value,
            2
        )

    # ======================================
    # FREIGHT VALUE
    # ======================================

    freight_match = re.search(
        r'FREIGHT\s+([0-9.]+)',
        clean_text,
        re.IGNORECASE
    )

    freight_value = 0.0

    if freight_match:

        freight_value = float(
            freight_match.group(1)
        )

    # ======================================
    # INSURANCE VALUE
    # ======================================

    insurance_match = re.search(
        r'INSURANCE\s+([0-9.]+)',
        clean_text,
        re.IGNORECASE
    )

    insurance_value = 0.0

    if insurance_match:

        insurance_value = float(
            insurance_match.group(1)
        )

    # ======================================
    # TOTAL TAXABLE VALUE
    # ======================================

    total_taxable = (
        fob_value +
        freight_value +
        insurance_value
    )

    data[
        "TAXABLE VALUE (FOB+FREIGHT+INSURANCE)"
    ] = round(
        total_taxable,
        2
    )

    # ======================================
    # IGST AMOUNT
    # ======================================

    igst_amt_match = re.search(
        r'3\.CESS AMT\s+([0-9.]+)\s+([0-9.]+)',
        clean_text
    )

    if igst_amt_match:

        data["IGST TAX AMOUNT"] = round(
            float(
                igst_amt_match.group(2)
            ),
            2
        )

    # ======================================
    # LUT VALUE
    # ======================================

    lut_match = re.search(
        r'1\.MODE\s+2\.ASSESS\s+3\.EXMN\s+4\.JOBBING\s+5\.MEIS\s+6\.DBK\s+7\.RODTP\s+8\.LICENCE\s+9\.DFRC\s+10\.RE-EXP\s+11\.LUT\s+'
        r'(\S+)\s+(\S+)\s+(\S+)\s+(\S+)\s+(\S+)\s+(\S+)\s+(\S+)\s+(\S+)\s+(\S+)\s+(\S+)\s+(\S+)',
        text
    )

    if lut_match:

        data["LUT"] = lut_match.group(11).upper()

    return data

# ==========================================
# FILE UPLOAD
# ==========================================

uploaded_files = st.file_uploader(
    "Upload PDF Files",
    type=["pdf"],
    accept_multiple_files=True
)

# ==========================================
# EXTRACT BUTTON
# ==========================================

if st.button("Extract Data"):

    if not uploaded_files:

        st.warning("Please upload PDF files.")

    else:

        results = []
        duplicate_checker = set()

        progress_bar = st.progress(0)

        for index, uploaded_file in enumerate(uploaded_files):

            progress = (index + 1) / len(uploaded_files)

            progress_bar.progress(progress)

            st.write(f"Processing: {uploaded_file.name}")

            text = extract_pdf_text(uploaded_file)

            extracted = extract_data(text)

            invoice_no = extracted[
                "Invoice No. (from Shipping Bill)"
            ]

            if invoice_no in duplicate_checker and invoice_no != "":

                continue

            duplicate_checker.add(invoice_no)

            results.append(extracted)

        # ======================================
        # CREATE DATAFRAME
        # ======================================

        columns = [
            "S NO.",
            "Invoice No. (from Shipping Bill)",
            "Port Code (from Shipping Bill)",
            "Shipping Bill No. (from Shipping Bill)",
            "Shipping Bill Date (from Shipping Bill)",
            "Invoice Date (from Shipping Bill)",
            "TAXABLE VALUE",
            "IGST TAX AMOUNT",
            "TAXABLE VALUE (FOB+FREIGHT+INSURANCE)",
            "LUT"
        ]

        df = pd.DataFrame(results)

        df["SORT_DATE"] = pd.to_datetime(
            df["Invoice Date (from Shipping Bill)"],
            format="%d-%b-%y",
            errors="coerce"
        )

        df = df.sort_values(
            by="SORT_DATE",
            ascending=True
        )

        df = df.drop(columns=["SORT_DATE"])

        df.insert(
            0,
            "S NO.",
            range(1, len(df) + 1)
        )

        df = df[columns]

        # ======================================
        # CREATE EXCEL IN MEMORY
        # ======================================

        output = BytesIO()

        with pd.ExcelWriter(output, engine='openpyxl') as writer:

            df.to_excel(writer, index=False)

            workbook = writer.book
            worksheet = writer.sheets['Sheet1']

            # HEADER FORMAT
            for cell in worksheet[1]:

                cell.font = Font(
                    bold=True,
                    color="FF0000"
                )

                cell.alignment = Alignment(
                    horizontal="center",
                    vertical="center"
                )

            # CENTER ALIGNMENT
            for row in worksheet.iter_rows():

                for cell in row:

                    cell.alignment = Alignment(
                        horizontal="center",
                        vertical="center"
                    )

            # AUTO WIDTH
            for column in worksheet.columns:

                max_length = 0

                column_letter = column[0].column_letter

                for cell in column:

                    try:

                        if len(str(cell.value)) > max_length:

                            max_length = len(str(cell.value))

                    except:
                        pass

                adjusted_width = max_length + 5

                worksheet.column_dimensions[
                    column_letter
                ].width = adjusted_width

        output.seek(0)

        # ======================================
        # SHOW DATA
        # ======================================

        st.success("Extraction Completed!")

        st.dataframe(df)

        timestamp = datetime.now().strftime(
            "%Y%m%d_%H%M%S"
        )

        # ======================================
        # DOWNLOAD BUTTON
        # ======================================

        st.download_button(
            label="📥 Download Excel File",
            data=output,
            file_name=f"Shipping_Bill_Data_{timestamp}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )