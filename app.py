import streamlit as st
import pdfplumber
import openpyxl
import re
import io
import copy

st.set_page_config(page_title="Invoice Lot Number Enricher", page_icon="📋", layout="wide")

st.title("📋 Invoice Lot Number Enricher")
st.markdown("Upload a PDF invoice and an Excel file. The app will match items by description and append lot/certificate info to the Excel.")

TARGET_SHEET = "Edit - Posted Sales Invoice - "
DESC_COL = "Description"


def extract_lot_info_from_pdf(pdf_file):
    """Extract description -> lot info mapping from PDF."""
    lot_map = {}
    text = ""
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            text += page.extract_text() + "\n"

    # Pattern: item line followed by Br. serije lines
    # Match item lines like: "1. 12067 OLITOR CAPS 10/10MG 720 Kut 1,214.10 10 6 821,702.88"
    # Then collect subsequent Br. serije lines

    lines = text.split("\n")
    current_desc = None
    lot_lines = []

    item_re = re.compile(r"^\d+\.\s+\d+\s+(.+?)\s+[\d,]+\s+Kut\s+[\d,.]+")
    br_re = re.compile(r"(Br\. serije:.+)")

    i = 0
    while i < len(lines):
        line = lines[i].strip()
        m = item_re.match(line)
        if m:
            # Save previous
            if current_desc and lot_lines:
                lot_map[current_desc] = " | ".join(lot_lines)
            current_desc = m.group(1).strip()
            lot_lines = []
        elif current_desc:
            bm = br_re.search(line)
            if bm:
                lot_lines.append(bm.group(1).strip())
        i += 1

    # Save last
    if current_desc and lot_lines:
        lot_map[current_desc] = " | ".join(lot_lines)

    return lot_map


def enrich_excel(excel_file, lot_map):
    """Add lot info to Description column in target sheet."""
    wb = openpyxl.load_workbook(excel_file)

    if TARGET_SHEET not in wb.sheetnames:
        return None, f"Sheet '{TARGET_SHEET}' not found. Available: {wb.sheetnames}"

    ws = wb[TARGET_SHEET]

    # Find header row and Description column index
    header_row = None
    desc_col_idx = None
    for row in ws.iter_rows():
        for cell in row:
            if cell.value == DESC_COL:
                header_row = cell.row
                desc_col_idx = cell.column
                break
        if header_row:
            break

    if not desc_col_idx:
        return None, f"Column '{DESC_COL}' not found in sheet."

    matched = []
    unmatched = []

    for row in ws.iter_rows(min_row=header_row + 1):
        cell = row[desc_col_idx - 1]
        desc = cell.value
        if not desc:
            continue
        desc_str = str(desc).strip()
        if desc_str in lot_map:
            cell.value = f"{desc_str}\n{lot_map[desc_str]}"
            cell.alignment = openpyxl.styles.Alignment(wrap_text=True)
            matched.append(desc_str)
        else:
            unmatched.append(desc_str)

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out, matched, unmatched


col1, col2 = st.columns(2)

with col1:
    pdf_file = st.file_uploader("📄 Upload PDF Invoice", type=["pdf"])

with col2:
    excel_file = st.file_uploader("📊 Upload Excel File", type=["xlsx"])

if pdf_file and excel_file:
    with st.spinner("Extracting lot info from PDF..."):
        lot_map = extract_lot_info_from_pdf(pdf_file)

    st.subheader("📦 Lot Info Extracted from PDF")
    if lot_map:
        for desc, lots in lot_map.items():
            st.markdown(f"**{desc}**")
            st.code(lots, language=None)
    else:
        st.warning("No lot info found in the PDF.")

    if lot_map:
        with st.spinner("Enriching Excel file..."):
            result = enrich_excel(excel_file, lot_map)

        if result[0] is None:
            st.error(result[1])
        else:
            out_bytes, matched, unmatched = result

            st.subheader("✅ Results")
            col3, col4 = st.columns(2)
            with col3:
                st.success(f"**{len(matched)} items matched and updated:**")
                for m in matched:
                    st.markdown(f"- {m}")
            with col4:
                if unmatched:
                    st.warning(f"**{len(unmatched)} items not matched in PDF:**")
                    for u in unmatched:
                        st.markdown(f"- {u}")
                else:
                    st.success("All items matched!")

            st.download_button(
                label="⬇️ Download Enriched Excel",
                data=out_bytes,
                file_name=f"enriched_{excel_file.name}",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
