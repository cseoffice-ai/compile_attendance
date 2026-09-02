import io
import re
import pandas as pd
import streamlit as st
import pdfplumber
from docx import Document
from rapidfuzz import fuzz

st.set_page_config(page_title="Attendance Compiler", layout="wide")

st.title("📊 Automated Attendance Compiler")
st.write("Upload multiple Word (.docx), PDF (.pdf), or Excel (.xlsx) attendance records to aggregate total lectures and present counts.")

# --- FILE EXTRACTION HELPERS ---

def extract_from_docx(file):
    """Extracts table rows and text from Word files."""
    doc = Document(file)
    rows = []
    for table in doc.tables:
        for row in table.rows:
            rows.append([cell.text.strip() for cell in row.cells])
    if not rows:
        for p in doc.paragraphs:
            if p.text.strip():
                tokens = re.split(r'\||\t|\s{2,}', p.text.strip())
                rows.append([t.strip() for t in tokens if t.strip()])
    return rows

def extract_from_pdf(file):
    """Extracts table rows and text from PDF files."""
    rows = []
    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                for row in table:
                    if row:
                        cleaned_row = [str(cell).strip() if cell is not None else "" for cell in row]
                        rows.append(cleaned_row)
            # Fallback if no structured tables found
            if not tables:
                text = page.extract_text()
                if text:
                    for line in text.split('\n'):
                        tokens = re.split(r'\||\t|\s{2,}', line.strip())
                        if tokens:
                            rows.append([t.strip() for t in tokens if t.strip()])
    return rows

def parse_to_dataframe(file):
    """Parses Excel, Word, or PDF files into a structured DataFrame."""
    file_name = file.name.lower()
    
    if file_name.endswith(('.xlsx', '.xls')):
        df = pd.read_excel(file)
        return clean_extracted_df(df)
        
    raw_rows = []
    if file_name.endswith('.docx'):
        raw_rows = extract_from_docx(file)
    elif file_name.endswith('.pdf'):
        raw_rows = extract_from_pdf(file)

    if not raw_rows:
        return None

    max_cols = max(len(r) for r in raw_rows)
    padded_rows = [r + [''] * (max_cols - len(r)) for r in raw_rows]
    df = pd.DataFrame(padded_rows)
    return clean_extracted_df(df)

def clean_extracted_df(df):
    """Detects and normalizes Roll/ID, Name, Total Classes, and Attendance columns."""
    df_clean = pd.DataFrame()
    id_col, name_col, total_col, att_col = None, None, None, None
    
    for col in df.columns:
        col_series = df[col].astype(str).str.upper()
        if not id_col and col_series.str.contains('ID|ROLL|S.NO|ENROLLMENT').any():
            id_col = col
        elif not name_col and col_series.str.contains('NAME|STUDENT').any():
            name_col = col
        elif not total_col and col_series.str.contains('TOTAL').any():
            total_col = col
        elif not att_col and col_series.str.contains('ATTENDANCE|PRESENT|HELD').any():
            att_col = col

    cols = list(df.columns)
    if len(cols) >= 4:
        id_col = id_col or cols[0]
        name_col = name_col or cols[1]
        total_col = total_col or cols[2]
        att_col = att_col or cols[3]
    elif len(cols) == 3:
        name_col = name_col or cols[0]
        total_col = total_col or cols[1]
        att_col = att_col or cols[2]

    if name_col:
        df_clean['NAME'] = df[name_col].astype(str).str.strip().str.upper()
    if id_col:
        df_clean['ID'] = df[id_col].astype(str).str.strip()
    else:
        df_clean['ID'] = ""
        
    df_clean['TOTAL'] = pd.to_numeric(df[total_col], errors='coerce').fillna(0) if total_col else 0
    df_clean['ATTENDANCE'] = pd.to_numeric(df[att_col], errors='coerce').fillna(0) if att_col else 0

    # Filter invalid/header rows
    df_clean = df_clean[~df_clean['NAME'].str.contains('NAME|STUDENT|SUBJECT|TOTAL', na=False)]
    df_clean = df_clean[df_clean['NAME'] != '']
    return df_clean

def match_and_aggregate(dfs):
    """Combines records using exact ID match or fuzzy name match."""
    master_records = []

    for df in dfs:
        for _, row in df.iterrows():
            curr_id = str(row['ID']).strip() if row['ID'] else ""
            curr_name = str(row['NAME']).strip().upper()
            curr_total = float(row['TOTAL'])
            curr_att = float(row['ATTENDANCE'])

            matched = False

            # Match by ID
            if curr_id and curr_id not in ["NAN", ""]:
                for rec in master_records:
                    if rec['ID'] == curr_id:
                        rec['TOTAL'] += curr_total
                        rec['ATTENDANCE'] += curr_att
                        matched = True
                        break

            # Fuzzy Match by Name (First, Middle, Last token matching)
            if not matched and curr_name:
                for rec in master_records:
                    score = fuzz.token_set_ratio(curr_name, rec['NAME'])
                    if score >= 85:
                        rec['TOTAL'] += curr_total
                        rec['ATTENDANCE'] += curr_att
                        if not rec['ID'] and curr_id:
                            rec['ID'] = curr_id
                        matched = True
                        break

            if not matched:
                master_records.append({
                    'ID': curr_id,
                    'NAME': curr_name,
                    'TOTAL': curr_total,
                    'ATTENDANCE': curr_att
                })

    result_df = pd.DataFrame(master_records)
    if not result_df.empty:
        result_df['PERCENTAGE (%)'] = (result_df['ATTENDANCE'] / result_df['TOTAL'] * 100).round(2).fillna(0)
    return result_df

# --- STREAMLIT UI ---

uploaded_files = st.file_uploader(
    "Upload Word, PDF, or Excel Files", 
    type=["docx", "pdf", "xlsx", "xls"], 
    accept_multiple_files=True
)

if uploaded_files:
    parsed_dfs = []
    
    for file in uploaded_files:
        parsed_df = parse_to_dataframe(file)
        if parsed_df is not None and not parsed_df.empty:
            parsed_dfs.append(parsed_df)
            st.success(f"Successfully processed: {file.name}")
        else:
            st.warning(f"Could not extract valid attendance table from: {file.name}")

    if parsed_dfs:
        st.subheader("Compiled Attendance Summary")
        final_df = match_and_aggregate(parsed_dfs)
        st.dataframe(final_df, use_container_width=True)

        excel_buffer = io.BytesIO()
        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
            final_df.to_excel(writer, index=False, sheet_name='Compiled Attendance')
        
        st.download_button(
            label="📥 Download Compiled Excel Report",
            data=excel_buffer.getvalue(),
            file_name="Compiled_Attendance_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
