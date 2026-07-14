import streamlit as st
import openpyxl
from openpyxl.utils import get_column_letter, range_boundaries
import io

st.set_page_config(page_title="Flagged Data Extractor", layout="centered")
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Inter', sans-serif !important;
    }
    
    [data-testid="stFileUploader"] {
        margin-bottom: 0.5rem !important;
    }
    
    [data-testid="stButton"],
    [data-testid="stDownloadButton"] {
        margin-top: 0.5rem !important;
        margin-bottom: 1rem !important;
    }

    /* Size formatting for BOTH buttons */
    [data-testid="stButton"] button,
    [data-testid="stDownloadButton"] button {
        font-size: 18px !important;
        font-weight: 600 !important;
        padding: 0.75rem !important;
    }
    
    /* Specific styling to make the Download button Green */
    [data-testid="stDownloadButton"] button {
        background-color: #16a34a !important; /* Bold green */
        border-color: #16a34a !important;
        color: white !important;
    }
    
    /* Hover effect for the green button */
    [data-testid="stDownloadButton"] button:hover {
        background-color: #15803d !important; /* Slightly darker green on hover */
        border-color: #15803d !important;
    }

    /* Force the dialog box to the absolute center of the screen */
    [data-testid="stModal"] {
        display: flex !important;
        align-items: center !important;
        justify-content: center !important;
    }
    
    [data-testid="stModal"] > div {
        margin: auto !important;
    }
    </style>
""", unsafe_allow_html=True)
st.title("Flagged Data Extractor")
st.write("Upload your data files(which have colored duplicate inputs) and submission file(where all the duplicate data will be saved).Processed results will be available for download.")

def find_first_empty_row(ws, key_columns):
    row = 2
    while True:
        if all(ws.cell(row=row, column=col).value in (None, '') for col in key_columns):
            return row
        row += 1
        if row > ws.max_row + 1000:
            return row

def clean_cnic(value):
    if value is None or value == '': return None
    if isinstance(value, str):
        stripped = value.strip()
        return int(stripped) if stripped.isdigit() else stripped
    if isinstance(value, float):
        return int(value) if value.is_integer() else value
    return value

def write_cnic(ws, row, col, value):
    cell = ws.cell(row=row, column=col, value=clean_cnic(value))
    if isinstance(cell.value, int):
        cell.number_format = '0'

@st.dialog("✅ Processing Complete!")
def show_download_popup(file_data):
    st.write("Your files have been merged. Tap the button below to save the result.")
    st.download_button(
        label="⬇️ Download Completed File",
        data=file_data,
        file_name="Flagged_Data_Extractor_Result.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True
    )

# Wrap the uploaders and the button in a form
with st.form("upload_and_process_form", clear_on_submit=False):
    submission_file = st.file_uploader(
        "1. Upload Submission File ", 
        type=["xlsx"], 
        key="submission_key"
    )
    
    source_files = st.file_uploader(
        "2. Upload Source Files (Colored-Data)", 
        type=["xlsx"], 
        accept_multiple_files=True, 
        key="source_key"
    )
    
    # The submit button belongs to the form now
    submitted = st.form_submit_button("Process Files", type="primary", use_container_width=True)

# Execute logic only when the form is submitted
if submitted:
    if not submission_file or not source_files:
        st.error("Please upload both the submission file and at least one source file.")
    else:
        try:
            target_wb = openpyxl.load_workbook(submission_file)
            target_ws = target_wb.active
            
            # ... [Keep all your existing processing logic exactly the same here] ...
            
            output = io.BytesIO()
            target_wb.save(output)
            output.seek(0)
            
            show_download_popup(output)

        except Exception as e:
            st.error(f"An error occurred: {e}")

if st.button("Process Files", type="primary", use_container_width=True):
    if not submission_file or not source_files:
        st.error("Please upload both the submission file and at least one source file.")
    else:
        try:
            target_wb = openpyxl.load_workbook(submission_file)
            target_ws = target_wb.active
            target_headers = {str(cell.value).strip(): idx for idx, cell in enumerate(target_ws[1], 1) if cell.value}

            key_cols_for_emptiness = [
                target_headers[h] for h in
                ['District', 'Tehsil', 'UC-Name', 'Head of family CNIC', 'CHI CNIC', 'CHI Name', 'House code']
                if h in target_headers
            ]
            
            target_row = find_first_empty_row(target_ws, key_cols_for_emptiness)
            start_row = target_row

            base_mapping = {
                'District': 'District',
                'Tehsil': 'Tehsil',
                'UC': 'UC-Name',
                'HeadOfFamilyCNIC': 'Head of family CNIC'
            }

            for uploaded_source in source_files:
                wb = openpyxl.load_workbook(uploaded_source, data_only=True)
                ws = wb.active

                source_headers = {str(cell.value).strip(): idx for idx, cell in enumerate(ws[1], 1) if cell.value}

                for row in ws.iter_rows(min_row=2):
                    for i in range(1, 5):
                        cadre_col = f'Cadre {i}'
                        cnic_col = f'CNICofCadre {i}'
                        house_col = f'HouseCode {i}'

                        if house_col in source_headers:
                            house_cell_idx = source_headers[house_col] - 1
                            house_cell = row[house_cell_idx]

                            if house_cell.fill.patternType == 'solid' and house_cell.value not in (None, ''):
                                
                                for src, tgt in base_mapping.items():
                                    if src in source_headers and tgt in target_headers:
                                        value = row[source_headers[src] - 1].value
                                        if src == 'HeadOfFamilyCNIC':
                                            write_cnic(target_ws, target_row, target_headers[tgt], value)
                                        else:
                                            target_ws.cell(row=target_row, column=target_headers[tgt], value=value)

                                if cadre_col in source_headers and 'CHI Name' in target_headers:
                                    target_ws.cell(row=target_row, column=target_headers['CHI Name'],
                                                    value=row[source_headers[cadre_col] - 1].value)

                                if cnic_col in source_headers and 'CHI CNIC' in target_headers:
                                    write_cnic(target_ws, target_row, target_headers['CHI CNIC'],
                                               row[source_headers[cnic_col] - 1].value)

                                if 'House code' in target_headers:
                                    target_ws.cell(row=target_row, column=target_headers['House code'],
                                                    value=house_cell.value)

                                target_row += 1

            last_written_row = target_row - 1

            for table in target_ws.tables.values():
                min_col, min_row, max_col, max_row = range_boundaries(table.ref)
                new_max_row = max(max_row, last_written_row)
                table.ref = f"{get_column_letter(min_col)}{min_row}:{get_column_letter(max_col)}{new_max_row}"

            # Save Output
            output = io.BytesIO()
            target_wb.save(output)
            output.seek(0)
            
            rows_written = max(0, last_written_row - start_row + 1)
            
            # This triggers the pop-up window
            show_download_popup(output)

        except Exception as e:
            st.error(f"An error occurred: {e}")
