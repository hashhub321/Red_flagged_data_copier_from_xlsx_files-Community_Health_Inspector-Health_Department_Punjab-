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
        margin-bottom: 2rem !important;
    }
    
    [data-testid="stButton"] {
        margin-top: 1rem !important;
        margin-bottom: 2rem !important;
    }

    /* Makes the button taller and the text larger */
    [data-testid="stButton"] button {
        font-size: 20px !important;
        font-weight: 800 !important;
        padding: 1.5rem !important;
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

submission_file = st.file_uploader("1. Upload Submission File ", type=["xlsx"])
source_files = st.file_uploader("2. Upload Source Files (Colored-Data)", type=["xlsx"], accept_multiple_files=True)

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

            output = io.BytesIO()
            target_wb.save(output)
            output.seek(0)
            
            rows_written = max(0, last_written_row - start_row + 1)
            st.success(f"Data procedure complete. {rows_written} lines written to submission file.")

            st.download_button(
                label="Download completed submission file",
                data=output,
                file_name="Flagged_Data_Extractor_Result.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"An error occurred: {e}")
