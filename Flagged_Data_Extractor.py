import streamlit as st
import openpyxl
from openpyxl.utils import get_column_letter, range_boundaries
import io
import zipfile  # NEW IMPORT REQUIRED

# ... [Keep your CSS, UI setup, and helper functions the same] ...

with st.form("upload_and_process_form", clear_on_submit=False):
    submission_file = st.file_uploader("1. Upload Submission File ", type=["xlsx"], key="sub_key")
    source_files = st.file_uploader("2. Upload Source Files (Colored-Data)", type=["xlsx"], accept_multiple_files=True, key="src_key")
    submitted = st.form_submit_button("Process Files", type="primary", use_container_width=True)

if submitted:
    if not submission_file or not source_files:
        st.error("Please upload both the submission file and at least one source file.")
    else:
        try:
            # Check if submission file is locked/empty
            if submission_file.size == 0:
                st.error("❌ Submission file is 0 bytes. If it is open in another app, close it and try again.")
                st.stop()
                
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
            base_mapping = {'District': 'District', 'Tehsil': 'Tehsil', 'UC': 'UC-Name', 'HeadOfFamilyCNIC': 'Head of family CNIC'}

            for uploaded_source in source_files:
                # Check for locked/empty source files
                if uploaded_source.size == 0:
                    st.error(f"❌ Cannot read '{uploaded_source.name}'. It may be open in another app or is a cloud shortcut. Please close it and re-upload.")
                    st.stop()

                try:
                    wb = openpyxl.load_workbook(uploaded_source, data_only=True)
                except zipfile.BadZipFile:
                    st.error(f"❌ '{uploaded_source.name}' is unreadable or locked by Android. Please make sure the file is closed on your phone before uploading.")
                    st.stop()
                    
                ws = wb.active
                source_headers = {str(cell.value).strip(): idx for idx, cell in enumerate(ws[1], 1) if cell.value}

                # ... [Keep your exact row processing logic here] ...
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
                                    target_ws.cell(row=target_row, column=target_headers['CHI Name'], value=row[source_headers[cadre_col] - 1].value)

                                if cnic_col in source_headers and 'CHI CNIC' in target_headers:
                                    write_cnic(target_ws, target_row, target_headers['CHI CNIC'], row[source_headers[cnic_col] - 1].value)

                                if 'House code' in target_headers:
                                    target_ws.cell(row=target_row, column=target_headers['House code'], value=house_cell.value)

                                target_row += 1

            last_written_row = target_row - 1
            for table in target_ws.tables.values():
                min_col, min_row, max_col, max_row = range_boundaries(table.ref)
                new_max_row = max(max_row, last_written_row)
                table.ref = f"{get_column_letter(min_col)}{min_row}:{get_column_letter(max_col)}{new_max_row}"

            output = io.BytesIO()
            target_wb.save(output)
            output.seek(0)
            
            show_download_popup(output)

        except zipfile.BadZipFile:
             st.error("❌ The Submission file is unreadable or locked. Make sure it is closed on your phone.")
        except Exception as e:
            st.error(f"An error occurred: {e}")
