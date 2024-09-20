import streamlit as st
import pandas as pd
import io
import zipfile
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.formatting.rule import Rule
from openpyxl import load_workbook, Workbook


def create_report(all_reports):
    wb = Workbook()
    ws = wb.active
    ws.title = "Processing Report"

    headers = ["File Name", "Row", "Original Version", "New Version", "Explicit Words Found"]
    ws.append(headers)

    for file_name, report in all_reports.items():
        for item in report:
            parts = item.split(": ")
            row = parts[0].split(" ")[1]
            versions = parts[1].split(" became ")
            original_version = versions[0].strip("'")
            new_version = versions[1].split(" >>>")[0].strip("'")
            explicit_words = versions[1].split(" >>> ")[1]

            ws.append([file_name, row, original_version, new_version, explicit_words])

    return wb
def reset_app():
    # Clear all session state variables
    for key in list(st.session_state.keys()):
        del st.session_state[key]
    # Set a flag to reinitialize the file uploader
    st.session_state.reset_uploader = True
def process_version(version, has_explicit_content):
    if has_explicit_content and "Explicit" not in version:
        parts = version.split(', ')
        if parts[0] in ["Full", "Full Mix", "Main"]:
            parts[0] += " Explicit"
        else:
            parts[0] += " Explicit"
        return ", ".join(parts)
    return version  # Return original version if no changes are needed


def process_excel(df, search_words):
    report = []


    # Convert column names to lowercase for case-insensitive comparison
    df.columns = df.columns.str.lower()



    if 'lyrics' not in df.columns or 'version' not in df.columns:
        missing_columns = []
        if 'lyrics' not in df.columns:
            missing_columns.append('Lyrics')
        if 'version' not in df.columns:
            missing_columns.append('Version')
        return df, [
            f"Error: Required column(s) {', '.join(missing_columns)} not found in the Excel file. Available columns are: {', '.join(df.columns)}"]

    modified_df = df.copy()
    for index, row in df.iterrows():
        lyrics = str(row['lyrics']).lower()
        version = str(row['version'])
        has_explicit_content = any(word.lower() in lyrics for word in search_words)
        new_version = process_version(version, has_explicit_content)
        if new_version != version:
            report.append(
                f"Row {index + 2}: '{version}' became '{new_version}' >>> {[word for word in search_words if word.lower() in lyrics]}")
        modified_df.at[index, 'version'] = new_version
    return modified_df, report


def highlight_modified_cells(writer, sheet_name, modified_rows, version_col_name='Version'):
    workbook = writer.book
    worksheet = workbook[sheet_name]

    # Find the column index for 'Version' (case-insensitive)
    version_col = None
    for idx, cell in enumerate(worksheet[1], start=1):
        if cell.value and cell.value.lower() == version_col_name.lower():
            version_col = idx
            break

    if version_col:
        # Define the highlighting style
        yellow_fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')

        # Apply the highlighting to the modified cells
        for row in modified_rows:
            cell = worksheet.cell(row=row, column=version_col)
            cell.fill = yellow_fill
    else:
        print(f"Warning: '{version_col_name}' column not found. Highlighting not applied.")
def on_file_upload():
    st.session_state.file_uploaded = True

def main():
    st.markdown(
        "<h1 style='text-align: center; color: white;'><b>XPLICIT</b></h1>",
        unsafe_allow_html=True
    )

    # Initialize session state variables
    if 'reset_uploader' not in st.session_state:
        st.session_state.reset_uploader = False
    if 'file_uploader_key' not in st.session_state:
        st.session_state.file_uploader_key = "file_uploader_0"
    if 'file_uploaded' not in st.session_state:
        st.session_state.file_uploaded = False
    if 'uploaded_files' not in st.session_state:
        st.session_state.uploaded_files = None

    # Check if we need to reset the file uploader
    if st.session_state.reset_uploader:
        st.session_state.file_uploader_key = f"file_uploader_{pd.Timestamp.now().timestamp()}"
        st.session_state.reset_uploader = False
        st.session_state.uploaded_files = None

    # File uploader
    uploaded_files = st.file_uploader("Choose Excel files", type=["xlsx", "xls"],
                                      accept_multiple_files=True,
                                      on_change=on_file_upload,
                                      key=st.session_state.file_uploader_key)

    # Update session state when files are uploaded
    if uploaded_files and st.session_state.file_uploaded:
        st.session_state.uploaded_files = uploaded_files
        st.session_state.file_uploaded = False

    if 'custom_words' not in st.session_state:
        st.session_state.custom_words = []

    search_words = st.multiselect(
        "Enter words to search for in lyrics:",
        options=["shit", "bullshit", "shithead", "piss", "fuck", "cunt", "cocksucker", "motherfucker", "tits", "pussy", "asshole", "wog", "wop", "nigger", "kike", "gook", "gypsy", "faggot", "goddamn"] + list(set(st.session_state.custom_words)),
        default=["shit", "bullshit", "shithead", "piss", "fuck", "cunt", "cocksucker", "motherfucker", "tits", "pussy", "asshole", "wog", "wop", "nigger", "kike", "gook", "gypsy", "faggot", "goddamn"]
    )

    st.session_state.custom_words = list(set(st.session_state.custom_words + search_words))

    if 'processed_files' not in st.session_state:
        st.session_state.processed_files = None
    if 'all_reports' not in st.session_state:
        st.session_state.all_reports = None
    if 'processing_report' not in st.session_state:
        st.session_state.processing_report = None

    if st.session_state.uploaded_files and st.button("CHECK FOR EXPLICIT WORDS!"):
        all_reports = {}
        processed_files = []
        processing_report = []

        for uploaded_file in st.session_state.uploaded_files:
            try:
                df = pd.read_excel(uploaded_file)
                modified_df, report = process_excel(df, search_words)
                file_report = [f"Processing Report for {uploaded_file.name}:"]
                if not report:
                    file_report.append("No changes were made to the file.")
                else:
                    file_report.extend(report)
                    all_reports[uploaded_file.name] = report

                processing_report.extend(file_report)

                if not any(item.startswith("Error") for item in report):
                    buffer = io.BytesIO()
                    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                        modified_df.to_excel(writer, index=False, sheet_name='Sheet1')
                        highlight_explicit_cells(writer, 'Sheet1')
                    processed_files.append((uploaded_file.name, buffer.getvalue()))
            except Exception as e:
                error_message = f"An error occurred processing {uploaded_file.name}: {str(e)}"
                processing_report.append(error_message)
                st.error(error_message)

        # Store results in session state
        st.session_state.processed_files = processed_files
        st.session_state.all_reports = all_reports
        st.session_state.processing_report = processing_report

    # Display processing report if available
    if st.session_state.processing_report:
        for item in st.session_state.processing_report:
            st.write(item)

    # Display download buttons if data is available in session state
    if st.session_state.processed_files:
        if len(st.session_state.processed_files) == 1:
            st.download_button(
                label=f"DOWNLOAD UPDATED {st.session_state.processed_files[0][0]}",
                data=st.session_state.processed_files[0][1],
                file_name=f"modified_{st.session_state.processed_files[0][0]}",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                for file_name, file_content in st.session_state.processed_files:
                    zip_file.writestr(f"modified_{file_name}", file_content)

            st.download_button(
                label="DOWNLOAD UPDATED XLS",
                data=zip_buffer.getvalue(),
                file_name="modified_excel_files.zip",
                mime="application/zip"
            )

    if st.session_state.all_reports:
        report_wb = create_report(st.session_state.all_reports)
        report_buffer = io.BytesIO()
        report_wb.save(report_buffer)
        report_buffer.seek(0)

        st.download_button(
            label="DOWNLOAD REPORT",
            data=report_buffer,
            file_name="Explicit Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    # Add RESET button after all other buttons
    st.button("RESET", on_click=reset_app)

if __name__ == "__main__":
    main()
