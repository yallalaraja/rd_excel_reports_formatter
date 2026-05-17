import io
import tempfile
import zipfile
from pathlib import Path
import html

import streamlit as st

from rd_report_formatter_final import format_report

st.set_page_config(
    page_title="RD Report Formatter",
    page_icon="📄",
    layout="centered",
)

st.markdown(
    """
    <style>
    .main-title {
        text-align: center;
        font-size: 30px;
        font-weight: 800;
        margin-bottom: 4px;
    }
    .sub-title {
        text-align: center;
        color: #555;
        font-size: 15px;
        margin-bottom: 24px;
    }
    .note-box {
        border: 1px solid #ddd;
        border-radius: 14px;
        padding: 14px 18px;
        background: #fafafa;
        margin-bottom: 16px;
    }
    .file-list-box {
        max-height: 240px;
        overflow-y: auto;
        border: 1px solid #ddd;
        border-radius: 10px;
        padding: 10px 14px;
        background: #fff;
        margin-top: 10px;
        margin-bottom: 12px;
    }
    .file-list-box ul {
        margin-top: 0;
        margin-bottom: 0;
        padding-left: 18px;
    }
    .small-note {
        color: #666;
        font-size: 13px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown('<div class="main-title">RD Installment Report Formatter</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="sub-title">Upload multiple RD Excel reports and download all formatted files in one ZIP.</div>',
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="note-box">
        <b>Output includes:</b> logo preserved, S.No column, leading-zero account numbers,
        formulas, borders, spacing, and print-friendly A4 setup.
    </div>
    """,
    unsafe_allow_html=True,
)

uploaded_files = st.file_uploader(
    "Upload RD report files",
    type=["xls", "xlsx"],
    accept_multiple_files=True,
    help="You can select multiple .xls or .xlsx files at once.",
)

if uploaded_files:
    st.success(f"{len(uploaded_files)} file(s) selected")

    file_items = "".join(f"<li>{html.escape(file.name)}</li>" for file in uploaded_files)
    st.markdown(
        f"""
        <div class="file-list-box">
            <b>Selected files</b>
            <ul>{file_items}</ul>
        </div>
        """,
        unsafe_allow_html=True,
    )

    if st.button("Format Files and Prepare ZIP", use_container_width=True):
        zip_buffer = io.BytesIO()
        errors = []
        success_files = []

        progress_bar = st.progress(0)
        status_text = st.empty()

        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            input_dir = temp_path / "input"
            output_dir = temp_path / "output"
            input_dir.mkdir()
            output_dir.mkdir()

            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                total_files = len(uploaded_files)

                for index, uploaded_file in enumerate(uploaded_files, start=1):
                    status_text.info(f"Formatting {uploaded_file.name} ...")

                    input_path = input_dir / uploaded_file.name
                    input_path.write_bytes(uploaded_file.getbuffer())

                    output_file_name = f"{input_path.stem}_formatted.xlsx"
                    output_path = output_dir / output_file_name

                    try:
                        format_report(input_path, output_path)
                        zip_file.write(output_path, arcname=output_file_name)
                        success_files.append(output_file_name)
                    except Exception as error:
                        errors.append(f"{uploaded_file.name}: {error}")

                    progress_bar.progress(index / total_files)

        zip_buffer.seek(0)
        status_text.empty()

        if success_files:
            st.success(f"Formatted {len(success_files)} file(s) successfully.")
            st.download_button(
                label="Download All Formatted Files as ZIP",
                data=zip_buffer,
                file_name="formatted_rd_reports.zip",
                mime="application/zip",
                use_container_width=True,
            )

        if errors:
            st.error("Some files failed:")
            for error in errors:
                st.write(f"• {error}")
else:
    st.info("Upload one or more Excel files to start.")
