Steps:
1. Clone the repository
2. Open PowerShell inside the extracted folder.
3. Run:
   pip install -r requirements.txt
   streamlit run rd_report_simple_ui.py

Fixes:
- Cleaner borders in print preview
- No broken partial border lines in top section/total block
- Removed Excel freeze panes so sheet scrolling feels normal
- Scrollable selected-file list in the Streamlit UI
- LibreOffice detection includes common Windows paths
