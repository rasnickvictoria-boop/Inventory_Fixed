# Grand Welcome Web Importer — Fixed Build

This version removes `artifact-tool`, which caused Streamlit's dependency installer to fail.

## Update your Streamlit app

1. Delete the old repository files, or replace them with every file from this folder.
2. Confirm `requirements.txt` contains only:
   - streamlit
   - openpyxl
   - pypdf
   - requests
3. Commit and push the changes.
4. In Streamlit Community Cloud, open the app menu and select **Reboot app**.

The main file remains `app.py`.
