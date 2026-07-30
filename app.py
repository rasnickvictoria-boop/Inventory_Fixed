from __future__ import annotations

import tempfile
from datetime import datetime
from pathlib import Path

import streamlit as st

from import_engine import download_pdf, import_report


APP_TITLE = "Grand Welcome Breezeway Importer"
DEFAULT_WORKBOOK = Path(__file__).parent / "Grand_Welcome_Phase_1_Master_Workbook.xlsx"


st.set_page_config(page_title=APP_TITLE, page_icon="🏔️", layout="wide")

st.title("🏔️ Grand Welcome Breezeway Importer")
st.write(
    "Upload Breezeway quarterly-inspection PDFs—or paste public PDF links—"
    "and download an updated copy of the master workbook."
)

with st.expander("How to use this app"):
    st.markdown(
        """
1. Upload your latest master workbook, or leave it blank to use the included template.
2. Upload one or more downloaded Breezeway PDFs.
3. You may also paste one public PDF URL per line.
4. Select **Process reports**.
5. Review the results and download the updated workbook.
        """
    )

left, right = st.columns(2, gap="large")

with left:
    st.subheader("1. Master workbook")
    workbook_upload = st.file_uploader(
        "Upload your latest master workbook",
        type=["xlsx"],
        help="Leave blank to use the included Phase 1 workbook.",
    )
    if workbook_upload is None:
        st.success("Using the included Phase 1 master workbook.")
    else:
        st.success(f"Using: {workbook_upload.name}")

    st.subheader("2. PDF reports")
    pdf_uploads = st.file_uploader(
        "Upload one or more Breezeway PDFs",
        type=["pdf"],
        accept_multiple_files=True,
    )

with right:
    st.subheader("3. Optional PDF links")
    url_text = st.text_area(
        "Paste one public Breezeway PDF URL per line",
        placeholder="https://reports.breezeway.io/task/maintenance/v4/....pdf",
        height=160,
    )
    force = st.checkbox(
        "Force duplicate imports",
        value=False,
        help="Normally matching property/date Inspection IDs are skipped.",
    )

    process = st.button(
        "Process reports",
        type="primary",
        use_container_width=True,
        disabled=not pdf_uploads and not url_text.strip(),
    )

if process:
    urls = [line.strip() for line in url_text.splitlines() if line.strip()]

    with tempfile.TemporaryDirectory(prefix="gw_streamlit_") as temp_name:
        temp_dir = Path(temp_name)
        workbook_path = temp_dir / "Grand_Welcome_Updated.xlsx"

        if workbook_upload is None:
            workbook_path.write_bytes(DEFAULT_WORKBOOK.read_bytes())
        else:
            workbook_path.write_bytes(workbook_upload.getvalue())

        jobs = []
        for uploaded in pdf_uploads or []:
            pdf_path = temp_dir / uploaded.name
            pdf_path.write_bytes(uploaded.getvalue())
            jobs.append((pdf_path, uploaded.name))

        download_errors = []
        downloaded_paths = []
        for url in urls:
            try:
                path = download_pdf(url)
                downloaded_paths.append(path)
                jobs.append((path, url))
            except Exception as exc:
                download_errors.append({
                    "Source": url,
                    "Property": "",
                    "Inspection Date": "",
                    "Inspection ID": "",
                    "Status": "Error",
                    "Details": f"Could not download PDF: {exc}",
                })

        results = list(download_errors)
        progress = st.progress(0, text="Preparing reports…")

        for index, (pdf_path, source) in enumerate(jobs, start=1):
            try:
                result = import_report(pdf_path, source, workbook_path, force=force)
                results.append({
                    "Source": source,
                    "Property": result["property"],
                    "Inspection Date": result["date"].strftime("%m/%d/%Y"),
                    "Inspection ID": result["inspection_id"],
                    "Status": result["status"],
                    "Details": result["message"],
                })
            except Exception as exc:
                results.append({
                    "Source": source,
                    "Property": "",
                    "Inspection Date": "",
                    "Inspection ID": "",
                    "Status": "Error",
                    "Details": str(exc),
                })

            progress.progress(
                index / max(len(jobs), 1),
                text=f"Processing report {index} of {len(jobs)}…",
            )

        for path in downloaded_paths:
            try:
                path.unlink(missing_ok=True)
            except Exception:
                pass

        progress.progress(1.0, text="Finished")

        imported = sum(row["Status"] == "Imported" for row in results)
        duplicates = sum(row["Status"] == "Duplicate" for row in results)
        errors = sum(row["Status"] == "Error" for row in results)

        st.divider()
        st.subheader("Import results")
        col1, col2, col3 = st.columns(3)
        col1.metric("Imported", imported)
        col2.metric("Duplicates skipped", duplicates)
        col3.metric("Errors", errors)

        st.dataframe(results, use_container_width=True, hide_index=True)

        timestamp = datetime.now().strftime("%Y-%m-%d_%H%M")
        st.download_button(
            "Download updated master workbook",
            data=workbook_path.read_bytes(),
            file_name=f"Grand_Welcome_Master_Updated_{timestamp}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )

        if errors:
            st.warning("Some reports need review. See the Details column above.")
        else:
            st.success("The reports finished processing.")

st.caption(
    "Phase 1 supports the Breezeway Maintenance Report v4 PDF layout. "
    "Keep a separate backup of your current workbook."
)
