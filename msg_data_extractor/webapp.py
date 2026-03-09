from __future__ import annotations

import sys
import tempfile
import warnings
from pathlib import Path

warnings.filterwarnings(
    "ignore",
    message=r"Pandas requires version '.*' or newer of 'numexpr'.*",
    category=UserWarning,
)

import streamlit as st

if __package__ in {None, ""}:
    sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
    from msg_data_extractor.bundle import (
        ZipValidationError,
        create_zip_bytes,
        extract_zip_bytes,
    )
    from msg_data_extractor.extractor import run
else:
    from .bundle import ZipValidationError, create_zip_bytes, extract_zip_bytes
    from .extractor import run


def _render_results(summary: dict):
    st.success(
        f"Processed {summary['total']} MSG files. Extracted photos from {summary['ok_count']} files."
    )

    if summary["results"]:
        rows = [
            {
                "Source Folder": result["source_folder"],
                "MSG Filename": result["msg_file"],
                "Applicant Name": result["applicant_name"],
                "Email": result["email"],
                "Phone": result["phone"],
                "Saved Photo File(s)": ", ".join(result["saved_paths"]),
                "Error": result["error"],
            }
            for result in summary["results"]
        ]
        st.dataframe(rows, use_container_width=True)


def render_app():
    st.set_page_config(page_title="MSG Photo Extractor", page_icon="📎", layout="wide")
    st.title("MSG Photo Extractor")
    st.write(
        "Upload a ZIP file containing Outlook MSG files and folders. The app will extract photo attachments and return a ZIP with the processed output and Excel log."
    )

    uploaded = st.file_uploader("Upload ZIP archive", type=["zip"])
    preserve_structure = st.checkbox("Keep folder structure in output", value=True)

    if st.button("Process ZIP", type="primary", disabled=uploaded is None):
        if uploaded is None:
            st.error("Choose a ZIP file first.")
            return

        try:
            with tempfile.TemporaryDirectory(prefix="msg-photo-web-") as temp_dir_str:
                temp_dir = Path(temp_dir_str)
                input_dir = temp_dir / "input"
                output_dir = temp_dir / "output"

                extracted = extract_zip_bytes(uploaded.getvalue(), input_dir)
                if not extracted:
                    st.error("The uploaded ZIP is empty.")
                    return

                summary = run(
                    msg_folder=input_dir,
                    output_folder=output_dir,
                    recursive=True,
                    preserve_structure=preserve_structure,
                )

                if summary["total"] == 0:
                    st.error("No .msg files were found in the uploaded ZIP.")
                    return

                output_zip = create_zip_bytes(output_dir)
                st.session_state["download_bytes"] = output_zip
                st.session_state["download_name"] = f"processed_{uploaded.name}"
                st.session_state["summary"] = summary
        except ZipValidationError as exc:
            st.error(str(exc))
        except Exception as exc:
            st.exception(exc)

    summary = st.session_state.get("summary")
    if summary:
        _render_results(summary)
        st.download_button(
            label="Download Processed ZIP",
            data=st.session_state["download_bytes"],
            file_name=st.session_state["download_name"],
            mime="application/zip",
            type="primary",
        )


def main():
    render_app()


if __name__ == "__main__":
    main()
