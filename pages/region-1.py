# Region 1 Daily Report Converter
# Upload Excel, convert, and download

import base64
import sys
from io import BytesIO
from pathlib import Path

import streamlit as st

# Import conversion logic from Region 1
sys.path.insert(0, str(Path(__file__).parent.parent / "Region 1"))
from app import convert_daily_report

st.set_page_config(page_title="Region 1 Converter", page_icon="📊", layout="centered")

st.title("Region 1 Daily Report Converter")

uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])
report_date_override = st.date_input(
    "Report date (optional)",
    value=None,
    help="Use if filename doesn't contain 'tanggal DD Mon YYYY'",
)

convert_clicked = st.button("Convert")

if convert_clicked:
    if not uploaded_file:
        st.warning("Please upload an Excel file first.")
    else:
        with st.spinner("Converting..."):
            try:
                report_date = report_date_override if report_date_override else None
                df_merged, report_date = convert_daily_report(
                    uploaded_file, report_date=report_date
                )
                st.session_state["region1_df"] = df_merged
                st.session_state["region1_report_date"] = report_date
            except ValueError as e:
                st.error(str(e))
            except Exception as e:
                st.error(f"Conversion failed: {e}")

if "region1_df" in st.session_state:
    df = st.session_state["region1_df"]
    report_date = st.session_state["region1_report_date"]

    st.dataframe(df, use_container_width=True)

    col1, col2 = st.columns(2)

    with col1:
        excel_buffer = BytesIO()
        df.to_excel(excel_buffer, index=False)
        excel_buffer.seek(0)
        st.download_button(
            label="Download Excel",
            data=excel_buffer,
            file_name=f"{report_date:%Y-%m-%d}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    # Copy to clipboard button removed
