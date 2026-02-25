# Region 2 Daily Report Converter
# Upload Excel, convert, and download

import sys
from io import BytesIO
from pathlib import Path

import streamlit as st

# Import conversion logic from Region 2
sys.path.insert(0, str(Path(__file__).parent.parent / "Region 2"))
from app import convert_daily_report

st.set_page_config(page_title="Region 2 Converter", page_icon="📊", layout="centered")

st.title("Region 2 Daily Report Converter")

uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])
convert_clicked = st.button("Convert")

if convert_clicked:
    if not uploaded_file:
        st.warning("Please upload an Excel file first.")
    else:
        with st.spinner("Converting..."):
            try:
                df_merged, report_date = convert_daily_report(uploaded_file)
                st.session_state["region2_df"] = df_merged
                st.session_state["region2_report_date"] = report_date
            except ValueError as e:
                st.error(str(e))
            except Exception as e:
                st.error(f"Conversion failed: {e}")

if "region2_df" in st.session_state:
    df = st.session_state["region2_df"]
    report_date = st.session_state["region2_report_date"]

    filter_by_date = st.checkbox("Filter by date", value=False, key="region2_filter_check")
    filter_date = None
    if filter_by_date:
        filter_date = st.date_input(
            "Report date",
            value=report_date,
            key="region2_date_filter",
        )

    if filter_by_date and filter_date is not None:
        df_display = df[df["Report Date"].dt.date == filter_date]
        download_date = filter_date
    else:
        df_display = df
        download_date = report_date

    st.dataframe(df_display, use_container_width=True)

    excel_buffer = BytesIO()
    df_display.to_excel(excel_buffer, index=False)
    excel_buffer.seek(0)
    st.download_button(
        label="Download Excel",
        data=excel_buffer,
        file_name=f"{download_date:%Y-%m-%d}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
