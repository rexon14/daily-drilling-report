"""
Daily Report Transformer for Zone 8
Transforms daily report .txt files into structured .xlsx files.

Input: daily-report/{date}.txt (Indonesian format - Laporan Harian DWI PHM)
Output: export-raw/{date}.xlsx and export-final/{date}.xlsx

Usage:
    python app.py <date>

Example:
    python app.py 2026-01-03
"""

import sys
import os
import re
from datetime import datetime, timedelta
import pandas as pd
import numpy as np


def parse_txt_file(path):
    """
    Parse daily drilling report .txt file and extract well data.
    
    Args:
        path: Path to the .txt file
    
    Returns:
        DataFrame with well information
    """
    with open(path, encoding='utf-8') as f:
        text_content = f.read()
    
    # Truncate content at WELL INTERVENTION keyword
    keyword = "WELL INTERVENTION"
    keyword_position = text_content.upper().find(keyword.upper())
    if keyword_position != -1:
        text_content = text_content[:keyword_position].rstrip()
    
    # Process lines
    lines = [l.rstrip("\n").replace("*", "").replace("_", "") for l in text_content.splitlines()]
    
    records = []
    current = {}
    i = 0
    
    while i < len(lines):
        line = lines[i].strip()
        
        # Detect new well entry (starts with number and dot)
        m = re.match(r"^(\d+)\.\s*", line)
        if m:
            if current:
                records.append(current)
            current = {
                "No": int(m.group(1)),
                "Nama Sumur": None,
                "Nama Rig": None,
                "Hari ke": None,
                "Kedalaman (mMD)": None,
                "Progress (mMD)": None,
                "AFE": None,
                "Realisasi Biaya": None,
                "EMD": None,
                "Summary Report": None,
                "Current Status": None,
                "Next Plan": None
            }
            i += 1
            continue

        if not current:
            i += 1
            continue

        # Extract field values
        if line.startswith("Nama Sumur"):
            current["Nama Sumur"] = line.split(":", 1)[1].strip()
        elif line.startswith("Nama Rig"):
            current["Nama Rig"] = line.split(":", 1)[1].strip()
        elif line.startswith("Hari ke"):
            current["Hari ke"] = int(line.split(":", 1)[1].strip())
        elif line.startswith("Kedalaman"):
            num = re.search(r"([\d\.]+)", line)
            if num:
                val = num.group(1)
                current["Kedalaman (mMD)"] = float(val) if "." in val else int(val)
        elif line.startswith("Progres"):
            num = re.search(r"([\d\.]+)", line)
            if num:
                val = num.group(1)
                current["Progress (mMD)"] = float(val) if "." in val else int(val)
        elif line.startswith("AFE"):
            num = re.sub(r"[^\d\.]", "", line.split(":", 1)[1])
            current["AFE"] = int(float(num))
        elif line.lower().startswith("realisasi biaya"):
            m2 = re.search(r"([\d,\.]+)", line)
            if m2:
                num = m2.group(1).replace(",", "")
                current["Realisasi Biaya"] = int(float(num))
        elif line.startswith("EMD"):
            val = line.split(":", 1)[1].strip()
            current["EMD"] = pd.to_datetime(val, dayfirst=True, errors="coerce")
        elif line.lower().startswith("summary report"):
            j = i + 1
            while j < len(lines) and not lines[j].strip():
                j += 1
            if j < len(lines):
                current["Summary Report"] = lines[j].strip()
                i = j
        elif line.lower().startswith("current status"):
            j = i + 1
            while j < len(lines) and not lines[j].strip():
                j += 1
            if j < len(lines):
                current["Current Status"] = lines[j].strip()
                i = j
        elif line.lower().startswith("next plan"):
            j = i + 1
            while j < len(lines) and not lines[j].strip():
                j += 1
            if j < len(lines):
                current["Next Plan"] = lines[j].strip()
                i = j

        i += 1

    if current:
        records.append(current)

    cols = [
        "No", "Nama Sumur", "Nama Rig", "Hari ke",
        "Kedalaman (mMD)", "Progress (mMD)", "AFE", "Realisasi Biaya",
        "EMD", "Summary Report", "Current Status", "Next Plan"
    ]
    return pd.DataFrame(records, columns=cols)


def clean_text_fields(df):
    """
    Clean text fields by removing leading special characters.
    
    Args:
        df: DataFrame to clean
    """
    text_fields = ['Summary Report', 'Current Status', 'Next Plan']
    for field in text_fields:
        if field in df.columns:
            df[field] = df[field].astype(str).str.replace(r'^[\-\:=]+\s*', '', regex=True)
            df[field] = df[field].replace('nan', np.nan)


def transform_txt_to_raw(date_str, base_path="."):
    """
    Transform daily report .txt file to export-raw .xlsx file.
    
    Args:
        date_str: Date string in format "YYYY-MM-DD"
        base_path: Base path containing daily-report and export-raw folders
    
    Returns:
        Path to the created Excel file
    """
    txt_file_path = os.path.join(base_path, "daily-report", f"{date_str}.txt")
    xlsx_file_path = os.path.join(base_path, "export-raw", f"{date_str}.xlsx")
    
    if not os.path.exists(txt_file_path):
        raise FileNotFoundError(f"Daily report file not found: {txt_file_path}")
    
    os.makedirs(os.path.dirname(xlsx_file_path), exist_ok=True)
    
    print(f"Reading daily report from: {txt_file_path}")
    df = parse_txt_file(txt_file_path)
    
    if df.empty:
        raise ValueError("No well data found in the report file")
    
    print(f"Exporting to Excel: {xlsx_file_path}")
    df.to_excel(xlsx_file_path, index=False, engine='openpyxl')
    
    print(f"Successfully transformed {len(df)} wells to Excel format")
    return xlsx_file_path


def transform_raw_to_final(date_str, base_path="."):
    """
    Transform export-raw .xlsx to export-final .xlsx with standardized columns.
    
    Args:
        date_str: Date string in format "YYYY-MM-DD"
        base_path: Base path containing export-raw and export-final folders
    
    Returns:
        Path to the created Excel file
    """
    raw_file_path = os.path.join(base_path, "export-raw", f"{date_str}.xlsx")
    final_file_path = os.path.join(base_path, "export-final", f"{date_str}.xlsx")
    
    if not os.path.exists(raw_file_path):
        raise FileNotFoundError(f"Export-raw file not found: {raw_file_path}")
    
    os.makedirs(os.path.dirname(final_file_path), exist_ok=True)
    
    print(f"Reading export-raw file from: {raw_file_path}")
    df = pd.read_excel(raw_file_path, engine='openpyxl')
    
    # Select required columns
    cols = ['Nama Sumur', 'Nama Rig', 'Summary Report', 'Current Status', 'Next Plan']
    df = df[cols]
    
    # Add standard columns
    df = df.assign(**{
        "Flag": "INC",
        "Region": "Region 3",
        "Zone": "Zone 8",
        "APH": "PHM",
        "Well Name [2]": np.nan,
        "Well Type": "Development",
        "Location": "Offshore",
        "Spud Date": np.nan,
        "Release Date": np.nan,
        "Status": np.nan,
        "Status Code [1]": np.nan,
        "Status Code [2]": np.nan
    })
    
    # Rename columns
    df.rename(columns={
        "Nama Sumur": "Well Name",
        "Nama Rig": "Rig Name",
    }, inplace=True)
    
    # Fill Well Name [2] with Well Name
    df['Well Name [2]'] = df['Well Name [2]'].fillna(df['Well Name'])
    
    # Clean text fields
    clean_text_fields(df)
    
    # Add date columns
    report_ts = pd.to_datetime(date_str)
    df['Report Date'] = report_ts
    df['Operation Date'] = report_ts - pd.Timedelta(days=1)
    df['Report Date'] = df['Report Date'].dt.date
    df['Operation Date'] = df['Operation Date'].dt.date
    
    # Reorder columns
    df['Rig Name'] = df['Rig Name'].astype(str)
    cols_order = [
        "Flag", "Region", "Zone", "APH", "Rig Name", "Well Name", "Well Name [2]",
        "Well Type", "Location", "Spud Date", "Release Date", "Status",
        "Status Code [1]", "Status Code [2]", "Summary Report", "Current Status",
        "Next Plan", "Report Date", "Operation Date"
    ]
    df = df[cols_order]
    
    # Sort by Rig Name (case-insensitive)
    df.sort_values(
        by='Rig Name',
        key=lambda s: s.str.lower(),
        ascending=True,
        inplace=True
    )
    df.reset_index(drop=True, inplace=True)
    
    print(f"Exporting final version to Excel: {final_file_path}")
    df.to_excel(final_file_path, index=False, engine='openpyxl')
    
    print(f"Successfully transformed {len(df)} wells to final Excel format")
    
    # Copy to clipboard (values only, no header)
    print("Copying data to clipboard...")
    try:
        df.to_clipboard(index=False, header=False)
        print("Data copied to clipboard successfully!")
    except Exception as e:
        print(f"Could not copy to clipboard: {e}")
    
    # Also output clipboard data with special markers for Streamlit to capture
    try:
        clipboard_data = df.to_csv(sep='\t', index=False, header=False)
        print("CLIPBOARD_DATA_START")
        print(clipboard_data, end='')
        print("CLIPBOARD_DATA_END")
    except Exception as e:
        print(f"Could not prepare clipboard data for Streamlit: {e}")
    
    return final_file_path


def main():
    """
    Main function to handle CLI arguments and execute transformation.
    """
    if len(sys.argv) < 2:
        print("Usage: python app.py <date>")
        print("Example: python app.py 2026-01-03")
        sys.exit(1)
    
    date_str = sys.argv[1]
    
    # Validate date format
    if not re.match(r'^\d{4}-\d{2}-\d{2}$', date_str):
        print("Error: Date must be in format YYYY-MM-DD (e.g., 2026-01-03)")
        sys.exit(1)
    
    try:
        # Step 1: Transform .txt to export-raw .xlsx
        raw_output_path = transform_txt_to_raw(date_str)
        
        # Step 2: Transform export-raw to export-final .xlsx
        final_output_path = transform_raw_to_final(date_str)
        
        print(f"\nTransformation completed successfully!")
        print(f"Raw export: {raw_output_path}")
        print(f"Final export: {final_output_path}")
        
    except FileNotFoundError as e:
        print(f"Error: {e}")
        sys.exit(1)
    except Exception as e:
        print(f"Error during transformation: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
