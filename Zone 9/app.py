"""
Daily Report Transformer for Zone 9
Transforms daily report .txt files into structured .xlsx files.

Input: daily-report/{date}.txt (Indonesian format - Laporan Pagi Drilling Region 3 Zona 9)
Output: export-raw/{date}.xlsx and export-final/{date}.xlsx

Usage:
    python app.py <date>

Example:
    python app.py 2026-02-03
"""

import sys
import os
import re
from datetime import datetime, timedelta
import pandas as pd
import numpy as np


def parse_txt_file(path):
    """
    Parse Zone 9 daily drilling report .txt file and extract well data.
    
    Handles:
    - FIELD sections (SANGASANGA, SANGATTA, MUTIARA, SEMBERAH, NILAM)
    - Well entries with optional alternative names: "MUT-385 (MUT-951 OS)"
    - Various day formats: M-8 (MIRU), D-34 (Drilling), WOL-D36 (Wait On Location)
    - Multi-line sections: Summary Report, Current Status, Plan/Next Plan
    
    Args:
        path: Path to the .txt file
    
    Returns:
        DataFrame with well information
    """
    with open(path, encoding='utf-8') as f:
        content = f.read()
    
    rows = []
    current_field = None
    current_item = None
    summary_flag = status_flag = plan_flag = False
    
    for raw in content.splitlines():
        line = raw.strip().replace("*", "")
        if not line:
            continue
        
        low = line.lower()
        
        # FIELD header detection
        if low.startswith('field'):
            parts = line.split(None, 1)
            current_field = parts[1].strip() if len(parts) > 1 else None
            continue
        
        # New record: "1. WELL-NAME (OPTIONAL-PAREN)"
        m = re.match(r'^(\d+)\.\s*(.+)$', line)
        if m:
            # Save previous item
            if current_item:
                rows.append(current_item)
            
            idx = int(m.group(1))
            desc = m.group(2).strip().rstrip('.')
            
            # Parse well name with optional alternative in parentheses
            p = re.match(r'(.+?)\s*\((.+?)\)', desc)
            if p:
                nm1, nm2 = p.group(1).strip(), p.group(2).strip()
            else:
                nm1, nm2 = desc, None
            
            current_item = {
                'No': idx,
                'Field': current_field,
                'Nama Sumur': nm1,
                'Nama Sumur_2': nm2,
                'Nama Rig': None,
                'WOL Hari ke': None,
                'Hari ke': None,
                'AFE Cost': None,
                'Realisasi Biaya': None,
                'Summary Report': None,
                'Current Status': None,
                'Plan': None
            }
            summary_flag = status_flag = plan_flag = False
            continue
        
        if current_item is None:
            continue
        
        # Section headers detection
        if 'summary report' in low or '24 hrs summary' in low:
            summary_flag, status_flag, plan_flag = True, False, False
            continue
        if 'current status' in low:
            status_flag, summary_flag, plan_flag = True, False, False
            continue
        if low.startswith('next plan') or low.startswith('plan') or low.startswith('plan ahead'):
            plan_flag, summary_flag, status_flag = True, False, False
            continue
        
        # Capture first line of Summary/Status/Plan (strip leading special chars)
        if summary_flag:
            val = line.lstrip('-=').strip().rstrip('.')
            current_item['Summary Report'] = val
            summary_flag = False
            continue
        if status_flag:
            val = line.lstrip('-=').strip().rstrip('.')
            current_item['Current Status'] = val
            status_flag = False
            continue
        if plan_flag:
            val = line.lstrip('-=').strip().rstrip('.')
            current_item['Plan'] = val
            plan_flag = False
            continue
        
        # Key: Value lines extraction
        kv = re.match(r'^([^:]+)\s*:\s*(.+)$', line)
        if not kv:
            continue
        
        key, val = kv.group(1).strip(), kv.group(2).strip().rstrip('.')
        kl = key.lower()
        
        # Nama Rig / Rig Name (strip "Rig" prefix)
        if kl in ('nama rig', 'rig name'):
            clean = re.sub(r'^[Rr]ig\s*', '', val).strip()
            current_item['Nama Rig'] = clean
            continue
        
        # WOL Hari ke (Wait On Location days)
        if 'wol' in kl and 'hari' in kl:
            num = re.sub(r'[^\d.-]', '', val)
            current_item['WOL Hari ke'] = float(num) if num else None
            continue
        
        # Hari ke / Days / Drilling Days
        if kl in ('hari ke', 'days') or 'drilling days' in kl:
            current_item['Hari ke'] = val
            continue
        
        # AFE Cost
        if 'afe' in kl:
            num = re.sub(r'[^\d.]', '', val)
            num = num.rstrip('.')  # Remove trailing periods
            current_item['AFE Cost'] = float(num) if num else None
            continue
        
        # Realisasi Biaya (Actual Cost) - take value before parentheses
        if 'realisasi' in kl:
            before_paren = val.split('(')[0].strip()
            num = re.sub(r'[^\d.]', '', before_paren)
            num = num.rstrip('.')  # Remove trailing periods
            current_item['Realisasi Biaya'] = float(num) if num else None
            continue
    
    # Append last item
    if current_item:
        rows.append(current_item)
    
    # Build DataFrame with specified column order
    cols = [
        'No', 'Field', 'Nama Sumur', 'Nama Sumur_2', 'Nama Rig',
        'WOL Hari ke', 'Hari ke', 'AFE Cost', 'Realisasi Biaya',
        'Summary Report', 'Current Status', 'Plan'
    ]
    return pd.DataFrame(rows, columns=cols)


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
    
    print(f"Successfully transformed {len(df)} wells to raw Excel format")
    return xlsx_file_path


def transform_raw_to_final(date_str, base_path="."):
    """
    Transform export-raw .xlsx to export-final .xlsx with standardized columns.
    
    Adds columns: Flag, Region, Zone, APH, Well Name [2], Well Type, Location,
    Spud Date, Release Date, Status, Status Code [1], Status Code [2],
    Report Date, Operation Date
    
    APH Logic:
    - Default: "PHSS"
    - If "PDSI" substring in Rig Name: "PEP"
    
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
    
    # Select required columns from raw
    cols = ['Nama Sumur', 'Nama Sumur_2', 'Nama Rig', 'Summary Report', 'Current Status', 'Plan']
    df = df[cols].copy()
    
    # Add standard columns
    df = df.assign(**{
        "Flag": "INC",
        "Region": "Region 3",
        "Zone": "Zone 9",
        "APH": "PHSS",  # Default value, will be updated based on Rig Name
        "Well Type": "Development",
        "Location": "Onshore",
        "Spud Date": np.nan,
        "Release Date": np.nan,
        "Status": np.nan,
        "Status Code [1]": np.nan,
        "Status Code [2]": np.nan
    })
    
    # Rename columns
    df.rename(columns={
        "Nama Sumur": "Well Name",
        "Nama Sumur_2": "Well Name [2]",
        "Nama Rig": "Rig Name",
        "Plan": "Next Plan"
    }, inplace=True)
    
    # Fill Well Name [2] with Well Name if empty
    df['Well Name [2]'] = df['Well Name [2]'].fillna(df['Well Name'])
    
    # Update APH based on Rig Name (if contains "PDSI", set APH to "PEP")
    df['Rig Name'] = df['Rig Name'].astype(str)
    df.loc[df['Rig Name'].str.contains('PDSI', case=False, na=False), 'APH'] = 'PEP'
    
    # Clean text fields
    clean_text_fields(df)
    
    # Add date columns
    report_ts = pd.to_datetime(date_str)
    df['Report Date'] = report_ts
    df['Operation Date'] = report_ts - pd.Timedelta(days=1)
    df['Report Date'] = df['Report Date'].dt.date
    df['Operation Date'] = df['Operation Date'].dt.date
    
    # Reorder columns to standard format
    cols_order = [
        "Flag", "Region", "Zone", "APH", "Rig Name", "Well Name", "Well Name [2]",
        "Well Type", "Location", "Spud Date", "Release Date", "Status",
        "Status Code [1]", "Status Code [2]", "Summary Report", "Current Status",
        "Next Plan", "Report Date", "Operation Date"
    ]
    df = df[cols_order]

    # Fix spacing in Rig Name if missing (e.g., "PDSI#21.2/OW 700-M" → "PDSI #21.2/OW 700-M")
    df["Rig Name"] = df["Rig Name"].replace("PDSI#21.2/OW 700-M", "PDSI #21.2/OW 700-M")
    
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
    
    # Output clipboard data with special markers for Streamlit to capture
    try:
        clipboard_data = df.to_csv(sep='\t', index=False, header=False)
        print("CLIPBOARD_DATA_START")
        print(clipboard_data, end='')
        print("CLIPBOARD_DATA_END")
    except Exception as e:
        print(f"Could not prepare clipboard data for Streamlit: {e}")
    
    return final_file_path


def main():
    """Main function to handle CLI arguments and execute transformation."""
    if len(sys.argv) < 2:
        print("Usage: python app.py <date>")
        print("Example: python app.py 2026-02-03")
        sys.exit(1)
    
    date_str = sys.argv[1]
    
    # Validate date format
    if not re.match(r'^\d{4}-\d{2}-\d{2}$', date_str):
        print("Error: Date must be in format YYYY-MM-DD (e.g., 2026-02-03)")
        sys.exit(1)
    
    # Get the script's directory as base path (script is now at Zone 9 root)
    base_path = os.path.dirname(os.path.abspath(__file__))
    
    try:
        # Step 1: Transform .txt to export-raw .xlsx
        raw_output_path = transform_txt_to_raw(date_str, base_path)
        
        # Step 2: Transform export-raw to export-final .xlsx
        final_output_path = transform_raw_to_final(date_str, base_path)
        
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
