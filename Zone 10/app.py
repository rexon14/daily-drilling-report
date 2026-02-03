"""
Daily Report Transformer for Zone 10
Transforms daily report .txt files into structured .xlsx files.

Input: daily-report/{date}.txt (Indonesian format - Laporan Pagi Drilling Zona-10)
Output: export-raw/{date}.xlsx and export-final/{date}.xlsx

Usage:
    python app.py <date>

Example:
    python app.py 2025-04-26
"""

import sys
import os
import re
from datetime import datetime, timedelta
import pandas as pd
import numpy as np


def parse_txt_file(path):
    """
    Parse Zone 10 daily drilling report .txt file and extract well data.
    
    Handles:
    - Well entries starting with number: "1. SBKD-001"
    - Fields: Nama Sumur, Nama Rig, Hari ke, Est. Tgl Selesai, Kedalaman, 
      Progres, AFE, Realisasi biaya, Summary report, Morning Status, Next plan
    - Multi-line sections: Summary report, Morning Status, Next plan
    
    Args:
        path: Path to the .txt file
    
    Returns:
        DataFrame with well information
    """
    with open(path, encoding='utf-8') as f:
        content = f.read()
    
    rows = []
    current_item = None
    summary_flag = status_flag = plan_flag = False
    
    lines = content.splitlines()
    i = 0
    
    while i < len(lines):
        raw = lines[i]
        line = raw.strip().replace("*", "")
        if not line:
            i += 1
            continue
        
        # New record: "1. WELL-NAME"
        m = re.match(r'^(\d+)\.\s*(.+)$', line)
        if m:
            # Save previous item
            if current_item:
                rows.append(current_item)
            
            idx = int(m.group(1))
            well_name = m.group(2).strip()
            
            current_item = {
                'No': idx,
                'Nama Sumur': well_name,
                'Nama Rig': None,
                'AFE': None,
                'Realisasi Biaya': None,
                'Summary Report': None,
                'Morning Status': None,
                'Next Plan': None
            }
            summary_flag = status_flag = plan_flag = False
            i += 1
            continue
        
        if current_item is None:
            i += 1
            continue
        
        # Helper function to check if a line is a known field header
        def is_field_header(text):
            """Check if line is a known field header (not content with colons)."""
            # List of known field headers that should stop section content capture
            known_fields = [
                'nama sumur', 'well name', 'nama rig', 'rig name',
                'hari ke', 'est. tgl selesai', 'kedalaman', 'progres',
                'afe', 'realisasi biaya', 'casing setting depth',
                'plan - actual', 'penambahan', 'dsr'
            ]
            text_lower = text.lower().strip()
            # Check if it matches a known field pattern exactly
            for field in known_fields:
                # Match field name followed by colon (with optional spaces)
                if re.match(rf'^{re.escape(field)}\s*:\s*', text_lower):
                    return True
            # Also check for new well entry
            if re.match(r'^\d+\.\s+', text):
                return True
            return False
        
        # Section headers detection (case-insensitive, handle variations)
        low = line.lower()
        
        # Summary Report section
        if 'summary report' in low or '24 hrs summary' in low:
            summary_flag, status_flag, plan_flag = True, False, False
            i += 1
            # Skip empty lines after header
            while i < len(lines) and not lines[i].strip():
                i += 1
            # Capture content (may span multiple lines)
            content_parts = []
            while i < len(lines):
                next_line = lines[i].strip().replace("*", "")
                if not next_line:
                    i += 1
                    continue
                # Stop if we hit a new well, section header, or known field header
                if (re.match(r'^\d+\.\s+', next_line) or 
                    'morning status' in next_line.lower() or 
                    'next plan' in next_line.lower() or
                    'current status' in next_line.lower() or
                    is_field_header(next_line)):
                    break
                # Clean line and add to content (preserve the dash prefix for now)
                cleaned_line = next_line.lstrip('=').strip()
                if cleaned_line:
                    content_parts.append(cleaned_line)
                i += 1
            if content_parts:
                current_item['Summary Report'] = ' '.join(content_parts).strip()
            summary_flag = False
            continue
        
        # Morning Status section
        if 'morning status' in low or 'current status' in low:
            status_flag, summary_flag, plan_flag = True, False, False
            i += 1
            # Skip empty lines after header
            while i < len(lines) and not lines[i].strip():
                i += 1
            # Capture content (may span multiple lines)
            content_parts = []
            while i < len(lines):
                next_line = lines[i].strip().replace("*", "")
                if not next_line:
                    i += 1
                    continue
                # Stop if we hit a new well, section header, or known field header
                if (re.match(r'^\d+\.\s+', next_line) or 
                    'summary report' in next_line.lower() or 
                    'next plan' in next_line.lower() or
                    '24 hrs summary' in next_line.lower() or
                    is_field_header(next_line)):
                    break
                # Clean line and add to content (preserve the dash prefix for now)
                cleaned_line = next_line.lstrip('=').strip()
                if cleaned_line:
                    content_parts.append(cleaned_line)
                i += 1
            if content_parts:
                current_item['Morning Status'] = ' '.join(content_parts).strip()
            status_flag = False
            continue
        
        # Next Plan section
        if 'next plan' in low or low.startswith('plan') or 'plan ahead' in low:
            plan_flag, summary_flag, status_flag = True, False, False
            i += 1
            # Skip empty lines after header
            while i < len(lines) and not lines[i].strip():
                i += 1
            # Capture content (may span multiple lines)
            content_parts = []
            while i < len(lines):
                next_line = lines[i].strip().replace("*", "")
                if not next_line:
                    i += 1
                    continue
                # Stop if we hit a new well, section header, or known field header
                if (re.match(r'^\d+\.\s+', next_line) or 
                    'summary report' in next_line.lower() or 
                    'morning status' in next_line.lower() or
                    'current status' in next_line.lower() or
                    '24 hrs summary' in next_line.lower() or
                    is_field_header(next_line)):
                    break
                # Clean line and add to content (preserve the dash prefix for now)
                cleaned_line = next_line.lstrip('=').strip()
                if cleaned_line:
                    content_parts.append(cleaned_line)
                i += 1
            if content_parts:
                current_item['Next Plan'] = ' '.join(content_parts).strip()
            plan_flag = False
            continue
        
        # Key: Value lines extraction
        kv = re.match(r'^([^:]+)\s*:\s*(.+)$', line)
        if kv:
            key, val = kv.group(1).strip(), kv.group(2).strip()
            kl = key.lower()
            
            # Nama Sumur
            if kl in ('nama sumur', 'well name'):
                current_item['Nama Sumur'] = val
            
            # Nama Rig / Rig Name
            elif kl in ('nama rig', 'rig name'):
                clean = re.sub(r'^[Rr]ig\s*', '', val).strip()
                current_item['Nama Rig'] = clean
            
            # AFE
            elif 'afe' in kl:
                num = re.sub(r'[^\d.]', '', val)
                num = num.rstrip('.')
                try:
                    current_item['AFE'] = float(num) if num else None
                except ValueError:
                    current_item['AFE'] = None
            
            # Realisasi Biaya (Actual Cost) - take value before parentheses
            elif 'realisasi' in kl:
                before_paren = val.split('(')[0].strip()
                num = re.sub(r'[^\d.]', '', before_paren)
                num = num.rstrip('.')
                try:
                    current_item['Realisasi Biaya'] = float(num) if num else None
                except ValueError:
                    current_item['Realisasi Biaya'] = None
        
        i += 1
    
    # Append last item
    if current_item:
        rows.append(current_item)
    
    # Build DataFrame with specified column order
    cols = [
        'No', 'Nama Sumur', 'Nama Rig', 'AFE', 'Realisasi Biaya',
        'Summary Report', 'Morning Status', 'Next Plan'
    ]
    return pd.DataFrame(rows, columns=cols)


def clean_text_for_excel(text):
    """
    Remove leading '-' from first line only to prevent Excel formula interpretation.
    Based on Zone 7 implementation pattern.
    
    Args:
        text: Text string to clean
    
    Returns:
        Cleaned text string
    """
    if not text:
        return text
    text = str(text).strip()
    lines = text.split('\n')
    # Only clean the first line's leading '-'
    while lines and lines[0].strip().startswith('-'):
        lines[0] = lines[0].strip()[1:].strip()
    return '\n'.join(lines)


def clean_text_fields(df):
    """
    Clean text fields by removing leading special characters.
    Enhanced with Excel-safe cleaning to prevent formula interpretation errors.
    
    Args:
        df: DataFrame to clean
    """
    text_fields = ['Summary Report', 'Morning Status', 'Next Plan', 'Current Status']
    for field in text_fields:
        if field in df.columns:
            # Apply Excel-safe cleaning (remove leading '-' from first line)
            df[field] = df[field].apply(lambda x: clean_text_for_excel(x) if pd.notna(x) else x)
            # Also apply regex cleaning for other special characters
            df[field] = df[field].astype(str).str.replace(r'^[\-\:=]+\s*', '', regex=True)
            df[field] = df[field].replace('nan', np.nan)


def transform_txt_to_raw(date_str, base_path="."):
    """
    Transform daily report .txt file to export-raw .xlsx file.
    
    Enhanced with better validation and error handling.
    
    Args:
        date_str: Date string in format "YYYY-MM-DD"
        base_path: Base path containing daily-report and export-raw folders
    
    Returns:
        Path to the created Excel file
    
    Raises:
        FileNotFoundError: If input file doesn't exist
        ValueError: If no well data found or parsing fails
    """
    txt_file_path = os.path.join(base_path, "daily-report", f"{date_str}.txt")
    xlsx_file_path = os.path.join(base_path, "export-raw", f"{date_str}.xlsx")
    
    if not os.path.exists(txt_file_path):
        raise FileNotFoundError(f"Daily report file not found: {txt_file_path}")
    
    # Create output directory if it doesn't exist
    os.makedirs(os.path.dirname(xlsx_file_path), exist_ok=True)
    
    print(f"Reading daily report from: {txt_file_path}")
    try:
        df = parse_txt_file(txt_file_path)
    except Exception as e:
        raise ValueError(f"Error parsing text file: {e}")
    
    if df.empty:
        raise ValueError("No well data found in the report file")
    
    # Validate required columns exist
    required_cols = ['No', 'Nama Sumur', 'Nama Rig', 'AFE', 'Realisasi Biaya', 
                     'Summary Report', 'Morning Status', 'Next Plan']
    missing_cols = [col for col in required_cols if col not in df.columns]
    if missing_cols:
        print(f"Warning: Missing columns in raw data: {missing_cols}")
    
    print(f"Parsed {len(df)} well(s) from report")
    print(f"Exporting to Excel: {xlsx_file_path}")
    
    try:
        df.to_excel(xlsx_file_path, index=False, engine='openpyxl')
        print(f"Successfully transformed {len(df)} wells to raw Excel format")
    except Exception as e:
        raise ValueError(f"Error writing Excel file: {e}")
    
    return xlsx_file_path


def transform_raw_to_final(date_str, base_path="."):
    """
    Transform export-raw .xlsx to export-final .xlsx with standardized columns.
    
    Adds columns: Flag, Region, Zone, APH, Well Name [2], Well Type, Location,
    Spud Date, Release Date, Status, Status Code [1], Status Code [2],
    Report Date, Operation Date
    
    APH Logic:
    - Default: "PHKT"
    - If "PDSI" substring in Rig Name: "PEP"
    
    Args:
        date_str: Date string in format "YYYY-MM-DD"
        base_path: Base path containing export-raw and export-final folders
    
    Returns:
        Path to the created Excel file
    
    Raises:
        FileNotFoundError: If raw file doesn't exist
        ValueError: If required columns are missing or transformation fails
    """
    raw_file_path = os.path.join(base_path, "export-raw", f"{date_str}.xlsx")
    final_file_path = os.path.join(base_path, "export-final", f"{date_str}.xlsx")
    
    if not os.path.exists(raw_file_path):
        raise FileNotFoundError(f"Export-raw file not found: {raw_file_path}")
    
    # Create output directory if it doesn't exist
    os.makedirs(os.path.dirname(final_file_path), exist_ok=True)
    
    print(f"Reading export-raw file from: {raw_file_path}")
    try:
        df = pd.read_excel(raw_file_path, engine='openpyxl')
    except Exception as e:
        raise ValueError(f"Error reading raw Excel file: {e}")
    
    if df.empty:
        raise ValueError("Raw Excel file is empty")
    
    # Validate required columns exist in raw data
    required_raw_cols = ['Nama Sumur', 'Nama Rig', 'Summary Report', 'Morning Status', 'Next Plan']
    missing_cols = [col for col in required_raw_cols if col not in df.columns]
    if missing_cols:
        raise ValueError(f"Missing required columns in raw data: {missing_cols}")
    
    # Select required columns from raw
    df = df[required_raw_cols].copy()
    
    # Add standard columns with Flag column (missing in original)
    df = df.assign(**{
        "Flag": "INC",
        "Region": "Region 3",
        "Zone": "Zone 10",
        "APH": "PHKT",  # Default value, will be updated based on Rig Name
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
        "Nama Rig": "Rig Name",
        "Morning Status": "Current Status"
    }, inplace=True)
    
    # Add Well Name [2] column that copies Well Name
    df['Well Name [2]'] = df['Well Name']
    
    # Update APH based on Rig Name (if contains "PDSI", set APH to "PEP")
    df['Rig Name'] = df['Rig Name'].astype(str)
    df.loc[df['Rig Name'].str.contains('PDSI', case=False, na=False), 'APH'] = 'PEP'
    
    # Clean text fields (Excel-safe cleaning)
    clean_text_fields(df)
    
    # Add date columns
    try:
        report_ts = pd.to_datetime(date_str)
        df['Report Date'] = report_ts
        df['Operation Date'] = report_ts - pd.Timedelta(days=1)
        df['Report Date'] = df['Report Date'].dt.date
        df['Operation Date'] = df['Operation Date'].dt.date
    except Exception as e:
        raise ValueError(f"Error processing dates: {e}")
    
    # Reorder columns to standard format (with Flag column first)
    cols_order = [
        "Flag", "Region", "Zone", "APH", "Rig Name", "Well Name", "Well Name [2]",
        "Well Type", "Location", "Spud Date", "Release Date", "Status",
        "Status Code [1]", "Status Code [2]", "Summary Report", "Current Status",
        "Next Plan", "Report Date", "Operation Date"
    ]
    
    # Ensure all columns exist before reordering
    available_cols = [col for col in cols_order if col in df.columns]
    missing_final_cols = [col for col in cols_order if col not in df.columns]
    if missing_final_cols:
        print(f"Warning: Missing columns in final output: {missing_final_cols}")
    
    df = df[available_cols]
    
    # Sort by Rig Name (case-insensitive)
    df.sort_values(
        by='Rig Name',
        key=lambda s: s.str.lower(),
        ascending=True,
        inplace=True
    )
    df.reset_index(drop=True, inplace=True)
    
    print(f"Exporting final version to Excel: {final_file_path}")
    try:
        df.to_excel(final_file_path, index=False, engine='openpyxl')
        print(f"Successfully transformed {len(df)} wells to final Excel format")
    except Exception as e:
        raise ValueError(f"Error writing final Excel file: {e}")
    
    # Copy to clipboard (values only, no header)
    print("Copying data to clipboard...")
    try:
        df.to_clipboard(index=False, header=False)
        print("Data copied to clipboard successfully!")
    except Exception as e:
        print(f"Could not copy to clipboard: {e}")
    
    # Output clipboard data with special markers for Streamlit to capture
    try:
        # Use utf-8-sig encoding to handle special characters better
        import io
        output = io.StringIO()
        df.to_csv(output, sep='\t', index=False, header=False, encoding='utf-8')
        clipboard_data = output.getvalue()
        print("CLIPBOARD_DATA_START")
        print(clipboard_data, end='')
        print("CLIPBOARD_DATA_END")
    except Exception as e:
        # Fallback: try without special characters handling
        try:
            clipboard_data = df.to_csv(sep='\t', index=False, header=False)
            print("CLIPBOARD_DATA_START")
            print(clipboard_data, end='')
            print("CLIPBOARD_DATA_END")
        except Exception as e2:
            print(f"Could not prepare clipboard data for Streamlit: {e2}")
    
    return final_file_path


def main():
    """
    Main function to handle CLI arguments and execute transformation.
    
    Enhanced with better error handling, validation, and informative logging.
    """
    if len(sys.argv) < 2:
        print("Usage: python app.py <date>")
        print("Example: python app.py 2025-04-26")
        sys.exit(1)
    
    date_str = sys.argv[1]
    
    # Validate date format
    if not re.match(r'^\d{4}-\d{2}-\d{2}$', date_str):
        print("Error: Date must be in format YYYY-MM-DD (e.g., 2025-04-26)")
        sys.exit(1)
    
    # Validate date is a valid date
    try:
        pd.to_datetime(date_str)
    except ValueError:
        print(f"Error: Invalid date: {date_str}")
        sys.exit(1)
    
    # Get the script's directory as base path
    base_path = os.path.dirname(os.path.abspath(__file__))
    
    # Validate base path structure
    daily_report_dir = os.path.join(base_path, "daily-report")
    if not os.path.exists(daily_report_dir):
        print(f"Warning: daily-report directory not found at {daily_report_dir}")
        print("Creating directory...")
        os.makedirs(daily_report_dir, exist_ok=True)
    
    try:
        print(f"\n{'='*60}")
        print(f"Zone 10 Daily Report Transformer")
        print(f"Processing date: {date_str}")
        print(f"{'='*60}\n")
        
        # Step 1: Transform .txt to export-raw .xlsx
        print("Step 1: Converting .txt to raw Excel format...")
        raw_output_path = transform_txt_to_raw(date_str, base_path)
        print(f"[OK] Raw export completed: {raw_output_path}\n")
        
        # Step 2: Transform export-raw to export-final .xlsx
        print("Step 2: Converting raw Excel to final Excel format...")
        final_output_path = transform_raw_to_final(date_str, base_path)
        print(f"[OK] Final export completed: {final_output_path}\n")
        
        print(f"{'='*60}")
        print(f"Transformation completed successfully!")
        print(f"{'='*60}")
        print(f"Raw export:   {raw_output_path}")
        print(f"Final export: {final_output_path}")
        print(f"{'='*60}\n")
        
    except FileNotFoundError as e:
        print(f"\n{'='*60}")
        print(f"ERROR: File not found")
        print(f"{'='*60}")
        print(f"{e}")
        print(f"{'='*60}\n")
        sys.exit(1)
    except ValueError as e:
        print(f"\n{'='*60}")
        print(f"ERROR: Validation error")
        print(f"{'='*60}")
        print(f"{e}")
        print(f"{'='*60}\n")
        sys.exit(1)
    except Exception as e:
        print(f"\n{'='*60}")
        print(f"ERROR: Unexpected error during transformation")
        print(f"{'='*60}")
        print(f"{e}")
        print(f"\nTraceback:")
        import traceback
        traceback.print_exc()
        print(f"{'='*60}\n")
        sys.exit(1)


if __name__ == "__main__":
    main()
