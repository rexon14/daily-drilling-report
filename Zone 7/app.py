"""
Daily Report Transformer for Zone 7
Transforms daily report .txt files into structured .xlsx files.

Input: daily-report/{date}.txt (Indonesian format - Laporan Pagi Pemboran PEP)
Output: export-raw/{date}.xlsx and export-final/{date}.xlsx

Usage:
    python app.py <date>

Example:
    python app.py 2026-01-19
"""

import sys
import os
import re
from datetime import datetime, timedelta
import pandas as pd


def clean_text_for_excel(text):
    """Remove leading '-' from first line only to prevent Excel formula interpretation."""
    if not text:
        return text
    text = str(text).strip()
    lines = text.split('\n')
    # Only clean the first line's leading '-'
    while lines[0].strip().startswith('-'):
        lines[0] = lines[0].strip()[1:].strip()
    return '\n'.join(lines)


def parse_well_section(lines, start_idx, field_name):
    """
    Parse a single well section from the report.
    
    Args:
        lines: List of all lines in the report
        start_idx: Starting line index for this well
        field_name: Current field name (JATIBARANG/SUBANG)
    
    Returns:
        Tuple of (well_data dict, next_index)
    """
    well = {
        'Field': field_name,
        'Well Name': '',
        'Location': '',
        'Rig Name': '',
        'Company Man': '',
        'Days': None,
        'Depth': '',
        'DSR': '',
        'AFE': '',
        'Realization': '',
        'Summary': '',
        'Current Status': '',
        'Next Plan': ''
    }
    
    i = start_idx
    current_section = None
    section_content = []
    
    while i < len(lines):
        line = lines[i].strip()
        
        # Check for next well entry or field section (end of current well)
        if i > start_idx:
            # Next numbered well entry
            if re.match(r'^\d+\.\s+\w+', line):
                break
            # Next field section
            if line.startswith('FIELD '):
                break
            # End of report markers
            if line.startswith('Terima kasih') or line.startswith('Salam'):
                break
        
        # Parse well number and name from first line (e.g., "1. AMJ-004")
        if i == start_idx:
            match = re.match(r'^\d+\.\s+(.+)$', line)
            if match:
                well['Well Name'] = match.group(1).strip()
        
        # Parse specific fields
        if line.startswith('Nama Lokasi'):
            match = re.search(r':\s*(.+)', line)
            if match:
                well['Location'] = match.group(1).strip()
        
        elif line.startswith('Nama Rig'):
            match = re.search(r':\s*(.+)', line)
            if match:
                well['Rig Name'] = match.group(1).strip()
        
        elif line.startswith('Coman on') or line.startswith('Coman On'):
            match = re.search(r':\s*(.+)', line)
            if match:
                well['Company Man'] = match.group(1).strip()
        
        elif line.startswith('Hari ke'):
            match = re.search(r':\s*(\d+)', line)
            if match:
                well['Days'] = int(match.group(1))
        
        elif line.startswith('Kedalaman'):
            match = re.search(r':\s*(.+)', line)
            if match:
                well['Depth'] = match.group(1).strip()
        
        elif line.startswith('DSR'):
            match = re.search(r':\s*(.+)', line)
            if match:
                well['DSR'] = match.group(1).strip()
        
        elif line.startswith('AFE'):
            match = re.search(r':\s*(.+)', line)
            if match:
                well['AFE'] = match.group(1).strip()
        
        elif line.startswith('Realisasi'):
            match = re.search(r':\s*(.+)', line)
            if match:
                well['Realization'] = match.group(1).strip()
        
        # Handle multi-line sections: Summary, Current Status, Next Plan
        elif re.match(r'^Summary\s*[Rr]eport\s*:', line, re.IGNORECASE):
            # Save previous section if exists
            if current_section and section_content:
                well[current_section] = '\n'.join(section_content).strip()
            current_section = 'Summary'
            section_content = []
            match = re.search(r':\s*(.+)', line)
            if match and match.group(1).strip():
                section_content.append(match.group(1).strip())
        
        elif re.match(r'^Current\s*[Ss]tatus\s*:', line, re.IGNORECASE):
            if current_section and section_content:
                well[current_section] = '\n'.join(section_content).strip()
            current_section = 'Current Status'
            section_content = []
            match = re.search(r':\s*(.+)', line)
            if match and match.group(1).strip():
                section_content.append(match.group(1).strip())
        
        elif re.match(r'^Next\s*[Pp]lan\s*:', line, re.IGNORECASE):
            if current_section and section_content:
                well[current_section] = '\n'.join(section_content).strip()
            current_section = 'Next Plan'
            section_content = []
            match = re.search(r':\s*(.+)', line)
            if match and match.group(1).strip():
                section_content.append(match.group(1).strip())
        
        # Continue collecting multi-line section content
        elif current_section and line and not line.startswith(('Nama', 'Coman', 'Hari', 'Kedalaman', 'DSR', 'AFE', 'Realisasi', 'Penambahan', 'Plan', 'Casing')):
            # Skip lines that start new sections or are field headers
            if not re.match(r'^\d+\.\s+', line) and not line.startswith('FIELD'):
                section_content.append(line)
        
        i += 1
    
    # Save last section
    if current_section and section_content:
        well[current_section] = '\n'.join(section_content).strip()
    
    # Clean up content (remove trailing section markers that may have been captured)
    for field in ['Summary', 'Current Status', 'Next Plan']:
        if well[field]:
            # Remove any trailing section markers (handle multiline)
            well[field] = re.sub(r'\n?\s*(Current Status|Next plan|Next Plan|Summary report|Summary Report)\s*:.*$', '', well[field], flags=re.IGNORECASE | re.MULTILINE)
            well[field] = well[field].strip()
    
    return well, i


def extract_wells_from_txt(text_content):
    """
    Extract well data from Zone 7 daily report text format.
    
    Args:
        text_content: Full text content of the report file
    
    Returns:
        List of dictionaries containing well information
    """
    wells = []
    lines = text_content.split('\n')
    
    current_field = ''
    i = 0
    
    while i < len(lines):
        line = lines[i].strip()
        
        # Detect field section
        if line.startswith('FIELD '):
            current_field = line.replace('FIELD ', '').strip()
            i += 1
            continue
        
        # Detect well entry (starts with number and dot)
        well_match = re.match(r'^(\d+)\.\s+(.+)$', line)
        if well_match and current_field:
            well_data, next_i = parse_well_section(lines, i, current_field)
            if well_data['Well Name']:
                wells.append(well_data)
            i = next_i
            continue
        
        i += 1
    
    return wells


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
    with open(txt_file_path, 'r', encoding='utf-8') as f:
        text_content = f.read()
    
    print("Parsing well data...")
    wells = extract_wells_from_txt(text_content)
    
    if not wells:
        raise ValueError("No well data found in the report file")
    
    # Create DataFrame with raw columns
    df = pd.DataFrame(wells)
    
    # Select and rename columns for raw export
    raw_columns = {
        'Field': 'Field',
        'Well Name': 'Well Name',
        'Location': 'Location Name',
        'Rig Name': 'Rig Name',
        'Company Man': 'Company Man',
        'Days': 'Days',
        'Depth': 'Depth',
        'Summary': 'Summary',
        'Current Status': 'Current Status',
        'Next Plan': 'Next Plan'
    }
    
    result_df = df[[col for col in raw_columns.keys() if col in df.columns]].copy()
    result_df.columns = [raw_columns[col] for col in result_df.columns]
    
    print(f"Exporting to Excel: {xlsx_file_path}")
    result_df.to_excel(xlsx_file_path, index=False, engine='openpyxl')
    
    print(f"Successfully transformed {len(wells)} wells to Excel format")
    return xlsx_file_path


def transform_raw_to_final(date_str, base_path="."):
    """
    Transform export-raw .xlsx to export-final .xlsx with standardized columns.
    
    Adds columns: Flag, Region, Zone, APH, Well Name_2, Well Type, Location,
    Spud Date, Release Date, Status, Status Code [1], Status Code [2],
    Current Status, Report Date, Operation Date
    
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
    
    # Add standard columns
    df['Flag'] = 'INC'
    df['Region'] = 'Region 2'
    df['Zone'] = 'Zone 7'
    df['APH'] = 'PEP'
    df['Well Name_2'] = df['Well Name'] if 'Well Name' in df.columns else ''
    df['Well Type'] = 'Development'
    df['Location'] = 'Onshore'
    
    # Parse dates
    report_date = pd.to_datetime(date_str)
    df['Report Date'] = report_date
    df['Operation Date'] = report_date - timedelta(days=1)
    
    # Add blank columns
    df['Spud Date'] = None
    df['Release Date'] = None
    df['Status'] = None
    df['Status Code [1]'] = None
    df['Status Code [2]'] = None
    
    # Fill blank Summary with Current Status, or Next Plan if both are blank
    if 'Summary' in df.columns:
        for idx in df.index:
            if pd.isna(df.loc[idx, 'Summary']) or str(df.loc[idx, 'Summary']).strip() == '':
                if 'Current Status' in df.columns and pd.notna(df.loc[idx, 'Current Status']) and str(df.loc[idx, 'Current Status']).strip() != '':
                    df.loc[idx, 'Summary'] = df.loc[idx, 'Current Status']
                elif 'Next Plan' in df.columns and pd.notna(df.loc[idx, 'Next Plan']) and str(df.loc[idx, 'Next Plan']).strip() != '':
                    df.loc[idx, 'Summary'] = df.loc[idx, 'Next Plan']
    
    # Clean text fields to prevent Excel formula errors (remove leading '-')
    text_columns = ['Summary', 'Current Status', 'Next Plan']
    for col in text_columns:
        if col in df.columns:
            df[col] = df[col].apply(lambda x: clean_text_for_excel(x) if pd.notna(x) else x)
    
    # Reorder columns to match standard format
    column_order = [
        'Flag', 'Region', 'Zone', 'APH', 'Rig Name', 'Well Name', 'Well Name_2',
        'Well Type', 'Location', 'Spud Date', 'Release Date', 'Status',
        'Status Code [1]', 'Status Code [2]', 'Summary', 'Current Status',
        'Next Plan', 'Report Date', 'Operation Date'
    ]
    
    available_columns = [col for col in column_order if col in df.columns]
    df = df[available_columns]
    
    # Sort by Zone -> Rig Name -> Well Name
    sort_cols = [col for col in ['Zone', 'Rig Name', 'Well Name'] if col in df.columns]
    if sort_cols:
        df = df.sort_values(by=sort_cols, ascending=True)
    df = df.reset_index(drop=True)
    
    print(f"Exporting final version to Excel: {final_file_path}")
    df.to_excel(final_file_path, index=False, engine='openpyxl')
    
    print(f"Successfully transformed {len(df)} wells to final Excel format")
    
    # Prepare clipboard data (values only, no header) in TSV format
    # Output with special markers so text_form.py can capture and copy it
    print("Preparing clipboard data...")
    try:
        # Convert dataframe to TSV format (tab-separated values)
        clipboard_data = df.to_csv(sep='\t', index=False, header=False)
        # Output with special markers for text_form.py to capture
        print("CLIPBOARD_DATA_START")
        print(clipboard_data, end='')
        print("CLIPBOARD_DATA_END")
        print("Clipboard data prepared successfully!")
    except Exception as e:
        print(f"Could not prepare clipboard data: {e}")
    
    return final_file_path


def main():
    """Main function to handle CLI arguments and execute transformation."""
    if len(sys.argv) < 2:
        print("Usage: python app.py <date>")
        print("Example: python app.py 2026-01-19")
        sys.exit(1)
    
    date_str = sys.argv[1]
    
    # Validate date format
    if not re.match(r'^\d{4}-\d{2}-\d{2}$', date_str):
        print("Error: Date must be in format YYYY-MM-DD (e.g., 2026-01-19)")
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
