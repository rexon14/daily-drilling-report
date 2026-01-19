"""
Daily Report Transformer
Transforms daily report .txt files into structured .xlsx files.

This script reads a daily report text file from the daily-report folder,
parses the structured well data, and exports it to an Excel file in the export-raw folder.

Usage:
    python app.py <date>
    
Example:
    python app.py 2026-01-18
"""

import sys
import os
import re
from datetime import datetime, timedelta
import pandas as pd
from pathlib import Path


def parse_date(date_str):
    """
    Parse date string in various formats to datetime object.
    
    Args:
        date_str: Date string in formats like "28 Dec 2025", "2025-12-28", etc.
    
    Returns:
        datetime object or None if parsing fails
    """
    # Try common date formats
    date_formats = [
        "%d %b %Y",      # 28 Dec 2025
        "%d %B %Y",      # 28 December 2025
        "%Y-%m-%d",      # 2025-12-28
        "%d-%m-%Y",      # 28-12-2025
    ]
    
    for fmt in date_formats:
        try:
            return datetime.strptime(date_str.strip(), fmt)
        except ValueError:
            continue
    
    return None


def extract_well_data(text_content):
    """
    Extract well data from the text report.
    
    Args:
        text_content: Full text content of the report file
    
    Returns:
        List of dictionaries containing well information
    """
    wells = []
    lines = text_content.split('\n')
    
    # Extract region (first line)
    region = lines[0].strip() if lines else "Unknown"
    
    # Extract report date
    report_date = None
    for line in lines[:5]:
        if "Report Date:" in line:
            date_match = re.search(r'Report Date:\s*(.+)', line)
            if date_match:
                report_date = date_match.group(1).strip()
            break
    
    # Split content by asset sections
    current_asset = None
    current_well = {}
    
    i = 0
    while i < len(lines):
        line = lines[i].strip()
        
        # Detect asset section
        if line.startswith("ASSET"):
            current_asset = line.replace("ASSET", "").strip()
            i += 1
            continue
        
        # Detect well entry (starts with number in parentheses)
        well_match = re.match(r'\((\d+)\.\)\s*(.+)', line)
        if well_match:
            # Save previous well if exists
            if current_well:
                wells.append(current_well)
            
            # Start new well entry
            well_info = well_match.group(2).strip()
            # Extract well name and rig: "OHD – QB-117 (Rig ENAFOR 45)"
            rig_match = re.search(r'\(Rig\s+(.+?)\)', well_info)
            rig_name = rig_match.group(1).strip() if rig_match else ""
            well_name = re.sub(r'\s*\(Rig\s+.+?\)', '', well_info).strip()
            # Clean up well name (remove prefix like "OHD – ")
            well_name = re.sub(r'^[^–]+–\s*', '', well_name).strip()
            
            current_well = {
                'Region': region,
                'Asset': current_asset or "",
                'Rig Name': rig_name,
                'Well Name': well_name,
                'Spud Date': None,
                'Current Operation': "",
                '24 hrs look ahead': ""
            }
            i += 1
            continue
        
        # Extract spud date (only "Spud date", not "Start date")
        if line.startswith("* Spud date:"):
            date_match = re.search(r':\s*(.+)', line)
            if date_match:
                date_str = date_match.group(1).strip()
                parsed_date = parse_date(date_str)
                if parsed_date:
                    current_well['Spud Date'] = parsed_date.strftime("%Y-%m-%d")
                else:
                    current_well['Spud Date'] = None
            i += 1
            continue
        
        # Extract current operation
        if "* Current Operation 24 hrs:" in line:
            operation = line.replace("* Current Operation 24 hrs:", "").strip()
            # Check if operation continues on next lines
            i += 1
            while i < len(lines) and not lines[i].strip().startswith("*") and not lines[i].strip().startswith("---"):
                next_line = lines[i].strip()
                if next_line:
                    operation += " " + next_line
                i += 1
            current_well['Current Operation'] = operation.strip()
            continue
        
        # Extract 24 hrs look ahead
        if "* 24 hrs look ahead:" in line:
            look_ahead = line.replace("* 24 hrs look ahead:", "").strip()
            # Check if look ahead continues on next lines
            i += 1
            while i < len(lines) and not lines[i].strip().startswith("*") and not lines[i].strip().startswith("---"):
                next_line = lines[i].strip()
                if next_line:
                    look_ahead += " " + next_line
                i += 1
            current_well['24 hrs look ahead'] = look_ahead.strip()
            continue
        
        i += 1
    
    # Add last well if exists
    if current_well:
        wells.append(current_well)
    
    return wells


def transform_txt_to_xlsx(date_str, base_path="."):
    """
    Transform daily report .txt file to .xlsx file.
    
    Args:
        date_str: Date string in format "YYYY-MM-DD" (e.g., "2026-01-18")
        base_path: Base path containing daily-report and export-raw folders (default: current directory)
    
    Returns:
        Path to the created Excel file
    """
    # Construct file paths
    txt_file_path = os.path.join(base_path, "daily-report", f"{date_str}.txt")
    xlsx_file_path = os.path.join(base_path, "export-raw", f"{date_str}.xlsx")
    
    # Check if input file exists
    if not os.path.exists(txt_file_path):
        raise FileNotFoundError(f"Daily report file not found: {txt_file_path}")
    
    # Create export-raw directory if it doesn't exist
    export_dir = os.path.dirname(xlsx_file_path)
    os.makedirs(export_dir, exist_ok=True)
    
    # Read text file
    print(f"Reading daily report from: {txt_file_path}")
    with open(txt_file_path, 'r', encoding='utf-8') as f:
        text_content = f.read()
    
    # Extract well data
    print("Parsing well data...")
    wells = extract_well_data(text_content)
    
    if not wells:
        raise ValueError("No well data found in the report file")
    
    # Create DataFrame
    df = pd.DataFrame(wells)
    
    # Ensure Spud Date column is properly formatted (handle NaT/None)
    df['Spud Date'] = pd.to_datetime(df['Spud Date'], errors='coerce')
    
    # Export to Excel
    print(f"Exporting to Excel: {xlsx_file_path}")
    df.to_excel(xlsx_file_path, index=False, engine='openpyxl')
    
    print(f"Successfully transformed {len(wells)} wells to Excel format")
    print(f"Output file: {xlsx_file_path}")
    
    return xlsx_file_path


def transform_raw_to_final(date_str, base_path="."):
    """
    Transform export-raw .xlsx file to export-final .xlsx file with additional columns and modifications.
    
    Modifications applied:
    - Add "Flag" column with default value "INC"
    - Apply sentence case to "Region" column
    - Add "Zone" column based on "Asset" column (Algeria: Zone 15, Iraq: Zone 16, Malaysia: Zone 17)
    - Add "APH" column with default value "PIEP"
    - Add "Well Name_2" column that copies "Well Name"
    - Add "Well Type" column with default value "Development"
    - Add "Location" column (Algeria: Onshore, Iraq: Onshore, Malaysia: Offshore)
    - Add "Release Date", "Status", "Status Code [1]", "Status Code [2]", and "Current Status" columns with blank values
    - Add "Report Date" column with value of input date
    - Add "Operation Date" column with value of Report Date minus 1 day
    - Remove "Asset" column
    
    Args:
        date_str: Date string in format "YYYY-MM-DD" (e.g., "2026-01-18")
        base_path: Base path containing export-raw and export-final folders (default: current directory)
    
    Returns:
        Path to the created Excel file
    """
    # Construct file paths
    raw_file_path = os.path.join(base_path, "export-raw", f"{date_str}.xlsx")
    final_file_path = os.path.join(base_path, "export-final", f"{date_str}.xlsx")
    
    # Check if input file exists
    if not os.path.exists(raw_file_path):
        raise FileNotFoundError(f"Export-raw file not found: {raw_file_path}")
    
    # Create export-final directory if it doesn't exist
    export_dir = os.path.dirname(final_file_path)
    os.makedirs(export_dir, exist_ok=True)
    
    # Read raw Excel file
    print(f"Reading export-raw file from: {raw_file_path}")
    df = pd.read_excel(raw_file_path, engine='openpyxl')
    
    # Apply sentence case to Region column
    df['Region'] = df['Region'].str.title()
    
    # Create Zone column based on Asset column
    zone_mapping = {
        'ALGERIA': 'Zone 15',
        'IRAQ': 'Zone 16',
        'MALAYSIA': 'Zone 17'
    }
    df['Zone'] = df['Asset'].str.upper().map(zone_mapping)
    
    # Add APH column with default value "PIEP"
    df['APH'] = 'PIEP'
    
    # Add Well Name_2 column that copies Well Name
    df['Well Name_2'] = df['Well Name']
    
    # Add Well Type column with default value "Development"
    df['Well Type'] = 'Development'
    
    # Create Location column based on Asset column
    location_mapping = {
        'ALGERIA': 'Onshore',
        'IRAQ': 'Onshore',
        'MALAYSIA': 'Offshore'
    }
    df['Location'] = df['Asset'].str.upper().map(location_mapping)
    
    # Parse input date and add Report Date column
    report_date = pd.to_datetime(date_str)
    df['Report Date'] = report_date
    
    # Add Operation Date column (Report Date minus 1 day)
    df['Operation Date'] = report_date - timedelta(days=1)
    
    # Add Flag column with default value "INC"
    df['Flag'] = 'INC'
    
    # Add blank columns: Release Date, Status, Status Code [1], Status Code [2], Current Status
    df['Release Date'] = None
    df['Status'] = None
    df['Status Code [1]'] = None
    df['Status Code [2]'] = None
    df['Current Status'] = None
    
    # Remove Asset column
    df = df.drop(columns=['Asset'])
    
    # Reorder columns to match expected format: Flag, Region, Zone, APH, Rig Name, Well Name, Well Name_2, Well Type, Location, Spud Date, Release Date, Status, Status Code [1], Status Code [2], Current Operation, 24 hrs look ahead, Report Date, Current Status, Operation Date
    column_order = [
        'Flag', 'Region', 'Zone', 'APH', 'Rig Name', 'Well Name', 'Well Name_2', 
        'Well Type', 'Location', 'Spud Date', 'Release Date', 'Status', 
        'Status Code [1]', 'Status Code [2]', 'Current Operation', 'Current Status',
        '24 hrs look ahead', 'Report Date', 'Operation Date'
    ]
    df = df[column_order]
    
    # Sort dataframe ascending by Zone -> Rig Name -> Well Name
    df = df.sort_values(by=['Zone', 'Rig Name', 'Well Name'], ascending=[True, True, True])
    df = df.reset_index(drop=True)
    
    # Export to Excel
    print(f"Exporting final version to Excel: {final_file_path}")
    df.to_excel(final_file_path, index=False, engine='openpyxl')
    
    print(f"Successfully transformed {len(df)} wells to final Excel format")
    print(f"Output file: {final_file_path}")
    
    # Read the exported file and copy to clipboard (values only)
    print("Copying data to clipboard...")
    df_clipboard = pd.read_excel(final_file_path, engine='openpyxl')
    df_clipboard.to_clipboard(index=False, header=False)
    print("Data copied to clipboard successfully!")
    
    return final_file_path


def main():
    """
    Main function to handle command-line arguments and execute transformation.
    """
    # Check if date argument is provided
    if len(sys.argv) < 2:
        print("Usage: python app.py <date>")
        print("Example: python app.py 2026-01-18")
        sys.exit(1)
    
    date_str = sys.argv[1]
    
    # Validate date format (YYYY-MM-DD)
    date_pattern = re.match(r'^\d{4}-\d{2}-\d{2}$', date_str)
    if not date_pattern:
        print("Error: Date must be in format YYYY-MM-DD (e.g., 2026-01-18)")
        sys.exit(1)
    
    try:
        # Step 1: Transform .txt to export-raw .xlsx
        raw_output_path = transform_txt_to_xlsx(date_str)
        
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
