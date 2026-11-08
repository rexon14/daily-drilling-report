"""
Daily Report Transformer for Zone 11
Transforms daily report .txt files into structured .xlsx files.

Input: daily-report/{date}.txt (Indonesian format - Laporan Pagi Drilling Zona 11)
Output: export-raw/{date}.xlsx and export-final/{date}.xlsx

Usage:
    python app.py <date>

Example:
    python app.py 2026-11-08
"""

import sys
import re
from pathlib import Path
import pandas as pd
import numpy as np

SCRIPT_DIR = Path(__file__).parent

# Rig name fixes (applied after spacing is normalized)
RIG_RENAMES = {
    "PDSI #04.3/N110M": "PDSI #04.3/N110-M",
}

RAW_COLUMNS = [
    "No", "Field", "Nama Sumur", "Nama Sumur_2", "Nama Lokasi", "Nama Rig", "Hari ke",
    "Company Man", "HSSE Kord", "Rig Supt", "DSR", "Cost DC", "Est CC", "AFE Cost",
    "Summary Report", "Current Status", "Plan",
]

FINAL_COLUMNS = [
    "Flag", "Region", "Zone", "APH", "Rig Name", "Well Name", "Well Name [2]",
    "Well Type", "Location", "Spud Date", "Release Date", "Status",
    "Status Code [1]", "Status Code [2]", "Summary Report", "Current Status",
    "Next Plan", "Report Date", "Operation Date",
]

# Section headers: "Operation Summary Report :", "Current Status.", "Plan :"
SECTION_HEADERS = [
    (r"^(?:operation summary report|summary report|24 hrs summary)\s*[:.,]?\s*(.*)$", "Summary Report"),
    (r"^current status\s*[:.,]?\s*(.*)$", "Current Status"),
    (r"^plan\s*:\s*(.*)$", "Plan"),
]

# Lines that end a multi-line section
SECTION_END = (
    r"^(?:dsr|cost dc|est\.?\s*cc|afe\b|note\s*:|casing setting depth|company man|hsse kord"
    r"|rig supt|nama (?:rig|lokasi|sumur)|salam|drilling zona|\d+\.\s)"
)


def clean_line(raw_line):
    """Remove invisible characters and asterisks, then trim spaces."""
    return re.sub(r"[\u200e\u200f\ufeff*]", "", raw_line).strip()


def split_well_names(desc):
    """Split "SKW-041(SKW-B26)" into (SKW-041, SKW-B26); both are the same if no brackets."""
    match = re.match(r"^(.+?)\s*\((.+?)\)?\s*$", desc)
    if match:
        return match.group(1).strip(), match.group(2).strip()
    return desc.strip(), desc.strip()


def parse_money(text):
    """Convert "US$ 1,014,133.53 (9.07%)" or "US$ 1.292.656,75" to a float."""
    number = re.sub(r"[^\d.,]", "", text.split("(")[0]).strip(".,")
    if not number:
        return None
    # A last separator with 1-2 digits after it is the decimal point, all others are thousands
    decimal_match = re.search(r"[.,](\d{1,2})$", number)
    if decimal_match:
        whole_part = re.sub(r"[.,]", "", number[:decimal_match.start()])
        return float(f"{whole_part}.{decimal_match.group(1)}")
    return float(re.sub(r"[.,]", "", number))


def clean_text_for_excel(text):
    """Remove leading '-', '=', '+', '@', ':' so Excel does not read the text as a formula."""
    if pd.isna(text):
        return text
    return re.sub(r"^[-=+@:\s]+", "", str(text)).strip()


def parse_txt_file(path):
    """
    Parse Zone 11 daily drilling report .txt file and extract well data.

    Handles:
    - FIELD sections (e.g. FIELD SUKOWATI)
    - Well entries with optional second name: "1. SKW-041(SKW-B26)"
    - Day count in "H: 04; ..." or "Rig Moving & Rig Up (MIRU), H: 7"
    - Multi-line sections: Operation Summary Report, Current Status, Plan

    Args:
        path: Path to the .txt file

    Returns:
        DataFrame with well information
    """
    content = Path(path).read_text(encoding="utf-8")

    rows = []
    current_field = None
    current_item = None
    current_section = None
    section_lines = []

    def save_section():
        """Join collected lines into the current well's section field."""
        if current_item and current_section and section_lines:
            current_item[current_section] = "\n".join(section_lines)

    for raw in content.splitlines():
        line = clean_line(raw)
        if not line:
            continue
        low = line.lower()

        # FIELD header
        if low.startswith("field"):
            save_section()
            current_section, section_lines = None, []
            names = [p.strip() for p in re.split(r"field", line, flags=re.I) if p.strip()]
            current_field = names[0] if names else None
            continue

        # New well: "1. WELL-NAME (OPTIONAL-NAME)"
        well_match = re.match(r"^(\d+)\.\s*(.+)$", line)
        if well_match:
            save_section()
            current_section, section_lines = None, []
            if current_item:
                rows.append(current_item)
            name_1, name_2 = split_well_names(well_match.group(2))
            current_item = dict.fromkeys(RAW_COLUMNS)
            current_item.update({
                "No": int(well_match.group(1)),
                "Field": current_field,
                "Nama Sumur": name_1,
                "Nama Sumur_2": name_2,
            })
            continue

        # Skip everything before the first well (title, highlights)
        if current_item is None:
            continue

        # Section header: start collecting (keep any text after the header)
        header = next(
            ((name, re.match(pattern, line, re.I)) for pattern, name in SECTION_HEADERS
             if re.match(pattern, line, re.I)),
            None,
        )
        if header:
            save_section()
            current_section, section_lines = header[0], []
            first_text = header[1].group(1).strip()
            if first_text:
                section_lines.append(first_text)
            continue

        # Collect lines until the next label
        if current_section and not re.match(SECTION_END, low):
            section_lines.append(line)
            continue
        if current_section:
            save_section()
            current_section, section_lines = None, []

        # Day count: "H: 04; D: ..." or "Rig Moving & Rig Up (MIRU), H: 7"
        day_match = re.match(r"^(?:.*,\s*)?H\s*:\s*(\d+)", line)
        if day_match:
            current_item["Hari ke"] = int(day_match.group(1))
            continue

        # Key: Value lines
        kv = re.match(r"^([^:]+?)\s*:\s*(.+)$", line)
        if not kv:
            continue
        key, val = kv.group(1).strip().lower(), kv.group(2).strip()

        if key == "nama lokasi":
            current_item["Nama Lokasi"] = val
        elif key == "nama rig":
            current_item["Nama Rig"] = re.sub(r"^rig\s*", "", val, flags=re.I).strip()
        elif key == "company man":
            current_item["Company Man"] = val
        elif key == "hsse kord":
            current_item["HSSE Kord"] = val
        elif key == "rig supt":
            current_item["Rig Supt"] = val
        elif key.startswith("dsr"):
            current_item["DSR"] = val
        elif key == "cost dc":
            current_item["Cost DC"] = parse_money(val)
        elif key.startswith("est"):
            current_item["Est CC"] = parse_money(val)
        elif key.startswith("afe"):
            current_item["AFE Cost"] = parse_money(val)

    # Append last item
    save_section()
    if current_item:
        rows.append(current_item)

    return pd.DataFrame(rows, columns=RAW_COLUMNS)


def transform_txt_to_raw(date_str, base_path="."):
    """
    Transform daily report .txt file to export-raw .xlsx file.

    Args:
        date_str: Date string in format "YYYY-MM-DD"
        base_path: Base path containing daily-report and export-raw folders

    Returns:
        Path to the created Excel file
    """
    base_path = Path(base_path)
    txt_file_path = base_path / "daily-report" / f"{date_str}.txt"
    xlsx_file_path = base_path / "export-raw" / f"{date_str}.xlsx"

    if not txt_file_path.exists():
        raise FileNotFoundError(f"Daily report file not found: {txt_file_path}")

    xlsx_file_path.parent.mkdir(parents=True, exist_ok=True)

    print(f"Reading daily report from: {txt_file_path}")
    df = parse_txt_file(txt_file_path)

    if df.empty:
        raise ValueError("No well data found in the report file")

    print(f"Exporting to Excel: {xlsx_file_path}")
    df.to_excel(xlsx_file_path, index=False, engine="openpyxl")

    print(f"Successfully transformed {len(df)} wells to raw Excel format")
    return xlsx_file_path


def transform_raw_to_final(date_str, base_path="."):
    """
    Transform export-raw .xlsx to export-final .xlsx with standardized columns.

    Fixed values: Region 4, Zone 11, APH PEP, Development, Onshore.

    Args:
        date_str: Date string in format "YYYY-MM-DD"
        base_path: Base path containing export-raw and export-final folders

    Returns:
        Path to the created Excel file
    """
    base_path = Path(base_path)
    raw_file_path = base_path / "export-raw" / f"{date_str}.xlsx"
    final_file_path = base_path / "export-final" / f"{date_str}.xlsx"

    if not raw_file_path.exists():
        raise FileNotFoundError(f"Export-raw file not found: {raw_file_path}")

    final_file_path.parent.mkdir(parents=True, exist_ok=True)

    print(f"Reading export-raw file from: {raw_file_path}")
    df = pd.read_excel(raw_file_path, engine="openpyxl")

    # Select required columns from raw
    df = df[["Nama Sumur", "Nama Sumur_2", "Nama Rig", "Summary Report", "Current Status", "Plan"]].copy()
    df = df.rename(columns={
        "Nama Sumur": "Well Name",
        "Nama Sumur_2": "Well Name [2]",
        "Nama Rig": "Rig Name",
        "Plan": "Next Plan",
    })

    # Add standard columns
    df = df.assign(**{
        "Flag": "INC",
        "Region": "Region 4",
        "Zone": "Zone 11",
        "APH": "PEP",
        "Well Type": "Development",
        "Location": "Onshore",
        "Spud Date": np.nan,
        "Release Date": np.nan,
        "Status": np.nan,
        "Status Code [1]": np.nan,
        "Status Code [2]": np.nan,
    })

    # Fill Well Name [2] with Well Name if empty
    df["Well Name [2]"] = df["Well Name [2]"].fillna(df["Well Name"])

    # Normalize rig name: "PDSI#40.3/ DS-1500E" -> "PDSI #40.3/DS-1500E"
    df["Rig Name"] = (
        df["Rig Name"].astype(str)
        .str.replace(r"PDSI\s*#", "PDSI #", regex=True)
        .str.replace(r"\s*/\s*", "/", regex=True)
        .str.strip()
        .replace(RIG_RENAMES)
    )

    # Clean text fields (Excel-safe)
    for column in ["Summary Report", "Current Status", "Next Plan"]:
        df[column] = df[column].apply(clean_text_for_excel)

    # Add date columns
    report_date = pd.to_datetime(date_str)
    df["Report Date"] = report_date.date()
    df["Operation Date"] = (report_date - pd.Timedelta(days=1)).date()

    # Reorder columns and sort by Rig Name (case-insensitive)
    df = df[FINAL_COLUMNS]
    df = df.sort_values(by="Rig Name", key=lambda s: s.str.lower()).reset_index(drop=True)

    print(f"Exporting final version to Excel: {final_file_path}")
    df.to_excel(final_file_path, index=False, engine="openpyxl")

    print(f"Successfully transformed {len(df)} wells to final Excel format")

    # Copy to clipboard (values only, no header) - works when run locally
    print("Copying data to clipboard...")
    try:
        df.to_clipboard(index=False, header=False)
        print("Data copied to clipboard successfully!")
    except Exception as e:
        print(f"Could not copy to clipboard: {e}")

    # Output clipboard data with markers for the Streamlit app to capture
    clipboard_data = df.to_csv(sep="\t", index=False, header=False)
    print("CLIPBOARD_DATA_START")
    print(clipboard_data, end="")
    print("CLIPBOARD_DATA_END")

    return final_file_path


def main():
    """Main function to handle CLI arguments and execute transformation."""
    if len(sys.argv) < 2:
        print("Usage: python app.py <date>")
        print("Example: python app.py 2026-11-08")
        sys.exit(1)

    date_str = sys.argv[1]

    # Validate date format
    if not re.match(r"^\d{4}-\d{2}-\d{2}$", date_str):
        print("Error: Date must be in format YYYY-MM-DD (e.g., 2026-11-08)")
        sys.exit(1)

    try:
        # Step 1: Transform .txt to export-raw .xlsx
        raw_output_path = transform_txt_to_raw(date_str, SCRIPT_DIR)

        # Step 2: Transform export-raw to export-final .xlsx
        final_output_path = transform_raw_to_final(date_str, SCRIPT_DIR)

        print("\nTransformation completed successfully!")
        print(f"Raw export: {raw_output_path}")
        print(f"Final export: {final_output_path}")

    except (FileNotFoundError, ValueError) as e:
        print(f"Error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
