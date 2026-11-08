# Region 1 Daily Drilling Report Processor
# Converts daily report Excel to standardized format and exports to export-final/

from pathlib import Path
from datetime import date
from io import BytesIO

import pandas as pd
import re
import sys
import warnings

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

SCRIPT_DIR = Path(__file__).parent

# Rig Name fixes
RIG_RENAMES = {
    "SO PEACE": "SINOOCEAN PEACE",
    "NYT19": "NYT-19",
    "EMSCO D2-M": "PDSI #20.2/EMSCOD2-M",
    "PDSI #09.2/N80UE-25 (1000HP)": "PDSI #09.2/N80UE-25",
    "PDSI#04.3/N110-M": "PDSI #04.3/N110-M",
    "PDSI#40.3": "PDSI #40.3",
    "TMMJ#05": "TMMJ-05",
    "PDSI#12.3/N110-M": "PDSI #12.3/N110-M",
    "GWDC-207": "GW-207",
    "PDSI#15.3/N110-M": "PDSI #15.3/N110-M",
    "PDSI#20.2/EMSCOD2-M": "PDSI #20.2/EMSCOD2-M",
    "PDSI#20.2/EMSCO D2-M": "PDSI #20.2/EMSCOD2-M",
    "Rig PDSI#12.3/N110-M": "PDSI #12.3/N110-M",
    "PDSI#21.2/OW700-M": "PDSI #21.2/OW 700-M",
    "ARJUNA#28": "ARJ#28",
    "PDSI #43.3 /AB-1500E": "PDSI #43.3/AB-1500E",
    "PDSI #31.03/D1500-E": "PDSI #31.3/D1500-E",
}

# Helper functions ( Zone 2&3, Zone 4 )


def split_well_name_z23(val):
    s = str(val).strip() if pd.notna(val) else ""
    if not s:
        return "", "", ""
    m = re.match(r"^([^\n]*)\n?\(([^)]*)\)\s*\n?\s*\(([^)]*)\)\s*$", s)
    if m:
        p1, p2, p3 = m.group(1).strip(), m.group(2).strip(), m.group(3).strip()
    else:
        m2 = re.match(r"^([^\n]*)\n?\(([^)]*)\)\s*$", s)
        if m2:
            p1, p2, p3 = m2.group(1).strip(), m2.group(2).strip(), ""
        else:
            p1, p2, p3 = s, "", ""
    if not p1 and p2:
        p1 = p2
    if not p2 and p1:
        p2 = p1
    if not p3 and p2:
        p3 = p2
    return p1, p2, p3


def split_summary_report_z23(val):
    s = str(val).strip() if pd.notna(val) else ""
    if not s:
        return "", "", ""
    s = s.replace("_x000D_", "\n")
    summary = ""
    # "Laporan:" label is optional (some reports start directly with the text)
    # "Rencana:" or "Plan:" starts the next plan; matching is case-insensitive
    flags = re.DOTALL | re.IGNORECASE
    m1 = re.search(r"Laporan:\s*(.*?)(?=Status Pagi|(?:Rencana|Plan)\s*:|$)", s, flags)
    if not m1:
        m1 = re.search(r"^(.*?)(?=Status Pagi|(?:Rencana|Plan)\s*:|$)", s, flags)
    if m1:
        summary = m1.group(1).strip()
        if summary.startswith("-"):
            summary = summary[1:].strip()
    status = ""
    # Time label is optional and varies: "05:00", "@05:00", "@05.00 :", "@05:00 hrs:"
    m2 = re.search(
        r"Status Pagi\s*(?:@?\s*\d{1,2}[:.]\d{2}\s*(?:hrs)?\s*)?:?\s*(.*?)(?=(?:Rencana|Plan)\s*:|$)",
        s,
        flags,
    )
    if m2:
        status = m2.group(1).strip()
    plan = ""
    m3 = re.search(r"(?:Rencana|Plan)\s*:\s*(.*)$", s, flags)
    if m3:
        plan = m3.group(1).strip()
    return summary, status, plan


def split_well_name_z4(val):
    s = str(val).strip() if pd.notna(val) else ""
    if not s:
        return "", ""
    # Closing ")" is optional (some reports forget it)
    m = re.match(r"^(.*?)\s*\(([^)]*)\)?\s*$", s)
    if m:
        p1, p2 = m.group(1).strip(), m.group(2).strip()
        if not p1:
            p1 = p2
        if not p2:
            p2 = p1
    else:
        p1, p2 = s, s
    return p1, p2


def split_well_name_z1(val):
    s = str(val).strip() if pd.notna(val) else ""
    if not s:
        return "", ""
    # Old format "A / B", new format "A (B)" or "(B)"
    if "/" in s and "(" not in s:
        p1, p2 = s.split("/", 1)
        return p1.strip(), p2.strip()
    return split_well_name_z4(s)


def split_summary_report_z4(val):
    s = str(val).strip() if pd.notna(val) else ""
    if not s:
        return "", "", ""
    s = s.replace("_x000D_", "\n")
    summary = ""
    m1 = re.search(r"^(.*?)(?=Status Pagi|Plan\s*:|$)", s, re.DOTALL | re.IGNORECASE)
    if m1:
        summary = m1.group(1).strip()
    status = ""
    # Colon and time label after "Status Pagi" are optional
    m2 = re.search(
        r"Status Pagi\s*(?:@?\s*\d{1,2}[:.]\d{2}\s*(?:hrs)?\s*)?:?\s*(.*?)(?=Plan\s*:|$)",
        s,
        re.DOTALL | re.IGNORECASE,
    )
    if m2:
        status = m2.group(1).strip()
    plan = ""
    m3 = re.search(r"Plan\s*:\s*(.*)$", s, re.DOTALL | re.IGNORECASE)
    if m3:
        plan = m3.group(1).strip()
    return summary, status, plan


# First 3 letters of month names (English + Indonesian) -> month number
MONTH_NUMBERS = {
    "jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "mei": 5, "jun": 6,
    "jul": 7, "aug": 8, "agu": 8, "agt": 8, "sep": 9, "oct": 10, "okt": 10,
    "nov": 11, "dec": 12, "des": 12,
}


def parse_report_date(text, default_year=None, from_filename=False):
    """
    Parse "4 November 2026", "01 Jan", "12 Maret 2026" into a date (or None).

    from_filename=True: find "tanggal DD Month YYYY" inside a file name.
    default_year: year to use when the text has no year (e.g. sheet names).
    """
    text = str(text)
    if from_filename:
        match = re.search(r"tanggal\s+(\d{1,2})\s+([A-Za-z]{3,9})\s+(\d{4})", text, re.IGNORECASE)
    else:
        match = re.fullmatch(r"\s*(\d{1,2})\s+([A-Za-z]{3,9})(?:\s+(\d{4}))?\s*", text)
    if not match:
        return None
    day, month_name, year = match.groups()
    month = MONTH_NUMBERS.get(month_name[:3].lower())
    year = int(year) if year else default_year
    if month is None or year is None:
        return None
    try:
        return date(year, month, int(day))
    except ValueError:
        return None


def convert_daily_report(input_file, report_date=None):
    """
    Convert Region 1 daily report Excel to standardized format.

    Args:
        input_file: Path (str/Path), or file-like object (BytesIO, Streamlit upload).
        report_date: Optional date. If None, extracted from filename
            (pattern: tanggal DD Mon YYYY, English or Indonesian month names).
            The sheet named after this date ("01 Jan", "1 Mar", "4 November") is read.

    Returns:
        (df_merged, report_date): Processed DataFrame and report date.

    Raises:
        ValueError: If date cannot be determined or sheet not found.
    """
    source_name = getattr(input_file, "name", None) or (
        Path(input_file).name if isinstance(input_file, (Path, str)) else None
    )

    # Report date: argument first, otherwise from filename
    if report_date is None:
        if not source_name:
            raise ValueError(
                "Could not extract date from filename. Provide report_date."
            )
        report_date = parse_report_date(source_name, from_filename=True)
        if report_date is None:
            raise ValueError(
                f"Could not extract date from filename: {source_name}. "
                "Filename must match 'tanggal DD Mon YYYY' (e.g. 'tanggal 4 November 2026') "
                "or provide report_date."
            )
    else:
        report_date = pd.to_datetime(report_date).date()

    # Find the sheet named after the report date ("01 Jan", "1 Mar", "4 November", ...)
    excel_file = pd.ExcelFile(input_file)
    sheet_dates = pd.Series(excel_file.sheet_names, dtype=str).map(
        lambda name: parse_report_date(name, default_year=report_date.year)
    )
    matching_sheets = sheet_dates.index[sheet_dates == report_date]
    if len(matching_sheets) == 0:
        raise ValueError(
            f"Could not find sheet for {report_date} in {source_name or input_file}. "
            f"Available sheets: {excel_file.sheet_names}"
        )
    sheet_name = excel_file.sheet_names[matching_sheets[0]]
    raw = excel_file.parse(sheet_name, header=None)

    # Warn if the date written in the sheet title differs from the report date
    title_dates = raw.iloc[:8].astype(str).stack().map(parse_report_date).dropna()
    if len(title_dates) > 0 and title_dates.iloc[0] != report_date:
        warnings.warn(
            f"Sheet title date ({title_dates.iloc[0]}) differs from report date ({report_date})."
        )

    # Find header row dynamically (the row that contains the "Zona" column title)
    header_rows = raw.index[raw.astype(str).apply(lambda col: col.str.strip()).eq("Zona").any(axis=1)]
    if len(header_rows) == 0:
        raise ValueError(f"Could not find header row (column 'Zona') in sheet '{sheet_name}'.")
    header_row = header_rows[0]
    df = raw.iloc[header_row + 1 :].copy()
    df.columns = [str(col).strip() for col in raw.iloc[header_row]]
    df = df.reset_index(drop=True)

    # Set date columns as pandas datetime objects
    df["Report Date"] = pd.to_datetime(report_date)
    df["Operation Date"] = pd.to_datetime(report_date) - pd.Timedelta(days=1)

    selected_columns = [
        "Zona",
        "Nama Sumur",
        "RIG",
        "Jenis Kegiatan",
        "Kegiatan",
        "Report Date",
        "Operation Date",
    ]
    missing_columns = [col for col in selected_columns if col not in df.columns]
    if missing_columns:
        raise ValueError(
            f"Missing columns in sheet '{sheet_name}': {missing_columns}. "
            f"Found: {list(df.columns)}"
        )
    df = df[selected_columns]

    df = df[
        df["Zona"].isin(["Zona 1", "Zona 2 & 3", "Zona 4", "Non Operator", "KSO"])
        & df["Zona"].notna()
    ]
    if df.empty:
        raise ValueError(f"No well rows found in sheet '{sheet_name}'.")
    df["Zona"] = df["Zona"].replace(
        {"Zona 1": "Zone 1", "Zona 2 & 3": "Zone 2&3", "Zona 4": "Zone 4", "Non Operator": "Non-Operator", "KSO": "Non-Operator"}
    )
    df = df.rename(
        columns={
            "Zona": "Zone",
            "Nama Sumur": "Well Name",
            "RIG": "Rig Name",
            "Jenis Kegiatan": "Well Type",
            "Kegiatan": "Summary Report",
        }
    )

    reference_columns = [
        "Flag",
        "Region",
        "Zone",
        "APH",
        "Rig Name",
        "Well Name",
        "Well Name [2]",
        "Well Type",
        "Location",
        "Spud Date",
        "Release Date",
        "Status",
        "Status Code [1]",
        "Status Code [2]",
        "Summary Report",
        "Current Status",
        "Next Plan",
        "Report Date",
        "Operation Date",
    ]

    for col in reference_columns:
        if col not in df.columns:
            df[col] = ""

    df = df[reference_columns]
    df["Flag"] = "INC"
    df["Region"] = "Region 1"
    df["Location"] = "Onshore"

    df["APH"] = df["Zone"].map({"Zone 1": "PEP", "Zone 4": "PEP", "Zone 2&3": "PHR"})
    df["Well Type"] = df["Well Type"].replace("Eksplorasi", "Exploration")

    # Zone 1
    df_z1 = df[df["Zone"] == "Zone 1"].copy()

    split_result = df_z1["Well Name"].apply(split_well_name_z1)
    df_z1["Well Name"] = split_result.apply(lambda x: x[0])
    df_z1["Well Name [2]"] = split_result.apply(lambda x: x[1])

    split_result = df_z1["Summary Report"].apply(split_summary_report_z4)
    df_z1["Summary Report"] = split_result.apply(lambda x: x[0])
    df_z1["Current Status"] = split_result.apply(lambda x: x[1])
    df_z1["Next Plan"] = split_result.apply(lambda x: x[2])

    df_z1["Rig Name"] = df_z1["Rig Name"].str.replace("Rig", "", regex=False).str.strip()
    df_z1.reset_index(drop=True, inplace=True)

    # Zone 2&3
    df_z23 = df[df["Zone"] == "Zone 2&3"].copy()
    split_result = df_z23["Well Name"].apply(split_well_name_z23)
    df_z23["Well Name"] = split_result.apply(lambda x: x[1])
    df_z23["Well Name [2]"] = split_result.apply(lambda x: x[2])
    df_z23 = df_z23.sort_values(by="Rig Name")

    split_result = df_z23["Summary Report"].apply(split_summary_report_z23)
    df_z23["Summary Report"] = split_result.apply(lambda x: x[0])
    df_z23["Current Status"] = split_result.apply(lambda x: x[1])
    df_z23["Next Plan"] = split_result.apply(lambda x: x[2])

    # Zone 4
    df_z4 = df[df["Zone"] == "Zone 4"].copy()
    df_z4["Well Name"] = df_z4["Well Name"].str.replace("\u2060", "", regex=False)

    split_result = df_z4["Well Name"].apply(split_well_name_z4)
    df_z4["Well Name"] = split_result.apply(lambda x: x[0])
    df_z4["Well Name [2]"] = split_result.apply(lambda x: x[1])
    df_z4.reset_index(drop=True, inplace=True)

    df_z4["Rig Name"] = df_z4["Rig Name"].str.replace("Rig", "", regex=False).str.strip()
    df_z4["Rig Name"] = (
        df_z4["Rig Name"]
        .replace(
            {
                "Airlangga #55": "Airlangga-55",
                "PDSI ACS#21": "ACS-21",
                "#36.1/Skytop 650M": "PDSI #36.1/Skytop 650M",
                "PDSI #32.2 / N80UE-E": "PDSI #32.2/N80UE-E"
            }
        )
        .apply(
            lambda x: re.sub(r"(PDSI #)\s+", r"\1", x)
            if isinstance(x, str) and x.startswith("PDSI #")
            else x
        )
    )
    df_z4 = df_z4.sort_values(by="Rig Name")

    split_result = df_z4["Summary Report"].apply(split_summary_report_z4)
    df_z4["Summary Report"] = split_result.apply(lambda x: x[0])
    df_z4["Current Status"] = split_result.apply(lambda x: x[1])
    df_z4["Next Plan"] = split_result.apply(lambda x: x[2])


    df_partnership = df[df["Zone"] == "Non-Operator"].copy()

    # Replace known names
    df_partnership["Rig Name"] = (
        df_partnership["Rig Name"]
        .replace({
            "EPI#9": "EPI-09",
            "Bohai #26": "Bohai-26",
            "Bohai#26": "Bohai-26",
        })
    )

    # Split Summary Report into 3 columns: Summary Report, Current Status, Next Plan
    split_result = df_partnership["Summary Report"].apply(split_summary_report_z4)
    df_partnership["Summary Report"] = split_result.apply(lambda x: x[0])
    df_partnership["Current Status"] = split_result.apply(lambda x: x[1])
    df_partnership["Next Plan"] = split_result.apply(lambda x: x[2])

    # Split Well Name "A (B)" into Well Name and Well Name [2] (both = name if no brackets)
    split_result = df_partnership["Well Name"].apply(split_well_name_z4)
    df_partnership["Well Name"] = split_result.apply(lambda x: x[0])
    df_partnership["Well Name [2]"] = split_result.apply(lambda x: x[1])

    # Merge all DataFrames
    df_merged = pd.concat([df_z1, df_z23, df_z4, df_partnership], ignore_index=True)


    
    # Normalize date columns to datetime64[ns] for consistent Excel round-trip and filtering
    df_merged["Report Date"] = pd.to_datetime(df_merged["Report Date"], format='mixed')
    df_merged["Operation Date"] = pd.to_datetime(df_merged["Operation Date"], format='mixed')

    df_merged["Summary Report"] = df_merged["Summary Report"].astype(str).str.lstrip("-").str.strip()
    df_merged["Current Status"] = df_merged["Current Status"].astype(str).str.lstrip("-").str.strip()
    df_merged["Next Plan"] = df_merged["Next Plan"].astype(str).str.lstrip("-").str.strip()

    # Rig name replacement
    df_merged["Rig Name"] = df_merged["Rig Name"].replace(RIG_RENAMES)
    # Standardize "TMMJ#18" -> "TMMJ-18" and "GW265" -> "GW-265"
    df_merged["Rig Name"] = (
        df_merged["Rig Name"]
        .str.replace(r"^TMMJ#(\d+)$", r"TMMJ-\1", regex=True)
        .str.replace(r"^GW(\d+)$", r"GW-\1", regex=True)
    )

    # Sort by Zone and Rig Name
    df_merged = df_merged.sort_values(by=["Zone", "Rig Name"]).reset_index(drop=True)
    return df_merged, report_date


if __name__ == "__main__":
    # Usage: python app.py [file1.xlsx file2.xlsx ...] (default: one sample file)
    default_file = SCRIPT_DIR / "daily-report/02-2026/Laporan Harian Pemboran Regional 1 tanggal 12 Feb 2026.xlsx"
    input_files = sys.argv[1:] or [default_file]
    export_dir = SCRIPT_DIR / "export-final"
    export_dir.mkdir(exist_ok=True)
    for input_file in input_files:
        df_merged, report_date = convert_daily_report(input_file)
        output_path = export_dir / f"{report_date:%Y-%m-%d}.xlsx"
        df_merged.to_excel(output_path, index=False)
        print(f"Exported to {output_path}")
