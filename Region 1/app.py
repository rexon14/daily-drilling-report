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
    if not p3 and p2:
        p3 = p2
    return p1, p2, p3


def split_summary_report_z23(val):
    s = str(val).strip() if pd.notna(val) else ""
    if not s:
        return "", "", ""
    s = s.replace("_x000D_", "\n")
    summary = ""
    m1 = re.search(r"Laporan:\s*(.*?)(?=Status Pagi|Rencana:|$)", s, re.DOTALL)
    if m1:
        summary = m1.group(1).strip()
        if summary.startswith("-"):
            summary = summary[1:].strip()
    status = ""
    m2 = re.search(
        r"Status Pagi\s*\d{1,2}:\d{2}\s*:\s*(.*?)(?=Rencana:|$)", s, re.DOTALL
    )
    if m2:
        status = m2.group(1).strip()
    plan = ""
    m3 = re.search(r"Rencana:\s*(.*)$", s, re.DOTALL)
    if m3:
        plan = m3.group(1).strip()
    return summary, status, plan


def split_well_name_z4(val):
    s = str(val).strip() if pd.notna(val) else ""
    if not s:
        return "", ""
    m = re.match(r"^(.+?)\s*\(([^)]*)\)\s*$", s)
    if m:
        p1, p2 = m.group(1).strip(), m.group(2).strip()
        if not p2:
            p2 = p1
    else:
        p1, p2 = s, s
    return p1, p2


def split_summary_report_z4(val):
    s = str(val).strip() if pd.notna(val) else ""
    if not s:
        return "", "", ""
    s = s.replace("_x000D_", "\n")
    summary = ""
    m1 = re.search(r"^(.+?)(?=Status Pagi|Plan:|$)", s, re.DOTALL)
    if m1:
        summary = m1.group(1).strip()
    status = ""
    m2 = re.search(
        r"Status Pagi(?:\s*\d{1,2}:\d{2})?\s*:\s*(.*?)(?=Plan:|$)", s, re.DOTALL
    )
    if m2:
        status = m2.group(1).strip()
    plan = ""
    m3 = re.search(r"Plan:\s*(.*)$", s, re.DOTALL)
    if m3:
        plan = m3.group(1).strip()
    return summary, status, plan


def convert_daily_report(input_file, report_date=None):
    """
    Convert Region 1 daily report Excel to standardized format.

    Args:
        input_file: Path (str/Path), or file-like object (BytesIO).
        report_date: Optional date. If None, extracted from filename (pattern: tanggal DD Mon YYYY).

    Returns:
        (df_merged, report_date): Processed DataFrame and report date.

    Raises:
        ValueError: If date cannot be determined or sheet not found.
    """
    if report_date is None:
        filename = getattr(input_file, "name", None) or (
            Path(input_file).name if isinstance(input_file, (Path, str)) else None
        )
        if not filename:
            raise ValueError(
                "Could not extract date from filename. Provide report_date."
            )
        match = re.search(r"tanggal (\d{1,2}) (\w{3}) (\d{4})", filename)
        if not match:
            raise ValueError(
                f"Could not extract date from filename: {filename}. "
                "Filename must match 'tanggal DD Mon YYYY' or provide report_date."
            )
        day, month, year = match.groups()
        report_date = pd.to_datetime(
            f"{day} {month} {year}", format="%d %b %Y"
        ).date()
    else:
        report_date = pd.to_datetime(report_date).date()

    day_no_zero = str(int(report_date.day))
    day_zero = str(report_date.day).zfill(2)
    month_abbr = report_date.strftime("%b")
    possible_sheet_names = [
        f"{day_no_zero} {month_abbr}",
        f"{day_zero} {month_abbr}",
    ]

    df = None
    for sheet in possible_sheet_names:
        try:
            df = pd.read_excel(input_file, sheet_name=sheet, header=13)
            break
        except Exception:
            continue

    if df is None:
        source = getattr(input_file, "name", str(input_file))
        raise ValueError(
            f"Could not find valid sheet in {source}. "
            f"Expected sheet names: {possible_sheet_names}"
        )

    # Set date columns as pandas datetime objects
    df["Report Date"] = pd.to_datetime(report_date)
    df["Operation Date"] = pd.to_datetime(report_date) - pd.Timedelta(days=1)

    selected_columns = [
        "Zona",
        "Nama Sumur",
        "RIG",
        "Jenis Kegiatan",
        "Kegiatan ",
        "Report Date",
        "Operation Date",
    ]
    df = df[selected_columns]

    df = df[
        df["Zona"].isin(["Zona 1", "Zona 2 & 3", "Zona 4"])
        & df["Zona"].notna()
    ]
    df["Zona"] = df["Zona"].replace(
        {"Zona 1": "Zone 1", "Zona 2 & 3": "Zone 2&3", "Zona 4": "Zone 4"}
    )
    df = df.rename(
        columns={
            "Zona": "Zone",
            "Nama Sumur": "Well Name",
            "RIG": "Rig Name",
            "Jenis Kegiatan": "Well Type",
            "Kegiatan ": "Summary Report",
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

    well_name_split = df_z1["Well Name"].str.split("/", n=1, expand=True)
    df_z1["Well Name"] = well_name_split[0].str.strip()
    df_z1["Well Name [2]"] = (
        well_name_split.get(1, pd.Series([""] * len(df_z1))).fillna("").str.strip()
    )

    summary_split = df_z1["Summary Report"].str.split(
        r"(?i)Plan\s*:\s*", n=1, expand=True, regex=True
    )
    df_z1["Summary Report"] = summary_split[0].fillna("").str.strip()
    df_z1["Next Plan"] = (
        summary_split.get(1, pd.Series([""] * len(df_z1))).fillna("").str.strip()
    )

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

    df_merged = pd.concat([df_z1, df_z23, df_z4], ignore_index=True)


    
    # Normalize date columns to datetime64[ns] for consistent Excel round-trip and filtering
    df_merged["Report Date"] = pd.to_datetime(df_merged["Report Date"], format='mixed')
    df_merged["Operation Date"] = pd.to_datetime(df_merged["Operation Date"], format='mixed')

    # Remove leading "-" and leading/trailing spaces in "Next Plan"
    df_merged["Next Plan"] = df_merged["Next Plan"].astype(str).str.lstrip("-").str.strip()
    return df_merged, report_date


if __name__ == "__main__":
    input_file = SCRIPT_DIR / "daily-report/02-2026/Laporan Harian Pemboran Regional 1 tanggal 12 Feb 2026.xlsx"
    df_merged, report_date = convert_daily_report(input_file)
    export_dir = SCRIPT_DIR / "export-final"
    export_dir.mkdir(exist_ok=True)
    output_path = export_dir / f"{report_date:%Y-%m-%d}.xlsx"
    df_merged.to_excel(output_path, index=False)
    print(f"Exported to {output_path}")
