# Region 2 Daily Drilling Report Processor
# Converts daily report Excel to standardized format

from pathlib import Path
import warnings

import pandas as pd

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

SCRIPT_DIR = Path(__file__).parent

COLUMN_MAPPING = {
    "Report Date": "Report Date",
    "Region": "Region",
    "Zone": "Zone",
    "Unit Name": "Rig Name",
    "Well Name/ Location": "Well Name_1",
    "Job Type": "Well Type",
    "Summary": "Summary Report",
    "Next Plan": "Next Plan",
}

COLUMN_ORDER = [
    "Flag", "Region", "Zone", "APH", "Rig Name", "Well Name_1", "Well Name_2",
    "Well Type", "Location", "Spud Date", "Release Date", "Status",
    "Status Code [1]", "Status Code [2]", "Summary Report", "Current Status",
    "Next Plan", "Report Date", "Operation Date",
]


def convert_daily_report(input_file):
    """
    Convert Region 2 daily report Excel to standardized format.

    Args:
        input_file: Path (str/Path), or file-like object (BytesIO).

    Returns:
        (df, report_date): Processed DataFrame and report date (max from data).

    Raises:
        ValueError: If sheet not found or required columns missing.
    """
    try:
        df = pd.read_excel(
            input_file,
            sheet_name="Bor Report Region 02",
            header=5,
        )
    except ValueError as e:
        source = getattr(input_file, "name", str(input_file))
        raise ValueError(
            f"Could not find sheet 'Bor Report Region 02' in {source}. {e}"
        ) from e

    missing = [k for k in COLUMN_MAPPING.keys() if k not in df.columns]
    if missing:
        raise ValueError(
            f"Missing required columns in Excel: {missing}. "
            f"Available: {list(df.columns)}"
        )

    df = df[list(COLUMN_MAPPING.keys())].rename(columns=COLUMN_MAPPING)
    df = df[df["Zone"].isin(["Zone_05", "Zone_06"])]

    df["Well Type"] = df["Well Type"].replace(
        {"BOR EKS": "Exploration", "BOR DEV": "Development"}
    )
    df["Region"] = df["Region"].replace({"Reg_02": "Region 2"})
    df["Zone"] = df["Zone"].replace({"Zone_05": "Zone 5", "Zone_06": "Zone 6"})

    df["APH"] = df["Zone"].map({"Zone 5": "ONWJ", "Zone 6": "OSES"})
    df["Well Name_2"] = df["Well Name_1"]

    df["Report Date"] = pd.to_datetime(df["Report Date"])
    df["Operation Date"] = df["Report Date"] - pd.Timedelta(days=1)

    df["Rig Name"] = df["Rig Name"].replace("PVD-I", "PVD-II")

    df["Flag"] = "INC"
    df["Location"] = "Offshore"
    df["Spud Date"] = ""
    df["Release Date"] = ""
    df["Status"] = ""
    df["Status Code [1]"] = ""
    df["Status Code [2]"] = ""
    df["Current Status"] = ""

    df = df.reindex(columns=COLUMN_ORDER)
    df = df.sort_values(by=["Zone", "Rig Name"]).reset_index(drop=True)

    report_date = df["Report Date"].max().date()
    return df, report_date


if __name__ == "__main__":
    input_file = SCRIPT_DIR / "daily-report/02-2026/sample.xlsx"
    if input_file.exists():
        df, report_date = convert_daily_report(input_file)
        export_dir = SCRIPT_DIR / "export-final"
        export_dir.mkdir(exist_ok=True)
        output_path = export_dir / f"{report_date:%Y-%m-%d}.xlsx"
        df.to_excel(output_path, index=False)
        print(f"Exported to {output_path}")
    else:
        print(f"Sample file not found: {input_file}")
