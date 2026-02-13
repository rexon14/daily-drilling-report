# Region 1 Daily Drilling Report Processor

**Source:** `Region 1/app.ipynb`

## Summary

This Jupyter notebook processes daily drilling report Excel files from Region 1 (Pertamina SHU). It reads multiple `.xlsx` files from a configurable month folder, extracts data from sheets named by date (e.g., `1 Feb`, `01 Feb`), and transforms the raw structure into a standardized format. The data is split and processed differently per zone (Zone 1, Zone 2&3, Zone 4) — each zone has its own parsing rules for well names and summary reports. The final merged result is displayed with a date picker widget and is copied to the clipboard for easy pasting into downstream tools.

## Key Dependencies

| Library | Role |
|---------|------|
| **pandas** | Data loading, manipulation, concatenation |
| **openpyxl** | Excel file reading (used by `pd.read_excel`) |
| **re** | Regex parsing for filename dates, well names, and summary report splits |
| **pathlib** | Path handling for finding Excel files |
| **ipywidgets** | Date picker and navigation buttons for interactive filtering |
| **IPython.display** | Display widgets and DataFrames in notebook |

> **Note:** `ipywidgets` may need to be installed separately (`pip install ipywidgets`). The project root `requirements.txt` includes `pandas`, `openpyxl`, and `streamlit` but not `ipywidgets`.

## Table of Contents

1. [Imports and Configuration](#1-imports-and-configuration)
2. [Load and Process Excel Files](#2-load-and-process-excel-files)
3. [Select and Filter Columns](#3-select-and-filter-columns)
4. [Add Reference Columns and Defaults](#4-add-reference-columns-and-defaults)
5. [Zone 1 Processing](#5-zone-1-processing)
6. [Zone 2&3 Processing](#6-zone-23-processing)
7. [Zone 4 Processing](#7-zone-4-processing)
8. [Merge and Display](#8-merge-and-display)

## Detailed Section Breakdown

### 1. Imports and Configuration

**What it does:** Imports required libraries, suppresses openpyxl warnings, and sets the month folder to read from. Lists all Excel files in that folder (excluding temp files starting with `~$`).

**Code:**

```python
from pathlib import Path
import pandas as pd
import re
import warnings

# Suppress openpyxl warnings
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

# Specify subdirectory to read from (e.g., "01-2026" for January 2026)
month_folder = "02-2026"

# Find all Excel files in the specified subdirectory
daily_report_path = Path("daily-report") / month_folder
excel_files = [f for f in daily_report_path.glob("*.xlsx") 
               if not f.name.startswith("~$")]
```

**How it works:**
- `Path("daily-report") / month_folder` builds the path relative to the notebook's working directory (e.g., `daily-report/02-2026`).
- `glob("*.xlsx")` finds all Excel files; the list comprehension excludes temporary files (`~$...`) that Excel creates when a file is open.

**Key variables/parameters:**
- `month_folder` — change this to switch months (e.g., `"01-2026"`, `"03-2026"`).
- `daily_report_path` — resolved path where Excel files live.
- `excel_files` — list of `Path` objects for each `.xlsx` file.

**Gotchas / Assumptions:**
- The notebook must be run from the `Region 1` folder so `daily-report` resolves correctly.
- Files must follow the naming pattern that includes `tanggal DD MMM YYYY` (see next section).

---

### 2. Load and Process Excel Files

**What it does:** Loops over each Excel file, extracts the date from the filename, finds the matching sheet by date (handling both `1 Feb` and `01 Feb`), reads the sheet with header row 13, and appends `Report Date` and `Operation Date` columns. All DataFrames are concatenated into a single `df`.

**Code:**

```python
# Process each file
dfs = []
for file_path in excel_files:
    # Extract date from filename pattern: "tanggal DD MMM YYYY"
    match = re.search(r"tanggal (\d{1,2}) (\w{3}) (\d{4})", file_path.name)
    
    if match:
        day, month, year = match.groups()
        # Handle both "1 Feb" and "01 Feb" sheet names
        day_no_zero = str(int(day))     # e.g., '1'
        day_zero = day.zfill(2)         # e.g., '01'
        possible_sheet_names = [f"{day_no_zero} {month}", f"{day_zero} {month}"]
        report_date = pd.to_datetime(
            f"{day} {month} {year}", format="%d %b %Y"
        ).date()

        # Try opening each possible sheet name until one works
        df = None
        for sheet in possible_sheet_names:
            try:
                df = pd.read_excel(file_path, sheet_name=sheet, header=13)
                break
            except Exception:
                continue
        if df is not None:
            df["Report Date"] = report_date
            df["Operation Date"] = report_date - pd.Timedelta(days=1)
            dfs.append(df)
        else:
            print(f"Warning: Could not find valid sheet name in {file_path.name}")
    else:
        print(f"Warning: Could not extract date from filename: {file_path.name}")

# Merge all dataframes
if dfs:
    df = pd.concat(dfs, ignore_index=True)
else:
    df = pd.DataFrame()
    print("No dataframes were successfully loaded.")
```

**How it works:**
- Regex `r"tanggal (\d{1,2}) (\w{3}) (\d{4})"` matches filenames like `...tanggal 12 Feb 2026.xlsx`.
- `day_no_zero` / `day_zero` handle both `1 Feb` and `01 Feb` sheet naming conventions.
- `pd.read_excel(..., header=13)` skips the first 13 rows and uses row 14 as the header.
- `Report Date` = date from filename; `Operation Date` = report date minus one day.
- `pd.concat(dfs, ignore_index=True)` merges all daily DataFrames into one.

**Key variables/parameters:**
- `header=13` — fixed offset; source Excel layout must have headers at row 14.
- `Operation Date` — always one day before `Report Date` (business rule).

**Gotchas / Assumptions:**
- Filename must contain `tanggal DD MMM YYYY` (month as 3-letter abbrev, e.g., Feb, Jan).
- Sheet name must match either `D MMM` or `DD MMM` (e.g., `1 Feb` or `01 Feb`).
- If no sheet matches, the file is skipped and a warning is printed.

---

### 3. Select and Filter Columns

**What it does:** Keeps only the needed columns, filters rows to zones 1, 2&3, and 4, normalizes zone names, and renames columns to English.

**Code:**

```python
# Keep only needed columns
selected_columns = ["Zona", "Nama Sumur", "RIG", "Jenis Kegiatan", "Kegiatan ", "Report Date", "Operation Date"]
df = df[selected_columns]

# Filter rows where Zona is in specified values and not NA
df = df[df['Zona'].isin(['Zona 1', 'Zona 2 & 3', 'Zona 4']) & df['Zona'].notna()]

# Rename Zona values for consistency
df['Zona'] = df['Zona'].replace({'Zona 1': 'Zone 1', 'Zona 2 & 3': 'Zone 2&3', 'Zona 4': 'Zone 4'})

# Rename columns for clarity and consistency
df = df.rename(columns={
    'Zona': 'Zone',
    'Nama Sumur': 'Well Name',
    'RIG': 'Rig Name',
    'Jenis Kegiatan': 'Well Type',
    'Kegiatan ': 'Summary Report'
})
```

**How it works:**
- Drops all columns except the seven listed. Note: `"Kegiatan "` has a trailing space — exact match required.
- `isin()` keeps only rows for Zona 1, 2&3, and 4; `notna()` drops missing zones.
- Zone labels are standardized to `Zone 1`, `Zone 2&3`, `Zone 4`.
- Column names are mapped from Indonesian to English for downstream use.

**Key variables/parameters:**
- `selected_columns` — must match the source Excel column names exactly.
- Zone filter — rows outside these three zones are discarded.

**Gotchas / Assumptions:**
- Source must have `Kegiatan ` (with trailing space) for the summary report column.
- Only Zone 1, 2&3, and 4 are processed; other zones are excluded.

---

### 4. Add Reference Columns and Defaults

**What it does:** Ensures all reference columns exist (adding blanks for missing ones), reorders columns, sets default values for `Flag`, `Region`, and `Location`, maps `APH` from `Zone`, and normalizes `Well Type` (Eksplorasi → Exploration).

**Code:**

```python
# Reference columns from the image
reference_columns = [
    "Flag", "Region", "Zone", "APH", "Rig Name", "Well Name", "Well Name [2]",
    "Well Type", "Location", "Spud Date", "Release Date", "Status",
    "Status Code [1]", "Status Code [2]", "Summary Report", "Current Status",
    "Next Plan", "Report Date", "Operation Date"
]

# Add any missing columns as blank
for col in reference_columns:
    if col not in df.columns:
        df[col] = ""  # blank

# Reorder according to the reference
df = df[reference_columns]

# Set default values for Flag and Region columns
df['Flag'] = "INC"
df['Region'] = "Region 1"
df["Location"] = "Onshore"
```

```python
# Fill APH column based on Zone values
df["APH"] = df["Zone"].map({"Zone 1": "PEP", "Zone 4": "PEP", "Zone 2&3": "PHR"})

# Replace "Eksplorasi" with "Exploration" in Well Type column
df["Well Type"] = df["Well Type"].replace("Eksplorasi", "Exploration")
```

**How it works:**
- `reference_columns` defines the target schema. Any column not yet present is added as an empty string.
- `df = df[reference_columns]` enforces the exact column order.
- `Flag` = `"INC"`, `Region` = `"Region 1"`, `Location` = `"Onshore"` for all rows.
- `APH` is derived from `Zone`: Zone 1 and 4 → `PEP`; Zone 2&3 → `PHR`.
- `Well Type` replaces Indonesian `Eksplorasi` with `Exploration`.

**Key variables/parameters:**
- `Flag` — all `"INC"` (likely "Inclusive" or similar).
- `APH` mapping — Zone 1 & 4 → PEP; Zone 2&3 → PHR.

**Gotchas / Assumptions:**
- `Well Name [2]` is added here as blank; it gets populated during zone-specific processing.
- `Spud Date`, `Release Date`, `Status`, `Status Code [1]`, `Status Code [2]` remain blank unless filled elsewhere.

---

### 5. Zone 1 Processing

**What it does:** Filters Zone 1 rows, splits `Well Name` on the first `/` into `Well Name` and `Well Name [2]`, splits `Summary Report` on `Plan:` into `Summary Report` and `Next Plan`, and strips `Rig` from `Rig Name`.

**Code:**

```python
df_z1 = df[df["Zone"] == "Zone 1"].copy()

# Split Well Name on first "/"
well_name_split = df_z1["Well Name"].str.split("/", n=1, expand=True)
df_z1["Well Name"] = well_name_split[0].str.strip()
df_z1["Well Name [2]"] = well_name_split.get(1, pd.Series([""] * len(df_z1))).fillna("").str.strip()
```

```python
# Split Summary Report into Summary and Plan (case-insensitive "Plan:" or "Plan :")
summary_split = df_z1["Summary Report"].str.split(r"(?i)Plan\s*:\s*", n=1, expand=True, regex=True)
df_z1["Summary Report"] = summary_split[0].fillna("").str.strip()
df_z1["Next Plan"] = summary_split.get(1, pd.Series([""] * len(df_z1))).fillna("").str.strip()
```

```python
# Remove "Rig" and extra spaces from "Rig Name" column
df_z1["Rig Name"] = df_z1["Rig Name"].str.replace("Rig", "", regex=False).str.strip()
```

```python
# Reset Index
df_z1.reset_index(drop=True, inplace=True)
```

**How it works:**
- `str.split("/", n=1, expand=True)` splits on the first `/`; part before → `Well Name`, part after → `Well Name [2]`.
- `(?i)Plan\s*:\s*` matches `Plan:`, `Plan :`, `plan:`, etc. Part before → `Summary Report`, part after → `Next Plan`.
- `str.replace("Rig", "", regex=False)` removes the literal `Rig` prefix from rig names (e.g., `Rig APS-752` → `APS-752`).

**Key variables/parameters:**
- Zone 1 `Well Name` format: `Part1/Part2` (e.g., `RNT-DZ51/P-475`).
- Zone 1 `Summary Report` format: free text then `Plan:` or `Plan :` then plan text.

**Gotchas / Assumptions:**
- Zone 1 does not use `Current Status`; it stays blank.
- `Well Name [2]` is empty when there is no `/` in the original name.

---

### 6. Zone 2&3 Processing

**What it does:** Filters Zone 2&3 rows, parses `Well Name` with regex (handles `Part1\n(Part2)\n(Part3)` or `Part1\n(Part2)`), maps Part 2 → `Well Name`, Part 3 → `Well Name [2]`, splits `Summary Report` on `Laporan:`, `Status Pagi HH:MM:`, and `Rencana:`, sorts by `Rig Name`.

**Code:**

```python
# Split Well Name into 3 columns: Part1, Part2 (first paren), Part3 (second paren)
def split_well_name_z23(val):
    s = str(val).strip() if pd.notna(val) else ""
    if not s:
        return "", "", ""
    # Pattern: Part1\n(Part2)\n (Part3)
    m = re.match(r"^([^\n]*)\n?\(([^)]*)\)\s*\n?\s*\(([^)]*)\)\s*$", s)
    if m:
        p1, p2, p3 = m.group(1).strip(), m.group(2).strip(), m.group(3).strip()
    else:
        # Try single paren: Part1\n(Part2)
        m2 = re.match(r"^([^\n]*)\n?\(([^)]*)\)\s*$", s)
        if m2:
            p1, p2, p3 = m2.group(1).strip(), m2.group(2).strip(), ""
        else:
            p1, p2, p3 = s, "", ""
    # Fallback: if Part 1 or Part 3 blank, copy from Part 2
    if not p1 and p2:
        p1 = p2
    if not p3 and p2:
        p3 = p2
    return p1, p2, p3

df_z23 = df[df["Zone"] == "Zone 2&3"].copy()
split_result = df_z23["Well Name"].apply(split_well_name_z23)
# Part 2 -> Well Name, Part 3 -> Well Name [2] (drop Part 1)
df_z23["Well Name"] = split_result.apply(lambda x: x[1])
df_z23["Well Name [2]"] = split_result.apply(lambda x: x[2])

# Sort by column Rig Name
df_z23 = df_z23.sort_values(by="Rig Name")
```

```python
# Split Summary Report into 3 columns: Summary Report, Current Status, Next Plan
# Keywords: Laporan:, Status Pagi HH:MM:, Rencana:

def split_summary_report_z23(val):
    s = str(val).strip() if pd.notna(val) else ""
    if not s:
        return "", "", ""
    s = s.replace("_x000D_", "\n")  # Normalize Excel carriage return

    # Extract Laporan (Summary Report) - strip leading "-"
    summary = ""
    m1 = re.search(r"Laporan:\s*(.*?)(?=Status Pagi|Rencana:|$)", s, re.DOTALL)
    if m1:
        summary = m1.group(1).strip()
        if summary.startswith("-"):
            summary = summary[1:].strip()

    # Extract Status Pagi (Current Status) - flexible HH:MM or H:MM
    status = ""
    m2 = re.search(r"Status Pagi\s*\d{1,2}:\d{2}\s*:\s*(.*?)(?=Rencana:|$)", s, re.DOTALL)
    if m2:
        status = m2.group(1).strip()

    # Extract Rencana (Next Plan)
    plan = ""
    m3 = re.search(r"Rencana:\s*(.*)$", s, re.DOTALL)
    if m3:
        plan = m3.group(1).strip()

    return summary, status, plan

split_result = df_z23["Summary Report"].apply(split_summary_report_z23)
df_z23["Summary Report"] = split_result.apply(lambda x: x[0])
df_z23["Current Status"] = split_result.apply(lambda x: x[1])
df_z23["Next Plan"] = split_result.apply(lambda x: x[2])
```

**How it works:**
- **Well Name:** Regex matches `Part1\n(Part2)\n(Part3)` or `Part1\n(Part2)`. Part 2 becomes `Well Name`, Part 3 becomes `Well Name [2]`; Part 1 is discarded. If Part 1 or Part 3 is blank, it falls back to Part 2.
- **Summary Report:** `_x000D_` (Excel newline) is replaced with `\n`. Sections are extracted by `Laporan:`, `Status Pagi HH:MM:`, and `Rencana:`. A leading `-` after `Laporan:` is stripped.
- Results are sorted by `Rig Name`.

**Key variables/parameters:**
- `_x000D_` — Excel carriage return; must be normalized for regex to work.
- Zone 2&3 uses `Current Status` (from `Status Pagi` section).

**Gotchas / Assumptions:**
- Well name format: `Line1\n(Part2)\n(Part3)` or `Line1\n(Part2)`.
- Summary format: `Laporan: ... Status Pagi HH:MM: ... Rencana: ...`.
- If regex doesn’t match, Part 1 becomes the full string and Part 2/3 stay blank.

---

### 7. Zone 4 Processing

**What it does:** Filters Zone 4 rows, splits `Well Name` as `Part1 (Part2)` → `Well Name` and `Well Name [2]`, removes zero-width space (`\u2060`), normalizes `Rig Name` (removes `Rig`, applies specific replacements), splits `Summary Report` on `Status Pagi` and `Plan:`, sorts by `Rig Name`.

**Code:**

```python
# Split Well Name: Part1 (Part2) -> Well Name, Well Name [2]
# If Well Name [2] is blank, copy from Well Name

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

df_z4 = df[df["Zone"] == "Zone 4"].copy()
df_z4["Well Name"] = df_z4["Well Name"].str.replace("\u2060", "", regex=False)

split_result = df_z4["Well Name"].apply(split_well_name_z4)
df_z4["Well Name"] = split_result.apply(lambda x: x[0])
df_z4["Well Name [2]"] = split_result.apply(lambda x: x[1])
df_z4.reset_index(drop=True, inplace=True)
```

```python
df_z4["Rig Name"] = df_z4["Rig Name"].str.replace("Rig", "", regex=False).str.strip()

# Replace known names, and normalize whitespace after '#' for 'PDSI' rigs
df_z4["Rig Name"] = (
    df_z4["Rig Name"]
    .replace({
        "Airlangga #55": "Airlangga-55",
        "PDSI ACS#21": "ACS-21",
        "#36.1/Skytop 650M": "PDSI #36.1/Skytop 650M"
    })
    .apply(lambda x: re.sub(r'(PDSI #)\s+', r'\1', x) if isinstance(x, str) and x.startswith('PDSI #') else x)
)

# Sort by column Rig Name
df_z4 = df_z4.sort_values(by="Rig Name")
```

```python
# Split Summary Report into 3 columns: Summary Report, Current Status, Next Plan
# Keywords: Status Pagi, Plan:

def split_summary_report_z4(val):
    s = str(val).strip() if pd.notna(val) else ""
    if not s:
        return "", "", ""
    s = s.replace("_x000D_", "\n")  # Normalize Excel carriage return

    # Summary Report: content from start until Status Pagi or Plan:
    summary = ""
    m1 = re.search(r"^(.+?)(?=Status Pagi|Plan:|$)", s, re.DOTALL)
    if m1:
        summary = m1.group(1).strip()

    # Current Status: after Status Pagi (optional time) : until Plan:
    status = ""
    m2 = re.search(r"Status Pagi(?:\s*\d{1,2}:\d{2})?\s*:\s*(.*?)(?=Plan:|$)", s, re.DOTALL)
    if m2:
        status = m2.group(1).strip()

    # Next Plan: after Plan: until end
    plan = ""
    m3 = re.search(r"Plan:\s*(.*)$", s, re.DOTALL)
    if m3:
        plan = m3.group(1).strip()

    return summary, status, plan

split_result = df_z4["Summary Report"].apply(split_summary_report_z4)
df_z4["Summary Report"] = split_result.apply(lambda x: x[0])
df_z4["Current Status"] = split_result.apply(lambda x: x[1])
df_z4["Next Plan"] = split_result.apply(lambda x: x[2])
```

**How it works:**
- **Well Name:** Regex `^(.+?)\s*\(([^)]*)\)\s*$` matches `Part1 (Part2)`. Part 1 → `Well Name`, Part 2 → `Well Name [2]`; if Part 2 is empty, it copies Part 1. `\u2060` (zero-width space) is removed.
- **Rig Name:** Strips `Rig`, applies mapping (e.g., `Airlangga #55` → `Airlangga-55`), and normalizes `PDSI #` spacing (removes extra space after `#`).
- **Summary Report:** Same `_x000D_` normalization. Extracts: content before `Status Pagi` or `Plan:` → summary; after `Status Pagi [HH:MM]:` until `Plan:` → current status; after `Plan:` → next plan.

**Key variables/parameters:**
- `\u2060` — zero-width space; removed for clean parsing.
- Rig name replacements are hardcoded for Zone 4-specific formatting.

**Gotchas / Assumptions:**
- Zone 4 well name format: `Part1 (Part2)`.
- Zone 4 uses `Plan:` (English) vs Zone 2&3’s `Rencana:` (Indonesian).
- `Status Pagi` time is optional in Zone 4.

---

### 8. Merge and Display

**What it does:** Concatenates Zone 1, Zone 2&3, and Zone 4 DataFrames, then adds a date picker with Previous/Next buttons. Changing the date filters the view and copies the filtered rows to the clipboard (without header, without index).

**Code:**

```python
# Merge Zone 1, Zone 2&3, and Zone 4 DataFrames
df_merged = pd.concat([df_z1, df_z23, df_z4], ignore_index=True)
```

```python
# Add date picker widget and display dataframe (for df_merged)
import ipywidgets as widgets
from IPython.display import display, clear_output
from datetime import timedelta

# Date picker and navigation buttons
date_picker_merged = widgets.DatePicker(description='Pick a date:', disabled=False)
btn_decrement_merged = widgets.Button(description='◀ Previous Day', button_style='info')
btn_increment_merged = widgets.Button(description='Next Day ▶', button_style='info')

# Navigation button callbacks
def decrement_date_merged(btn):
    if date_picker_merged.value is not None:
        date_picker_merged.value = date_picker_merged.value - timedelta(days=1)

def increment_date_merged(btn):
    if date_picker_merged.value is not None:
        date_picker_merged.value = date_picker_merged.value + timedelta(days=1)

btn_decrement_merged.on_click(decrement_date_merged)
btn_increment_merged.on_click(increment_date_merged)
widget_box_merged = widgets.HBox([btn_decrement_merged, date_picker_merged, btn_increment_merged])

# Filtering and display function
def show_df_by_date_merged(change):
    clear_output(wait=True)
    display(widget_box_merged)
    selected_date = date_picker_merged.value
    if selected_date is not None:
        filtered_df = df_merged[df_merged["Report Date"] == selected_date]
    else:
        filtered_df = df_merged
    display(filtered_df.head(3))
    filtered_df.to_clipboard(header=False, index=False)

# Initial display
clear_output(wait=True)
display(widget_box_merged)
display(df_merged.head(3))
date_picker_merged.observe(show_df_by_date_merged, names='value')
```

**How it works:**
- `pd.concat([df_z1, df_z23, df_z4], ignore_index=True)` merges in order: Zone 1, then Zone 2&3, then Zone 4.
- `DatePicker` holds the selected date; buttons add/subtract one day.
- `show_df_by_date_merged` runs when the date changes: clears output, shows the widget, filters `df_merged` by `Report Date`, displays first 3 rows, and copies the **full** filtered DataFrame to clipboard (`header=False`, `index=False`).
- `display(df_merged.head(3))` shows a preview; the clipboard holds the complete filtered table for pasting elsewhere.

**Key variables/parameters:**
- `filtered_df.head(3)` — only the first 3 rows are shown; clipboard gets all filtered rows.
- `to_clipboard(header=False, index=False)` — tab-separated, no header or index row.

**Gotchas / Assumptions:**
- Clipboard access requires proper permissions in the Jupyter environment.
- Date picker defaults to today; `Report Date` may not match if no data for that date.
- On date change, both display and clipboard are updated; user can paste directly into Excel or similar.

---

## Data Flow

```
Excel files (daily-report/{month_folder}/*.xlsx)
    → extract date from filename, read sheet by date (header=13)
    → concat into df
    → select columns, filter zones, rename
    → add reference columns, defaults, APH mapping
    → split by Zone
        → Zone 1: parse Well Name (/), Summary Report (Plan:), Rig Name
        → Zone 2&3: parse Well Name (parens), Summary Report (Laporan/Status Pagi/Rencana)
        → Zone 4: parse Well Name (parens), Rig Name (replacements), Summary Report (Status Pagi/Plan)
    → concat df_z1, df_z23, df_z4 → df_merged
    → display with date picker, copy filtered rows to clipboard
```

## Quick Reference

**How to run:**
1. Open `Region 1/app.ipynb` in Jupyter.
2. Ensure `daily-report/{month_folder}/` contains Excel files (e.g., `daily-report/02-2026/`).
3. Set `month_folder` in the first code cell (e.g., `"02-2026"`).
4. Run all cells (top to bottom).

**Environment setup:**
```bash
pip install pandas openpyxl ipywidgets
```

**Expected input:**
- Path: `Region 1/daily-report/{month_folder}/*.xlsx`
- Filename pattern: must contain `tanggal DD MMM YYYY` (e.g., `...tanggal 12 Feb 2026.xlsx`)
- Sheet names: `D MMM` or `DD MMM` (e.g., `1 Feb`, `01 Feb`)
- Header row: row 14 (0-indexed: 13)

**Output:**
- Merged DataFrame `df_merged` with 19 columns.
- Interactive date picker + Previous/Next buttons.
- Filtered rows copied to clipboard on date change (tab-separated, no header/index).

**Known limitations:**
- Fixed `header=13`; source layout must not change.
- Zone-specific parsing assumes specific text formats; non-standard formats may parse incorrectly.
- Clipboard may not work in headless or remote Jupyter environments.
- Only Zone 1, 2&3, and 4 are processed; other zones are excluded.
