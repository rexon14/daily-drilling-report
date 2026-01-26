#!/usr/bin/env python3
import re
import pandas as pd
import numpy as np
import sys

def main(base):
    def parse_txt_file(path):
        with open(path, encoding='utf-8') as f:
            lines = [l.rstrip("\n").replace("*", "").replace("_", "") for l in f]
        records = []
        current = {}
        i = 0
        while i < len(lines):
            line = lines[i].strip()
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
                    "Realiasi Biaya": None,
                    "EMD": None,
                    "Summary Report": None,
                    "Cuurent Status": None,
                    "Next Plan": None
                }
                i += 1
                continue

            if not current:
                i += 1
                continue

            if line.startswith("Nama Sumur"):
                current["Nama Sumur"] = line.split(":",1)[1].strip()
            elif line.startswith("Nama Rig"):
                current["Nama Rig"] = line.split(":",1)[1].strip()
            elif line.startswith("Hari ke"):
                current["Hari ke"] = int(line.split(":",1)[1].strip())
            elif line.startswith("Kedalaman"):
                num = re.search(r"([\d\.]+)", line)
                val = num.group(1)
                current["Kedalaman (mMD)"] = float(val) if "." in val else int(val)
            elif line.startswith("Progres"):
                num = re.search(r"([\d\.]+)", line)
                val = num.group(1)
                current["Progress (mMD)"] = float(val) if "." in val else int(val)
            elif line.startswith("AFE"):
                num = re.sub(r"[^\d\.]", "", line.split(":",1)[1])
                current["AFE"] = int(float(num))
            elif line.lower().startswith("realisasi biaya"):
                m2 = re.search(r"([\d,\.]+)", line)
                num = m2.group(1).replace(",", "")
                current["Realiasi Biaya"] = int(float(num))
            elif line.startswith("EMD"):
                val = line.split(":",1)[1].strip()
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
                    current["Cuurent Status"] = lines[j].strip()
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
            "Kedalaman (mMD)", "Progress (mMD)", "AFE", "Realiasi Biaya",
            "EMD", "Summary Report", "Cuurent Status", "Next Plan"
        ]
        return pd.DataFrame(records, columns=cols)
    
    in_path = f"DDR\\{base}.txt"
    out_path = f"Output_raw\\{base}.xlsx"

    df = parse_txt_file(in_path)
    df.to_excel(out_path, index=False)

    # --- FINAL ---
    def clean():
        try:
            df['Summary Report'] = df['Summary Report'].str.replace(r'^[\-\:=]+\s*', '', regex=True)
        except:
            print("Error")

        try:
            df['Current Status'] = df['Current Status'].str.replace(r'^[\-\:=]+\s*', '', regex=True)
        except:
            print("Error")

        try:
            df['Next Plan'] = df['Next Plan'].str.replace(r'^[\-\:=]+\s*', '', regex=True)
        except:
            print("Error")


    def add_date():
        # convert the string to a Timestamp
        report_ts = pd.to_datetime(date)

        # add both columns
        df['Report Date']    = report_ts
        df['Operation Date'] = report_ts - pd.Timedelta(days=1)

        # strip off the time component
        df['Report Date']    = df['Report Date'].dt.date
        df['Operation Date'] = df['Operation Date'].dt.date

    # Reorder Column
    def reorder(df):
        df['Rig Name'] = df['Rig Name'].astype(str)
        cols = ["Region", "Zone", "APH", "Rig Name", "Well Name", "Well Name [2]", "Well Type", "Location", 
                "Spud Date", "Release Date", "Status", 
                "Status Code [1]", "Status Code [2]", "Summary Report", "Current Status", "Next Plan",
            "Report Date", "Operation Date"]
        return(df[cols])


    date = base
    cols = ['Nama Sumur', 'Nama Rig', 'Summary Report','Cuurent Status', 'Next Plan' ]

    df = pd.read_excel(f"Output_raw\\{date}.xlsx")
    df = df[cols]

    # Add Column
    df = df.assign(**{"Region":"Region 3",
                    "Zone": "Zone 8",
                    "APH":"PHM",
                    "Well Name [2]": np.nan,
                    "Well Type": "Development",
                    "Location": "Offshore",
                    "Spud Date": np.nan,
                    "Release Date": np.nan,
                    "Status": np.nan,
                    "Status Code [1]": np.nan,
                    "Status Code [2]": np.nan
                    })

    # Rename Column
    df.rename(columns={
        "Nama Sumur": "Well Name",
        "Nama Rig": "Rig Name",
        "Cuurent Status": "Current Status",
    }, inplace = True)

    # Fill [Well Name]
    df['Well Name [2]'] = df['Well Name [2]'].fillna(df['Well Name'])

    # Clean/Add Column
    clean()
    add_date()

    df = reorder(df)

    # Copies all values (no header, no index) to your clipboard
    df.sort_values(
        by='Rig Name',
        key=lambda s: s.str.lower(),      # lowercase for comparison
        ascending=True, inplace = True
    )

    df.reset_index(drop=True).to_clipboard(index=False, header=False)

    df.to_excel(f"Output\\{date}.xlsx", index = False)

# with guard:
if __name__=="__main__":
    if len(sys.argv)!=2:
        print("Usage: python main.py <argument>")
        sys.exit(1)

    main(sys.argv[1])