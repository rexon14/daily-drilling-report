import sys
import os
import re
import pandas as pd
import numpy as np

def main(date):
    def parse_file(date_str, input_dir="DDR", output_dir="Output_raw"):
        input_path = os.path.join(input_dir, f"{date_str}.txt")
        output_path = os.path.join(output_dir, f"{date_str}.xlsx")
        os.makedirs(output_dir, exist_ok=True)

        rows = []
        current_field = None
        current_item = None
        summary_flag = status_flag = plan_flag = False

        with open(input_path, encoding='utf-8') as f:
            for raw in f:
                line = raw.strip().replace("*", "")
                if not line:
                    continue

                low = line.lower()

                # FIELD header
                if low.startswith('field'):
                    parts = line.split(None, 1)
                    current_field = parts[1].strip() if len(parts) > 1 else None
                    continue

                # New record: "1. WELL-NAME (OPTIONAL-PAREN)"
                m = re.match(r'^(\d+)\.\s*(.+)$', line)
                if m:
                    if current_item:
                        rows.append(current_item)
                    idx = int(m.group(1))
                    desc = m.group(2).strip()
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
                        'Realiasi Biaya': None,
                        'Summary Report': None,
                        'Curent Status': None,
                        'Plan': None
                    }
                    summary_flag = status_flag = plan_flag = False
                    continue

                if current_item is None:
                    continue

                # Section headers
                if 'summary report' in low:
                    summary_flag, status_flag, plan_flag = True, False, False
                    continue
                if 'current status' in low:
                    status_flag, summary_flag, plan_flag = True, False, False
                    continue
                if low.startswith('next plan') or low.startswith('plan'):
                    plan_flag, summary_flag, status_flag = True, False, False
                    continue

                # Capture first line of Summary/Status/Plan (strip "=" too)
                if summary_flag:
                    val = line.lstrip('-').strip().rstrip('.')
                    val = val.strip('= ').strip()
                    current_item['Summary Report'] = val
                    summary_flag = False
                    continue
                if status_flag:
                    val = line.lstrip('-').strip().rstrip('.')
                    val = val.strip('= ').strip()
                    current_item['Curent Status'] = val
                    status_flag = False
                    continue
                if plan_flag:
                    val = line.lstrip('-').strip().rstrip('.')
                    val = val.strip('= ').strip()
                    current_item['Plan'] = val
                    plan_flag = False
                    continue

                # Key: Value lines
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

                # WOL Hari ke
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
                    current_item['AFE Cost'] = float(num) if num else None
                    continue

                # Realiasi Biaya
                if 'realisasi' in kl:
                    before_paren = val.split('(')[0].strip()
                    num = re.sub(r'[^\d.]', '', before_paren)
                    current_item['Realiasi Biaya'] = float(num) if num else None
                    continue

        # append last item
        if current_item:
            rows.append(current_item)

        # Build DataFrame
        df = pd.DataFrame(rows, columns=[
            'No','Field','Nama Sumur','Nama Sumur_2','Nama Rig',
            'WOL Hari ke','Hari ke','AFE Cost','Realiasi Biaya',
            'Summary Report','Curent Status','Plan'
        ])

        # Write to Excel (no styling) using XlsxWriter
        with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False)

    parse_file(date)

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


    cols = ['Field', 'Nama Sumur', 'Nama Sumur_2', 'Nama Rig', 'Summary Report', 'Curent Status', 'Plan']

    df = pd.read_excel(f"Output_raw\\{date}.xlsx")
    df = df[cols]

    # Add Column
    df = df.assign(**{"Region":"Region 3",
                    "Zone": "Zone 9",
                    'APH':"PHSS",
                    "Well Type": "Development",
                    "Location": "Onshore",
                    "Spud Date": np.nan,
                    "Release Date": np.nan,
                    "Status": np.nan,
                    "Status Code [1]": np.nan,
                    "Status Code [2]": np.nan
                    })

    # Rename Column
    df.rename(columns={
        "Nama Sumur": "Well Name",
        "Nama Sumur_2": "Well Name [2]",
        "Nama Rig": "Rig Name",
        "Curent Status": "Current Status",
        'Plan': "Next Plan"
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