import pandas as pd
import os
from datetime import datetime

# === INPUT FILE PATHS (update as needed) ===
data_folder = "./input"
ref_racf_file = os.path.join(data_folder, "Reference Table - RACF with SF.xlsx")
ref_comm_file = os.path.join(data_folder, "Reference Table - COMM.xlsx")
raw_data_file = os.path.join(data_folder, "PlenaBillcodesDIT.xlsx")

# === LOAD DATA ===
raw_df = pd.read_excel(raw_data_file, sheet_name="BillCodeRates", dtype=str)
raw_df['Effective Date*'] = pd.to_datetime(raw_df['Effective Date*'], errors='coerce')
raw_df = raw_df.dropna(subset=['Effective Date*'])

# === CLEAN FUNDER TYPE AND CATEGORISE ===
def classify_row(row):
    funder = str(row['FunderCode*']).lower()
    bill = str(row['BillCode*']).lower()
    if 'racf' in funder and 'racfpff' not in funder:
        return 'RACF'
    elif 'comm' in funder:
        if '(hrly)' in bill or '(hourly)' in bill:
            return 'COMM'
        else:
            return 'OTHERS'
    else:
        return 'OTHERS'

raw_df['CATEGORY'] = raw_df.apply(classify_row, axis=1)

# === KEEP LATEST EFFECTIVE DATE PER BILL CODE IN EACH CATEGORY ===
def get_latest_by_billcode(df):
    return df.sort_values('Effective Date*').groupby('BillCode*', as_index=False).last()

racf_df = get_latest_by_billcode(raw_df[raw_df['CATEGORY'] == 'RACF'])
comm_df = get_latest_by_billcode(raw_df[raw_df['CATEGORY'] == 'COMM'])
other_df = get_latest_by_billcode(raw_df[raw_df['CATEGORY'] == 'OTHERS'])

# === MATCH RACF WITH SF REF TABLE ===
ref_racf = pd.read_excel(ref_racf_file, sheet_name=0)
ref_racf.columns = ref_racf.columns.str.strip()

racf_df['Matched Salesforce Code'] = racf_df['FunderCode*'].str.strip().map(
    lambda x: next((ref_racf.iloc[i, 0] for i in range(len(ref_racf))
                    if str(x).strip() == str(ref_racf.iloc[i, 1]).strip()), None))

racf_df['New CPI Adjusted Rate'] = racf_df.apply(lambda row:
    round(float(row['Rate*']) * float(ref_racf[ref_racf.iloc[:, 1].astype(str).str.strip() == str(row['FunderCode*']).strip()].iloc[0, 2]), 2)
    if pd.notnull(row['Matched Salesforce Code']) and str(row['Rate*']).replace('.', '', 1).isdigit()
    else None,
    axis=1
)

unmatched_racf = racf_df[racf_df['Matched Salesforce Code'].isna()]
racf_df = racf_df.dropna(subset=['Matched Salesforce Code'])

# === MATCH COMM WITH REF TABLE ===
ref_comm = pd.read_excel(ref_comm_file, sheet_name=0)
ref_comm.columns = ref_comm.columns.str.strip()

comm_df['Matched Rate'] = comm_df['BillCode*'].str.strip().map(
    lambda x: next((ref_comm.iloc[i, ref_comm.columns.get_loc("Rate*")] for i in range(len(ref_comm))
                    if str(x).strip().lower() == str(ref_comm.iloc[i, 0]).strip().lower()), "No Match"))

unmatched_comm = comm_df[comm_df['Matched Rate'] == "No Match"]
comm_df = comm_df[comm_df['Matched Rate'] != "No Match"]

# === COMBINE UNMATCHED TO OTHERS ===
final_others = pd.concat([other_df, unmatched_racf, unmatched_comm], ignore_index=True)

# === EXPORT TO EXCEL ===
output_path = "Working_File_Cleaned.xlsx"
with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
    racf_df.to_excel(writer, index=False, sheet_name="RACF")
    comm_df.to_excel(writer, index=False, sheet_name="COMM")
    final_others.to_excel(writer, index=False, sheet_name="OTHERS")
    ref_racf.to_excel(writer, index=False, sheet_name="Reference Table - RACF with SF")
    ref_comm.to_excel(writer, index=False, sheet_name="Reference Table - COMM")

print(f"✅ Processing complete. File saved to: {output_path}")
