
import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Plena Billcode Processor", layout="wide")
st.title("📋 Plena Billcode Processor")

# Step 1: Upload required files
st.sidebar.header("Upload Files")
raw_file = st.sidebar.file_uploader("PlenaBillcodesDIT.xlsx", type=["xlsx"])
racf_ref_file = st.sidebar.file_uploader("Reference Table - RACF with SF.xlsx", type=["xlsx"])
comm_ref_file = st.sidebar.file_uploader("Reference Table - COMM.xlsx", type=["xlsx"])

# Utility to clean and trim text
def clean_trim(text):
    return str(text).strip().replace("\n", "").replace("\r", "")

# Function to get column index safely
def get_col_index(header_row, search_terms):
    for idx, col in enumerate(header_row):
        col_clean = col.strip().lower()
        for term in search_terms:
            if term in col_clean:
                return idx
    return -1

# Step 2: Process the uploaded raw file
if raw_file:
    try:
        raw_df = pd.read_excel(raw_file, sheet_name="BillCodeRates", dtype=str)
        raw_df.columns = [str(c).strip() for c in raw_df.columns]
        st.success("✅ Raw file loaded successfully.")
        st.dataframe(raw_df.head())

        # Prepare dictionaries
        racf_dict, comm_dict, others_dict = {}, {}, {}
        skipped = 0

        for _, row in raw_df.iterrows():
            billcode = str(row.get("BillCode*", "")).strip()
            fundercode = str(row.get("FunderCode*", "")).lower().strip()
            eff_date = row.get("Effective Date*", None)

            if pd.isna(eff_date):
                skipped += 1
                continue

            key = billcode
            target = others_dict

            if "racf" in fundercode and "racfpff" not in fundercode:
                target = racf_dict
            elif "comm" in fundercode:
                if "(hrly)" in billcode.lower() or "(hourly)" in billcode.lower():
                    target = comm_dict
                else:
                    target = others_dict

            if key not in target or pd.to_datetime(eff_date) > pd.to_datetime(target[key]["Effective Date*"]):
                target[key] = row

        st.write(f"✅ RACF entries: {len(racf_dict)}")
        st.write(f"✅ COMM entries: {len(comm_dict)}")
        st.write(f"✅ OTHERS entries: {len(others_dict)}")
        st.write(f"⚠️ Skipped rows with invalid date: {skipped}")

        # Convert dictionaries to dataframes
        df_racf = pd.DataFrame(racf_dict.values())
        df_comm = pd.DataFrame(comm_dict.values())
        df_others = pd.DataFrame(others_dict.values())

        # Step 3: Match RACF and COMM against Reference Tables if uploaded
        if racf_ref_file:
            ref_racf = pd.read_excel(racf_ref_file, dtype=str)
            ref_racf.columns = [str(c).strip() for c in ref_racf.columns]
            df_racf["Matched Salesforce Code"] = df_racf["FunderCode*"].apply(
                lambda x: ref_racf.loc[ref_racf.iloc[:, 1].str.strip() == clean_trim(x), ref_racf.columns[0]].values[0]
                if any(ref_racf.iloc[:, 1].str.strip() == clean_trim(x)) else "Not Found"
            )

        if comm_ref_file:
            ref_comm = pd.read_excel(comm_ref_file, dtype=str)
            ref_comm.columns = [str(c).strip() for c in ref_comm.columns]
            df_comm["Matched Rate"] = df_comm["BillCode*"].apply(
                lambda x: ref_comm.loc[ref_comm.iloc[:, 0].str.strip() == clean_trim(x), ref_comm.columns[1]].values[0]
                if any(ref_comm.iloc[:, 0].str.strip() == clean_trim(x)) else "No Match"
            )

        # Step 4: Export combined file
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            df_racf.to_excel(writer, index=False, sheet_name="RACF")
            df_comm.to_excel(writer, index=False, sheet_name="COMM")
            df_others.to_excel(writer, index=False, sheet_name="OTHERS")
            if racf_ref_file:
                ref_racf.to_excel(writer, index=False, sheet_name="Reference Table - RACF with SF")
            if comm_ref_file:
                ref_comm.to_excel(writer, index=False, sheet_name="Reference Table - COMM")
        output.seek(0)

        st.download_button(
            label="📥 Download Processed Excel",
            data=output,
            file_name="Processed_Billcodes.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ Error processing file: {e}")
