import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Plena Billcode Processor", layout="wide")
st.title("📋 Plena Billcode Processor")

# Sidebar for uploads
st.sidebar.header("Upload Files")
raw_file = st.sidebar.file_uploader("PlenaBillcodesDIT.xlsx", type=["xlsx"])
racf_ref_file = st.sidebar.file_uploader("Reference Table - RACF with SF.xlsx", type=["xlsx"])
comm_ref_file = st.sidebar.file_uploader("Reference Table - COMM.xlsx", type=["xlsx"])

# Utility
def clean_trim(text):
    return str(text).strip().replace("\n", "").replace("\r", "")

# App logic
if raw_file:
    try:
        raw_df = pd.read_excel(raw_file, sheet_name="BillCodeRates", dtype=str)
        raw_df.columns = [str(c).strip() for c in raw_df.columns]
        st.success("✅ Raw file loaded successfully.")
        st.dataframe(raw_df.head())

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

        df_racf = pd.DataFrame(racf_dict.values())
        df_comm = pd.DataFrame(comm_dict.values())
        df_others = pd.DataFrame(others_dict.values())

        # Apply CPI rate from reference tables
        if comm_ref_file:
            ref_comm = pd.read_excel(comm_ref_file, dtype=str)
            ref_comm.columns = [str(c).strip().lower() for c in ref_comm.columns]

            def match_and_apply_comm_rate(row):
                code = clean_trim(row["BillCode*"]).lower()
                match = ref_comm[ref_comm.iloc[:, 0].str.strip().str.lower() == code]
                if not match.empty and pd.notnull(row["Rate*"]):
                    try:
                        return round(float(row["Rate*"]) * float(match.iloc[0, 1]), 2)
                    except:
                        return row["Rate*"]
                return row["Rate*"]

            df_comm["Rate*"] = df_comm.apply(match_and_apply_comm_rate, axis=1)

        if racf_ref_file:
            ref_racf = pd.read_excel(racf_ref_file, dtype=str)
            ref_racf.columns = [str(c).strip().lower() for c in ref_racf.columns]

            def match_and_apply_racf_rate(row):
                code = clean_trim(row["BillCode*"]).lower()
                match = ref_racf[ref_racf["alayacare funder code"].str.strip().str.lower() == code]
                if not match.empty and pd.notnull(row["Rate*"]):
                    try:
                        return round(float(row["Rate*"]) * float(match["cpi rate"].values[0]), 2)
                    except:
                        return row["Rate*"]
                return row["Rate*"]

            df_racf["Rate*"] = df_racf.apply(match_and_apply_racf_rate, axis=1)

        # Combine all into single sheet
        combined_df = pd.concat([df_racf, df_comm, df_others], ignore_index=True)

        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            combined_df.to_excel(writer, index=False, sheet_name="All Rates")
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
