import streamlit as st
import pandas as pd
import io
from datetime import datetime

st.title("🧾 Plena Bill Code Processor (CSV Version)")
st.write("Upload 3 CSV files (exported from Excel) to process and generate a cleaned output.")

uploaded_raw = st.file_uploader("1️⃣ PlenaBillcodesDIT.csv", type=["csv"])
uploaded_racf_ref = st.file_uploader("2️⃣ Reference Table - RACF with SF.csv", type=["csv"])
uploaded_comm_ref = st.file_uploader("3️⃣ Reference Table - COMM.csv", type=["csv"])

def clean_trim(text):
    return str(text).strip()

def deduplicate_by_latest(df, key_col, date_col):
    df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
    df = df.dropna(subset=[date_col])
    return df.sort_values(by=date_col).drop_duplicates(subset=[key_col], keep='last')

if uploaded_raw and uploaded_racf_ref and uploaded_comm_ref:
    try:
        raw_df = pd.read_csv(uploaded_raw)
        racf_ref = pd.read_csv(uploaded_racf_ref, header=None)
        comm_ref = pd.read_csv(uploaded_comm_ref, header=None)

        raw_df.columns = [col.lower().strip() for col in raw_df.columns]
        funder_col = next(col for col in raw_df.columns if "fundercode" in col)
        billcode_col = next(col for col in raw_df.columns if "billcode" in col)
        effdate_col = next(col for col in raw_df.columns if "effective date" in col)
        rate_col = next((col for col in raw_df.columns if "rate" in col), None)

        racf_df, comm_df, others_df = [], [], []
        for _, row in raw_df.iterrows():
            funder = str(row[funder_col]).lower()
            billcode = str(row[billcode_col]).lower()
            if pd.isnull(row[effdate_col]):
                continue
            if "racf" in funder and "racfpff" not in funder:
                racf_df.append(row)
            elif "comm" in funder:
                if "(hrly)" in billcode or "(hourly)" in billcode:
                    comm_df.append(row)
                else:
                    others_df.append(row)
            else:
                others_df.append(row)

        racf_df = deduplicate_by_latest(pd.DataFrame(racf_df), billcode_col, effdate_col)
        comm_df = deduplicate_by_latest(pd.DataFrame(comm_df), billcode_col, effdate_col)
        others_df = deduplicate_by_latest(pd.DataFrame(others_df), billcode_col, effdate_col)

        racf_df["Matched Salesforce Code"] = racf_df[funder_col].map(
            lambda x: racf_ref.set_index(1).to_dict()[0].get(clean_trim(x), "Not Found")
        )

        if rate_col:
            racf_df["New CPI Adjusted Rate"] = racf_df.apply(
                lambda r: round(float(r[rate_col]) * float(
                    racf_ref[racf_ref[1].apply(clean_trim) == clean_trim(r[funder_col])][2].values[0]
                ), 2) if clean_trim(r[funder_col]) in racf_ref[1].apply(clean_trim).values else None, axis=1
            )

        comm_df["Matched Rate"] = comm_df[billcode_col].map(
            lambda x: comm_ref.set_index(0).to_dict()[1].get(clean_trim(x), "No Match")
        )

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            racf_df.to_excel(writer, index=False, sheet_name="RACF")
            comm_df.to_excel(writer, index=False, sheet_name="COMM")
            others_df.to_excel(writer, index=False, sheet_name="Others")
            racf_ref.to_excel(writer, index=False, sheet_name="Reference Table - RACF with SF")
            comm_ref.to_excel(writer, index=False, sheet_name="Reference Table - COMM")

        st.success("✅ Processing complete!")

        st.download_button(
            label="📥 Download Excel Output",
            data=output.getvalue(),
            file_name=f"Plena_Output_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ Error during processing: {e}")
