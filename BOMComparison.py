import streamlit as st
import pandas as pd
import difflib
from io import BytesIO
from openpyxl import load_workbook

# --- Streamlit Config ---
st.set_page_config(layout="wide", page_title="SAP & PLM BOM Validation")

st.title("📊 SAP vs PLM BOM Validation Tool")
st.write("Upload SAP and PLM Excel files to compare material numbers and consumption quantities.")

# --- File Uploads ---
sap_file = st.file_uploader("Upload SAP File", type=["xlsx", "xls"])
plm_file = st.file_uploader("Upload PLM File", type=["xlsx", "xls"])

if sap_file and plm_file:
    try:
        sap = pd.read_excel(sap_file, sheet_name="SAP")
        plm = pd.read_excel(plm_file, sheet_name="PLM")

        # --- Step 1: Direct Material Match ---
        direct_matches = pd.merge(
            plm, sap, on="Material", how="inner", suffixes=("_PLM", "_SAP")
        )
        missing_in_sap = plm[~plm["Material"].isin(sap["Material"])]
        missing_in_plm = sap[~sap["Material"].isin(plm["Material"])]

        # --- Step 2: Fuzzy Matching (difflib) ---
        plm["Combined"] = (
            plm["Material"].astype(str).str.strip() + " " +
            plm["Vendor Reference"].astype(str).str.strip() + " " +
            plm["Color Reference"].astype(str).str.strip() + " " +
            plm["Color Name"].astype(str).str.strip()
        )

        fuzzy_matches = []
        for _, row in plm.iterrows():
            best_match = difflib.get_close_matches(
                row["Combined"], sap["Material Description"], n=1, cutoff=0.7
            )
            if best_match:
                sap_row = sap[sap["Material Description"] == best_match[0]].iloc[0]
                fuzzy_matches.append({
                    "Combined_PLMMeta": row["Combined"],
                    "MaterialDescription_SAP": best_match[0],
                    "Qty(Cons.)": row.get("Qty(Cons.)", 0),
                    "Comp.Qty.": sap_row.get("Comp.Qty.", 0),
                    "ConsumptionDiff": row.get("Qty(Cons.)", 0) - sap_row.get("Comp.Qty.", 0)
                })

        fuzzy_df = pd.DataFrame(fuzzy_matches)

        # --- Save Results to Excel ---
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            direct_matches.to_excel(writer, sheet_name="Direct_Matches", index=False)
            missing_in_sap.to_excel(writer, sheet_name="PLM_Not_in_SAP", index=False)
            missing_in_plm.to_excel(writer, sheet_name="SAP_Not_in_PLM", index=False)
            fuzzy_df.to_excel(writer, sheet_name="Fuzzy_Matches", index=False)

        output.seek(0)

        # --- Download Button ---
        st.success("✅ Comparison complete! Download the results below.")
        st.download_button(
            label="📥 Download Comparison Report",
            data=output,
            file_name="comparison_results.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # --- Preview ---
        st.subheader("🔍 Preview of Results")
        tab1, tab2, tab3, tab4 = st.tabs(["Direct Matches", "PLM Not in SAP", "SAP Not in PLM", "Fuzzy Matches"])
        with tab1:
            st.dataframe(direct_matches)
        with tab2:
            st.dataframe(missing_in_sap)
        with tab3:
            st.dataframe(missing_in_plm)
        with tab4:
            st.dataframe(fuzzy_df)

    except Exception as e:
        st.error(f"❌ Error while processing: {e}")
else:
    st.info("⬆️ Please upload both SAP and PLM Excel files to begin.")
