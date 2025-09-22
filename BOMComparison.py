import streamlit as st
import pandas as pd
from thefuzz import fuzz, process
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

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
            plm, sap, left_on="Material", right_on="Material", how="inner", suffixes=("_PLM", "_SAP")
        )
        missing_in_sap = plm[~plm["Material"].isin(sap["Material"])]
        missing_in_plm = sap[~sap["Material"].isin(plm["Material"])]

        # --- Step 2: Fuzzy Matching ---
        plm["Combined"] = (
            plm["Material"].astype(str).str.strip() + " " +
            plm["Vendor Reference"].astype(str).str.strip() + " " +
            plm["Color Reference"].astype(str).str.strip() + " " +
            plm["Color Name"].astype(str).str.strip()
        )

        fuzzy_matches = []
        for idx, row in plm.iterrows():
            best_match, score = process.extractOne(
                row["Combined"], sap["Material Description"], scorer=fuzz.token_sort_ratio
            )
            if score >= 70:
                sap_row = sap[sap["Material Description"] == best_match].iloc[0]
                fuzzy_matches.append({
                    "Combined_PLMMeta": row["Combined"],
                    "MaterialDescription_SAP": best_match,
                    "Similarity": score,
                    "Qty(Cons.)": row.get("Qty(Cons.)", None),
                    "Comp.Qty.": sap_row.get("Comp.Qty.", None),
                    "ConsumptionDiff": (row.get("Qty(Cons.)", 0) - sap_row.get("Comp.Qty.", 0))
                })

        fuzzy_df = pd.DataFrame(fuzzy_matches)
        if not fuzzy_df.empty:
            fuzzy_df = fuzzy_df.sort_values(by="ConsumptionDiff", key=abs, ascending=False)

        # --- Save Results to Excel ---
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            direct_matches.to_excel(writer, sheet_name="Direct_Matches", index=False)
            missing_in_sap.to_excel(writer, sheet_name="PLM_Not_in_SAP", index=False)
            missing_in_plm.to_excel(writer, sheet_name="SAP_Not_in_PLM", index=False)
            fuzzy_df.to_excel(writer, sheet_name="Fuzzy_Matches", index=False)

        # --- Apply Formatting ---
        output.seek(0)
        wb = load_workbook(output)
        
        # Highlight unmatched PLM materials
        if "PLM_Not_in_SAP" in wb.sheetnames:
            ws = wb["PLM_Not_in_SAP"]
            red_fill = PatternFill(start_color="FF9999", end_color="FF9999", fill_type="solid")
            for row in ws.iter_rows(min_row=2, max_col=1, max_row=ws.max_row):
                for cell in row:
                    cell.fill = red_fill

        # Highlight fuzzy matches consumption
        if "Fuzzy_Matches" in wb.sheetnames:
            ws = wb["Fuzzy_Matches"]
            headers = [cell.value for cell in ws[1]]
            plm_col = headers.index("Qty(Cons.)") + 1
            sap_col = headers.index("Comp.Qty.") + 1
            diff_col = headers.index("ConsumptionDiff") + 1

            green_fill = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")
            red_fill = PatternFill(start_color="FF6666", end_color="FF6666", fill_type="solid")

            for row in range(2, ws.max_row + 1):
                plm_val = ws.cell(row=row, column=plm_col).value
                sap_val = ws.cell(row=row, column=sap_col).value
                diff_val = ws.cell(row=row, column=diff_col).value

                if plm_val == sap_val:
                    ws.cell(row=row, column=plm_col).fill = green_fill
                    ws.cell(row=row, column=sap_col).fill = green_fill
                else:
                    ws.cell(row=row, column=plm_col).fill = red_fill
                    ws.cell(row=row, column=sap_col).fill = red_fill
                    ws.cell(row=row, column=diff_col).fill = red_fill

        # Save workbook into memory
        final_output = BytesIO()
        wb.save(final_output)
        final_output.seek(0)

        # --- Download Button ---
        st.success("✅ Comparison complete! Download the results below.")
        st.download_button(
            label="📥 Download Comparison Report",
            data=final_output,
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
