import streamlit as st
import pandas as pd
import difflib
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# --- Streamlit Config ---
st.set_page_config(layout="wide", page_title="SAP vs PLM BOM Validation")

st.title("📊 SAP vs PLM BOM Validation Tool")
st.write("Upload SAP and PLM Excel files to compare material numbers and consumption quantities.")

# --- File Uploads ---
sap_file = st.file_uploader("Upload SAP File", type=["xlsx", "xls"])
plm_file = st.file_uploader("Upload PLM File", type=["xlsx", "xls"])

if sap_file and plm_file:
    try:
        # --- Read sheets ---
        sap = pd.read_excel(sap_file, sheet_name="SAP")
        plm = pd.read_excel(plm_file, sheet_name="PLM")

        # --- Add suffixes for clarity ---
        sap = sap.add_suffix("_SAP")
        plm = plm.add_suffix("_PLM")

        # --- Step 1: Direct Material Match ---
        direct_matches = pd.merge(
            plm, sap, left_on="Material_PLM", right_on="Material_SAP", how="inner"
        )

        missing_in_sap = plm[~plm["Material_PLM"].isin(sap["Material_SAP"])]
        missing_in_plm = sap[~sap["Material_SAP"].isin(plm["Material_PLM"])]

        # --- Add consumption comparison ---
        if not direct_matches.empty:
            direct_matches["ConsumptionDiff"] = direct_matches["Qty(Cons.)_PLM"] - direct_matches["Comp.Qty._SAP"]
            direct_matches["DifferenceFlag"] = direct_matches["ConsumptionDiff"].apply(
                lambda x: "SAP consumption is higher" if x < 0 else "OK"
            )
            direct_matches = direct_matches.reindex(
                direct_matches["ConsumptionDiff"].abs().sort_values(ascending=False).index
            )

        # Keep only requested columns for Direct Matches tab
        direct_cols = [
            "Material_PLM",
            "Material Description_SAP",
            "Vendor Reference_PLM",
            "Vendor Reference_SAP",
            "Color Reference_PLM",
            "Comp. Colour_SAP",
            "Qty(Cons.)_PLM",
            "Comp.Qty._SAP",
            "ConsumptionDiff",
            "DifferenceFlag"
        ]
        direct_matches_tab = direct_matches[[col for col in direct_cols if col in direct_matches.columns]]

        # --- Step 2: Build Combined Column for Fuzzy Matching ---
        plm["Combined_PLMMeta"] = (
            plm["Material_PLM"].astype(str).str.strip() + " " +
            plm["Vendor Reference_PLM"].astype(str).str.strip() + " " +
            plm["Color Reference_PLM"].astype(str).str.strip()
        )

        # --- Step 3: Fuzzy Matching (≥70%) ---
        fuzzy_matches = []
        for _, row in direct_matches.iterrows():
            combined_val = (
                str(row["Material_PLM"]).strip() + " " +
                str(row.get("Vendor Reference_PLM", "")).strip() + " " +
                str(row.get("Color Reference_PLM", "")).strip()
            )
            best_match = difflib.get_close_matches(
                combined_val, sap.get("Material Description_SAP", []), n=1, cutoff=0.7
            )
            if best_match:
                sap_row = sap[sap["Material Description_SAP"] == best_match[0]].iloc[0]
                fuzzy_matches.append({
                    "Material_PLM": row["Material_PLM"],
                    "Combined_PLMMeta": combined_val,
                    "MaterialDescription_SAP": best_match[0],
                    "Qty(Cons.)_PLM": row.get("Qty(Cons.)_PLM", 0),
                    "Comp.Qty._SAP": sap_row.get("Comp.Qty._SAP", 0),
                    "ConsumptionDiff": row.get("Qty(Cons.)_PLM", 0) - sap_row.get("Comp.Qty._SAP", 0),
                    "DifferenceFlag": "SAP consumption is higher" if sap_row.get("Comp.Qty._SAP", 0) > row.get("Qty(Cons.)_PLM", 0) else "OK",
                    "Vendor Reference_PLM": row.get("Vendor Reference_PLM", ""),
                    "Vendor Reference_SAP": sap_row.get("Vendor Reference_SAP", ""),
                    "Color Reference_PLM": row.get("Color Reference_PLM", ""),
                    "Comp. Colour_SAP": sap_row.get("Comp. Colour_SAP", "")
                })

        fuzzy_df = pd.DataFrame(fuzzy_matches)
        if not fuzzy_df.empty:
            fuzzy_df = fuzzy_df.reindex(
                fuzzy_df["ConsumptionDiff"].abs().sort_values(ascending=False).index
            )

        # --- Summary Counts ---
        matched_count = len(direct_matches_tab) + len(fuzzy_df) if not fuzzy_df.empty else len(direct_matches_tab)
        unmatched_sap_count = len(missing_in_sap)
        unmatched_plm_count = len(missing_in_plm)

        st.subheader("📈 Summary Counts")
        col1, col2, col3 = st.columns(3)
        col1.metric("✅ Total Matched Records", matched_count)
        col2.metric("❌ PLM Not in SAP", unmatched_sap_count)
        col3.metric("❌ SAP Not in PLM", unmatched_plm_count)

        # --- Save Results to Excel ---
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            direct_matches_tab.to_excel(writer, sheet_name="Direct_Matches", index=False)
            missing_in_sap.to_excel(writer, sheet_name="PLM_Not_in_SAP", index=False)
            missing_in_plm.to_excel(writer, sheet_name="SAP_Not_in_PLM", index=False)
            if not fuzzy_df.empty:
                fuzzy_df.to_excel(writer, sheet_name="70%_or_more_Matches", index=False)
        output.seek(0)

        # --- Conditional Formatting ---
        wb = load_workbook(output)

        def apply_coloring(ws, headers, diff_col_name, flag_col_name):
            if not all(col in headers for col in [diff_col_name, flag_col_name]):
                return
            diff_col = headers.index(diff_col_name) + 1
            flag_col = headers.index(flag_col_name) + 1

            red_fill = PatternFill(start_color="FF6666", end_color="FF6666", fill_type="solid")
            green_fill = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")

            for row in range(2, ws.max_row + 1):
                flag = ws.cell(row=row, column=flag_col).value
                fill = red_fill if flag == "SAP consumption is higher" else green_fill
                ws.cell(row=row, column=diff_col).fill = fill
                ws.cell(row=row, column=flag_col).fill = fill

        for sheet_name in ["Direct_Matches", "70%_or_more_Matches"]:
            if sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
                headers = [cell.value for cell in ws[1]]
                apply_coloring(ws, headers, "ConsumptionDiff", "DifferenceFlag")

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

        # --- Preview Tabs ---
        st.subheader("🔍 Preview of Results")
        tab1, tab2, tab3, tab4 = st.tabs(["Direct Matches", "PLM Not in SAP", "SAP Not in PLM", "70% or more Matches"])
        with tab1:
            st.dataframe(direct_matches_tab)
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
