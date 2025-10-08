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
        sap = pd.read_excel(sap_file, sheet_name="SAP")
        plm = pd.read_excel(plm_file, sheet_name="PLM")

        # --- Step 1: Direct Material Match ---
        direct_matches = pd.merge(
            plm, sap, on="Material", how="inner", suffixes=("_PLM", "_SAP")
        )
        missing_in_sap = plm[~plm["Material"].isin(sap["Material"])]
        missing_in_plm = sap[~sap["Material"].isin(plm["Material"])]

        # --- Add consumption comparison for direct matches ---
        if not direct_matches.empty:
            direct_matches["ConsumptionDiff"] = direct_matches["Qty(Cons.)"] - direct_matches["Comp.Qty."]
            direct_matches["DifferenceFlag"] = direct_matches["ConsumptionDiff"].apply(lambda x: "RED" if x < 0 else "GREEN")
            direct_matches = direct_matches.reindex(
                direct_matches["ConsumptionDiff"].abs().sort_values(ascending=False).index
            )

        # --- Step 2: Build Combined Column for PLM ---
        plm["Combined"] = (
            plm["Material"].astype(str).str.strip() + " " +
            plm["Vendor Reference"].astype(str).str.strip() + " " +
            plm["Color Reference"].astype(str).str.strip() + " " +
            plm["Color Name"].astype(str).str.strip()
        )

        # --- Step 3: Fuzzy Matching (≥70%) ---
        fuzzy_matches = []
        for _, row in direct_matches.iterrows():
            combined_val = (
                str(row["Material"]).strip() + " " +
                str(row.get("Vendor Reference_PLM", "")).strip() + " " +
                str(row.get("Color Reference", "")).strip() + " " +
                str(row.get("Color Name", "")).strip()
            )
            best_match = difflib.get_close_matches(
                combined_val, sap["Material Description"], n=1, cutoff=0.7
            )
            if best_match:
                sap_row = sap[sap["Material Description"] == best_match[0]].iloc[0]
                fuzzy_matches.append({
                    "Material": row["Material"],
                    "Combined_PLMMeta": combined_val,
                    "MaterialDescription_SAP": best_match[0],
                    "Qty(Cons.)_PLM": row.get("Qty(Cons.)", 0),
                    "Comp.Qty._SAP": sap_row.get("Comp.Qty.", 0),
                    "ConsumptionDiff": row.get("Qty(Cons.)", 0) - sap_row.get("Comp.Qty.", 0),
                    "DifferenceFlag": "RED" if sap_row.get("Comp.Qty.", 0) > row.get("Qty(Cons.)", 0) else "GREEN",
                    "Vendor Reference_PLM": row.get("Vendor Reference_PLM", ""),
                    "Vendor Reference_SAP": sap_row.get("Vendor Reference", ""),
                    "Color Reference_PLM": row.get("Color Reference", ""),
                    "Comp. Colour_SAP": sap_row.get("Comp. Colour", "")
                })

        fuzzy_df = pd.DataFrame(fuzzy_matches)
        if not fuzzy_df.empty:
            fuzzy_df = fuzzy_df.reindex(
                fuzzy_df["ConsumptionDiff"].abs().sort_values(ascending=False).index
            )

        # --- Summary Counts for Streamlit ---
        matched_count = len(direct_matches) + len(fuzzy_df) if not fuzzy_df.empty else len(direct_matches)
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
            direct_matches.to_excel(writer, sheet_name="Direct_Matches", index=False)
            missing_in_sap.to_excel(writer, sheet_name="PLM_Not_in_SAP", index=False)
            missing_in_plm.to_excel(writer, sheet_name="SAP_Not_in_PLM", index=False)
            if not fuzzy_df.empty:
                fuzzy_df.to_excel(writer, sheet_name="70%_or_more_Matches", index=False)
        output.seek(0)

        # --- Conditional Formatting ---
        wb = load_workbook(output)

        def apply_coloring(ws, headers, plm_col_name, sap_col_name, diff_col_name, flag_col_name):
            if not all(col in headers for col in [plm_col_name, sap_col_name, diff_col_name, flag_col_name]):
                return
            plm_col = headers.index(plm_col_name) + 1
            sap_col = headers.index(sap_col_name) + 1
            diff_col = headers.index(diff_col_name) + 1
            flag_col = headers.index(flag_col_name) + 1

            red_fill = PatternFill(start_color="FF6666", end_color="FF6666", fill_type="solid")
            green_fill = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")

            for row in range(2, ws.max_row + 1):
                flag = ws.cell(row=row, column=flag_col).value
                fill = red_fill if flag == "RED" else green_fill
                ws.cell(row=row, column=plm_col).fill = fill
                ws.cell(row=row, column=sap_col).fill = fill
                ws.cell(row=row, column=diff_col).fill = fill
                ws.cell(row=row, column=flag_col).fill = fill

        if "Direct_Matches" in wb.sheetnames:
            ws = wb["Direct_Matches"]
            headers = [cell.value for cell in ws[1]]
            apply_coloring(ws, headers, "Qty(Cons.)", "Comp.Qty.", "ConsumptionDiff", "DifferenceFlag")

        if "70%_or_more_Matches" in wb.sheetnames:
            ws = wb["70%_or_more_Matches"]
            headers = [cell.value for cell in ws[1]]
            apply_coloring(ws, headers, "Qty(Cons.)_PLM", "Comp.Qty._SAP", "ConsumptionDiff", "DifferenceFlag")

        # Save workbook with formatting
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
