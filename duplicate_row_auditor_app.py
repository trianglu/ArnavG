
import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Duplicate Row Auditor", layout="wide")

st.title("📊 Duplicate Row Auditor")
st.write("Upload up to 50 Excel files and extract duplicate groups where multiple rows are marked 'Sold To'.")

uploaded_files = st.file_uploader(
    "Upload Excel files",
    type=["xlsx"],
    accept_multiple_files=True
)

if uploaded_files:
    if len(uploaded_files) > 50:
        st.error("⚠️ Please upload 50 files or fewer.")
    else:
        if st.button("Run Audit"):
            output_groups = []
            summary = []

            for file in uploaded_files:
                df = pd.read_excel(file, engine="openpyxl")

                if "group_id" not in df.columns or "AccountGroup" not in df.columns:
                    st.warning(f"Skipping {file.name} (missing required columns)")
                    continue

                grouped = df.groupby("group_id")
                total_groups = 0
                flagged_groups = 0

                for gid, g in grouped:
                    if str(gid).upper() == "UNIQUE":
                        continue

                    total_groups += 1
                    sold_to_count = (g["AccountGroup"].astype(str).str.lower() == "sold to").sum()

                    if sold_to_count > 1:
                        flagged_groups += 1
                        output_groups.append(g)

                summary.append((file.name, total_groups, flagged_groups))

            if output_groups:
                result = pd.concat(output_groups, ignore_index=True)
                output = io.BytesIO()
                result.to_excel(output, index=False, engine="openpyxl")
                output.seek(0)

                st.success("✅ Audit complete!")

                st.download_button(
                    label="Download Results",
                    data=output,
                    file_name="Multiple Sold To Duplicates.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.info("No groups found with multiple 'Sold To' rows.")

            st.subheader("📋 Summary")
            for name, total, flagged in summary:
                st.write(f"**{name}** → Groups: {total}, Flagged: {flagged}")
