import streamlit as st
from openpyxl import load_workbook, Workbook
import io

st.set_page_config(page_title="Duplicate Row Auditor", layout="wide")

st.title("📊 Duplicate Row Auditor")
st.write("Upload up to 50 Excel files. All highlighted rows (any fill color) will be treated as duplicate groups.")

uploaded_files = st.file_uploader(
    "Upload Excel files",
    type=["xlsx"],
    accept_multiple_files=True
)

def is_highlighted(row):
    # ANY filled cell → counts as highlighted
    return any(cell.fill and cell.fill.fill_type for cell in row)

if uploaded_files:
    if len(uploaded_files) > 50:
        st.error("⚠️ Please upload 50 files or fewer.")
    else:
        if st.button("Run Audit"):

            output_rows = []
            summary = []
            headers_written = False
            headers = None

            for file in uploaded_files:

                wb = load_workbook(file, data_only=True)
                ws = wb.active

                rows = list(ws.iter_rows())

                headers = [cell.value for cell in rows[0]]

                if "AccountGroup" not in headers:
                    st.warning(f"⚠️ {file.name} skipped (missing AccountGroup column)")
                    continue

                acc_idx = headers.index("AccountGroup")

                groups = []
                current_group = []

                # --- Step 1: Identify highlighted groups ---
                for row in rows[1:]:

                    if is_highlighted(row):
                        current_group.append(row)
                    else:
                        if current_group:
                            groups.append(current_group)
                            current_group = []

                if current_group:
                    groups.append(current_group)

                total_groups = len(groups)
                flagged_groups = 0

                # --- Step 2: Apply Sold To rule ---
                for group in groups:

                    sold_to_count = 0

                    for row in group:
                        value = str(row[acc_idx].value).strip().lower()
                        if value == "sold to":
                            sold_to_count += 1

                    if sold_to_count > 1:
                        flagged_groups += 1
                        output_rows.append((file.name, group))

                summary.append((file.name, total_groups, flagged_groups))

            # --- Step 3: Build output file ---
            if output_rows:

                out_wb = Workbook()
                out_ws = out_wb.active

                out_ws.title = "Flagged Groups"

                # Write headers once
                out_ws.append(["Source File"] + headers)

                for file_name, group in output_rows:

                    # Optional separator row for readability
                    out_ws.append([f"--- {file_name} ---"] + [""] * (len(headers)))

                    for row in group:
                        values = [cell.value for cell in row]
                        out_ws.append([file_name] + values)

                        # Preserve highlights
                        for col_idx, cell in enumerate(row):
                            out_cell = out_ws.cell(
                                row=out_ws.max_row,
                                column=col_idx + 2
                            )
                            out_cell.fill = cell.fill

                buffer = io.BytesIO()
                out_wb.save(buffer)
                buffer.seek(0)

                st.success("✅ Audit complete!")

                st.download_button(
                    label="Download Multiple Sold To Duplicates.xlsx",
                    data=buffer,
                    file_name="Multiple Sold To Duplicates.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            else:
                st.info("No duplicate groups found with multiple 'Sold To' rows.")

            # --- Summary ---
            st.subheader("📋 Summary")
            for name, total, flagged in summary:
                st.write(f"**{name}** → Groups Reviewed: {total}, Flagged: {flagged}")
