import streamlit as st
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill
import io
import pandas as pd

st.set_page_config(page_title="Duplicate Row Auditor", layout="wide")

st.title("📊 Duplicate Row Auditor")
st.write("Upload up to 50 Excel files. Highlighted rows (any fill color) are treated as duplicate groups.")

uploaded_files = st.file_uploader(
    "Upload Excel files",
    type=["xlsx"],
    accept_multiple_files=True
)

# --- Normalize column names ---
def normalize(col):
    if not col:
        return ""
    return str(col).strip().lower().replace(" ", "").replace("_", "")

# --- Detect highlighted rows ---
def is_highlighted(row):
    return any(cell.fill and cell.fill.fill_type for cell in row)

# --- SAFE fill copy (fixes crash) ---
def safe_copy_fill(cell):
    try:
        if not cell.fill or not cell.fill.fill_type:
            return None

        start = cell.fill.start_color
        end = cell.fill.end_color

        start_color = start.rgb if start and start.rgb else None
        end_color = end.rgb if end and end.rgb else None

        return PatternFill(
            start_color=start_color,
            end_color=end_color,
            fill_type=cell.fill.fill_type
        )
    except:
        return None


if uploaded_files:
    if len(uploaded_files) > 50:
        st.error("⚠️ Please upload 50 files or fewer.")
    else:
        if st.button("Run Audit"):

            output_rows = []
            summary = []
            headers = None

            for file in uploaded_files:

                try:
                    wb = load_workbook(file, data_only=True)
                    ws = wb.active
                except:
                    st.warning(f"⚠️ Could not read {file.name}")
                    continue

                rows = list(ws.iter_rows())

                if not rows:
                    continue

                headers_raw = [cell.value for cell in rows[0]]
                headers = headers_raw

                normalized_headers = [normalize(h) for h in headers_raw]

                # ✅ FIXED AccountGroup detection
                if "accountgroup" not in normalized_headers:
                    st.warning(f"⚠️ {file.name} skipped (missing AccountGroup column)")
                    continue

                acc_idx = normalized_headers.index("accountgroup")

                groups = []
                current_group = []

                # --- Identify highlighted groups ---
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

                # --- Apply Sold To rule ---
                for group in groups:
                    sold_to_count = 0

                    for row in group:
                        val = row[acc_idx].value
                        if val and str(val).strip().lower() == "sold to":
                            sold_to_count += 1

                    if sold_to_count > 1:
                        flagged_groups += 1
                        output_rows.append((file.name, group))

                summary.append((file.name, total_groups, flagged_groups))

            # ✅ GROUPED PREVIEW PANEL
            if output_rows:
                st.subheader("🔍 Preview of Flagged Duplicate Groups")

                group_counter = {}

                for file_name, group in output_rows:

                    if file_name not in group_counter:
                        group_counter[file_name] = 1
                    else:
                        group_counter[file_name] += 1

                    group_num = group_counter[file_name]

                    with st.expander(f"📁 {file_name} — Group {group_num}", expanded=False):

                        group_data = []
                        for row in group:
                            group_data.append([cell.value for cell in row])

                        group_df = pd.DataFrame(group_data, columns=headers)
                        st.dataframe(group_df, use_container_width=True)

                        # ✅ Sold To count display
                        sold_to_count = sum(
                            1 for r in group
                            if r[acc_idx].value and str(r[acc_idx].value).strip().lower() == "sold to"
                        )

                        st.markdown(f"**✅ Sold To Count: {sold_to_count}**")

            else:
                st.info("No duplicate groups found with multiple 'Sold To' rows.")

            # ✅ EXPORT OUTPUT
            if output_rows:

                out_wb = Workbook()
                out_ws = out_wb.active
                out_ws.title = "Flagged Groups"

                out_ws.append(["Source File"] + headers)

                for file_name, group in output_rows:

                    out_ws.append([f"--- {file_name} ---"] + [""] * len(headers))

                    for row in group:
                        values = [cell.value for cell in row]
                        out_ws.append([file_name] + values)

                        # ✅ SAFE highlight copy
                        for col_idx, cell in enumerate(row):
                            out_cell = out_ws.cell(
                                row=out_ws.max_row,
                                column=col_idx + 2
                            )

                            fill = safe_copy_fill(cell)
                            if fill:
                                try:
                                    out_cell.fill = fill
                                except:
                                    pass

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

            # ✅ SUMMARY
            st.subheader("📋 Summary")
            for name, total, flagged in summary:
                st.write(f"**{name}** → Groups Reviewed: {total}, Flagged: {flagged}")
