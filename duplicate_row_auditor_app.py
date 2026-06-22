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

# --- Clean headers (fix duplicates + blanks) ---
def clean_headers(headers):
    cleaned = []
    seen = {}

    for i, h in enumerate(headers):
        if h is None or str(h).strip() == "":
            h = f"Column_{i+1}"
        else:
            h = str(h).strip()

        if h in seen:
            seen[h] += 1
            h = f"{h}_{seen[h]}"
        else:
            seen[h] = 0

        cleaned.append(h)

    return cleaned

# --- Detect highlighted rows ---
def is_highlighted(row):
    return any(cell.fill and cell.fill.fill_type for cell in row)

# --- Safe fill copy ---
def safe_copy_fill(cell):
    try:
        if not cell.fill or not cell.fill.fill_type:
            return None

        start = cell.fill.start_color
        end = cell.fill.end_color

        return PatternFill(
            start_color=start.rgb if start and start.rgb else None,
            end_color=end.rgb if end and end.rgb else None,
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

                raw_headers = [cell.value for cell in rows[0]]
                headers = clean_headers(raw_headers)

                normalized_headers = [normalize(h) for h in raw_headers]

                if "accountgroup" not in normalized_headers:
                    st.warning(f"⚠️ {file.name} skipped (missing AccountGroup column)")
                    continue

                acc_idx = normalized_headers.index("accountgroup")

                groups = []
                current_group = []

                # --- Identify groups from highlights ---
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

                for group in groups:

                    sold_to_count = 0

                    for row in group:
                        val = row[acc_idx].value
                        if val and str(val).strip().lower() == "sold to":
                            sold_to_count += 1

                    if sold_to_count > 1:
                        flagged_groups += 1
                        output_rows.append((file.name, group, headers, acc_idx, sold_to_count))

                summary.append((file.name, total_groups, flagged_groups))

            # ✅ SORT: Critical first, then Warning
            output_rows = sorted(
                output_rows,
                key=lambda x: x[4],  # sold_to_count
                reverse=True
            )

            # ✅ GROUPED PREVIEW WITH SEVERITY
            if output_rows:
                st.subheader("🔍 Preview of Flagged Duplicate Groups")

                group_counter = {}

                for file_name, group, headers, acc_idx, sold_to_count in output_rows:

                    group_counter[file_name] = group_counter.get(file_name, 0) + 1
                    group_num = group_counter[file_name]

                    # --- Severity logic ---
                    if sold_to_count >= 3:
                        severity = "🚨 Critical"
                        color = "#FFCDD2"
                    else:
                        severity = "⚠️ Warning"
                        color = "#FFF9C4"

                    with st.expander(f"📁 {file_name} — Group {group_num} | {severity}"):

                        # --- Header block ---
                        st.markdown(f"""
                        <div style="padding:10px; border-radius:8px; background-color:{color};">
                        <b>File:</b> {file_name}<br>
                        <b>Group #:</b> {group_num}<br>
                        <b>Sold To Count:</b> {sold_to_count}<br>
                        <b>Status:</b> {severity}
                        </div>
                        """, unsafe_allow_html=True)

                        group_data = [[cell.value for cell in row] for row in group]
                        df = pd.DataFrame(group_data, columns=headers)

                        # --- Highlight Sold-To rows in preview ---
                        def highlight_soldto(row):
                            val = str(row.iloc[acc_idx]).strip().lower()
                            if val == "sold to":
                                return ["background-color: #FFF59D"] * len(row)
                            return [""] * len(row)

                        styled_df = df.style.apply(highlight_soldto, axis=1)

                        st.dataframe(styled_df, width="stretch")

            else:
                st.info("No duplicate groups found with multiple 'Sold To' rows.")

            # ✅ EXPORT FILE
            if output_rows:

                out_wb = Workbook()
                out_ws = out_wb.active
                out_ws.title = "Flagged Groups"

                first_headers = output_rows[0][2]
                out_ws.append(["Source File"] + first_headers)

                for file_name, group, headers, acc_idx, sold_to_count in output_rows:

                    out_ws.append([f"--- {file_name} ---"] + [""] * len(headers))

                    for row in group:
                        values = [cell.value for cell in row]
                        out_ws.append([file_name] + values)

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
