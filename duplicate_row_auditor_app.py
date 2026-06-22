import streamlit as st
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill
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

def safe_copy_fill(cell):
    """
    Safely recreate a fill (avoids StyleProxy crash)
    """
    try:
        if not cell.fill or not cell.fill.fill_type:
            return None

        # Handle RGB colors safely
        start = cell.fill.start_color
        end = cell.fill.end_color

        start_color = start.rgb if start.rgb else None
        end_color = end.rgb if end.rgb else None

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
                        val = row[acc_idx].value
                        if val and str(val).strip().lower() == "sold to":
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

                # Write headers
                out_ws.append(["Source File"] + headers)

                for file_name, group in output_rows:

                    # separator row
                    out_ws.append([f"--- {file_name} ---"] + [""] * len(headers))

                    for row in group:
                        values = [cell.value for cell in row]
                        out_ws.append([file_name] + values)

                        # ✅ SAFE highlight copy (NO CRASH)
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

            else:
                st.info("No duplicate groups found with multiple 'Sold To' rows.")

            # --- Summary ---
            st.subheader("📋 Summary")
            for name, total, flagged in summary:
                st.write(f"**{name}** → Groups Reviewed: {total}, Flagged: {flagged}")
