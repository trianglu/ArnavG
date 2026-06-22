import streamlit as st
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill
import io
import pandas as pd

st.set_page_config(page_title="Duplicate Row Auditor", layout="wide")

st.title("📊 Duplicate Row Auditor")
st.write("Upload up to 50 Excel files. Highlighted rows are treated as duplicate groups and prioritized.")

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

# --- Clean headers ---
def clean_headers(headers):
    cleaned = []
    seen = {}

    for i, h in enumerate(headers):
        if not h or str(h).strip() == "":
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

# --- Detect highlight ---
def is_highlighted(row):
    return any(cell.fill and cell.fill.fill_type for cell in row)

# --- Safe fill copy ---
def safe_copy_fill(cell):
    try:
        if not cell.fill or not cell.fill.fill_type:
            return None

        return PatternFill(
            start_color=cell.fill.start_color.rgb,
            end_color=cell.fill.end_color.rgb,
            fill_type=cell.fill.fill_type
        )
    except:
        return None

# --- Session state ---
if "results" not in st.session_state:
    st.session_state.results = None

# --- Button trigger ---
if st.button("Run Audit"):
    st.session_state.results = "processing"

# --- Run audit ---
if st.session_state.results == "processing":

    output_groups = []
    summary = []

    for file in uploaded_files:

        try:
            wb = load_workbook(file, data_only=True)
            ws = wb.active
        except:
            continue

        rows = list(ws.iter_rows())
        if not rows:
            continue

        raw_headers = [cell.value for cell in rows[0]]
        headers = clean_headers(raw_headers)
        normalized_headers = [normalize(h) for h in raw_headers]

        if "accountgroup" not in normalized_headers:
            continue

        acc_idx = normalized_headers.index("accountgroup")

        groups = []
        current_group = []

        for row in rows[1:]:
            if is_highlighted(row):
                current_group.append(row)
            else:
                if current_group:
                    groups.append(current_group)
                    current_group = []

        if current_group:
            groups.append(current_group)

        for group in groups:

            sold_to_count = sum(
                1 for r in group
                if r[acc_idx].value and str(r[acc_idx].value).strip().lower() == "sold to"
            )

            if sold_to_count > 1:
                priority = "CRITICAL"
            elif sold_to_count == 1:
                priority = "MEDIUM"
            else:
                priority = "LOW"

            output_groups.append((priority, file.name, group))

    st.session_state.results = {
        "groups": output_groups,
        "headers": headers,
        "acc_idx": acc_idx
    }

# --- Display results ---
if isinstance(st.session_state.results, dict):

    data = st.session_state.results
    output_groups = data["groups"]
    headers = data["headers"]
    acc_idx = data["acc_idx"]

    if not output_groups:
        st.warning("⚠️ No duplicate groups detected.")
    else:

        priority_order = {"CRITICAL": 0, "MEDIUM": 1, "LOW": 2}
        output_groups.sort(key=lambda x: priority_order[x[0]])

        st.subheader("🔍 Grouped Preview (Sorted by Priority)")

        group_counter = {}

        for priority, file_name, group in output_groups:

            group_counter[file_name] = group_counter.get(file_name, 0) + 1
            group_num = group_counter[file_name]

            with st.expander(f"{priority} — 📁 {file_name} — Group {group_num}"):

                group_data = [[cell.value for cell in row] for row in group]
                df = pd.DataFrame(group_data, columns=headers)

                st.dataframe(df, width="stretch")

                sold_to_count = sum(
                    1 for r in group
                    if r[acc_idx].value and str(r[acc_idx].value).strip().lower() == "sold to"
                )

                st.markdown(f"**Sold To Count: {sold_to_count}**")

    # ✅ EXPORT FILE
    out_wb = Workbook()
    out_ws = out_wb.active
    out_ws.title = "Prioritized Groups"

    out_ws.append(["Priority", "Source File"] + headers)

    for priority, file_name, group in output_groups:

        for row in group:
            values = [cell.value for cell in row]
            out_ws.append([priority, file_name] + values)

            for col_idx, cell in enumerate(row):
                out_cell = out_ws.cell(
                    row=out_ws.max_row,
                    column=col_idx + 3
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

    st.download_button(
        "Download Prioritized Duplicate Groups.xlsx",
        data=buffer,
        file_name="Multiple Sold To Duplicates.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
