import streamlit as st
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill
import io
import pandas as pd

st.set_page_config(page_title="Duplicate Row Auditor", layout="wide")

st.title("📊 Duplicate Row Auditor")

# ✅ SETTINGS (performance controls)
group_filter = st.selectbox(
    "Filter Groups",
    ["All", "🚨 Critical Only", "⚠️ Warning Only"]
)

show_highlighting = st.checkbox("Highlight Sold-To rows (slower)", value=False)

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

# ✅ CACHE FILE PROCESSING
@st.cache_data
def process_file(file_bytes, filename):

    wb = load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows())

    if not rows:
        return None

    raw_headers = [cell.value for cell in rows[0]]
    headers = clean_headers(raw_headers)
    normalized_headers = [normalize(h) for h in raw_headers]

    if "accountgroup" not in normalized_headers:
        return None

    acc_idx = normalized_headers.index("accountgroup")

    groups = []
    current_group = []

    # Faster highlight detection
    def is_highlighted(row):
        return any(cell.fill and cell.fill.fill_type for cell in row)

    for row in rows[1:]:
        if is_highlighted(row):
            current_group.append(row)
        else:
            if current_group:
                groups.append(current_group)
                current_group = []

    if current_group:
        groups.append(current_group)

    output = []

    for group in groups:
        sold_to_count = sum(
            1 for r in group
            if r[acc_idx].value and str(r[acc_idx].value).strip().lower() == "sold to"
        )

        if sold_to_count > 1:
            output.append((filename, group, headers, acc_idx, sold_to_count))

    return output


# ✅ STORE RESULTS (prevents recompute on refresh)
if "results" not in st.session_state:
    st.session_state.results = None


# ✅ RUN BUTTON
if uploaded_files:
    if st.button("Run Audit"):

        with st.spinner("🔄 Processing..."):

            all_results = []
            summary = []

            for file in uploaded_files:

                file_bytes = file.read()

                result = process_file(file_bytes, file.name)

                if result is None:
                    st.warning(f"⚠️ {file.name} skipped (missing AccountGroup column)")
                    continue

                all_results.extend(result)

            # sort by severity
            all_results = sorted(all_results, key=lambda x: x[4], reverse=True)

            st.session_state.results = all_results

            st.success("✅ Audit complete!")


# ✅ MAIN DISPLAY (no recompute)
if st.session_state.results:

    all_results = st.session_state.results

    # ✅ FILTER
    filtered = all_results

    if group_filter == "🚨 Critical Only":
        filtered = [x for x in all_results if x[4] >= 3]
    elif group_filter == "⚠️ Warning Only":
        filtered = [x for x in all_results if x[4] == 2]

    # ✅ PREVIEW
    st.subheader("🔍 Preview of Flagged Groups")

    group_counter = {}

    for file_name, group, headers, acc_idx, sold_to_count in filtered:

        group_counter[file_name] = group_counter.get(file_name, 0) + 1
        group_num = group_counter[file_name]

        severity = "🚨 Critical" if sold_to_count >= 3 else "⚠️ Warning"
        color = "#FFCDD2" if sold_to_count >= 3 else "#FFF9C4"

        with st.expander(f"{file_name} — Group {group_num} | {severity}"):

            st.markdown(f"""
            <div style="padding:10px; border-radius:8px; background-color:{color};">
            <b>Sold To Count:</b> {sold_to_count}<br>
            <b>Status:</b> {severity}
            </div>
            """, unsafe_allow_html=True)

            group_data = [[cell.value for cell in row] for row in group]

            # ✅ LIMIT rows (huge speed gain)
            MAX_ROWS = 75
            if len(group_data) > MAX_ROWS:
                st.warning(f"Showing first {MAX_ROWS} rows")

            df = pd.DataFrame(group_data[:MAX_ROWS], columns=headers)

            if show_highlighting:
                def highlight(row):
                    val = str(row.iloc[acc_idx]).strip().lower()
                    return ["background-color: #FFF59D"]*len(row) if val == "sold to" else [""]*len(row)

                st.dataframe(df.style.apply(highlight, axis=1), width="stretch")
            else:
                st.dataframe(df, width="stretch")


    # ✅ EXPORT BUTTON (always visible)
    st.subheader("📥 Export")

    from openpyxl import Workbook

    if st.button("Generate Excel File"):

        out_wb = Workbook()
        out_ws = out_wb.active

        first_headers = all_results[0][2]
        out_ws.append(["Source File"] + first_headers)

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

        for file_name, group, headers, acc_idx, sold_to_count in all_results:

            out_ws.append([f"--- {file_name} ---"] + [""] * len(headers))

            for row in group:
                values = [cell.value for cell in row]
                out_ws.append([file_name] + values)

                for col_idx, cell in enumerate(row):
                    out_cell = out_ws.cell(row=out_ws.max_row, column=col_idx + 2)

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
            "Download Multiple Sold To Duplicates.xlsx",
            buffer,
            file_name="Multiple Sold To Duplicates.xlsx"
        )
