import streamlit as st
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill
import io
import pandas as pd

st.set_page_config(page_title="Duplicate Row Auditor", layout="wide")

st.title("📊 Duplicate Row Auditor")

# --- Controls ---
group_filter = st.selectbox(
    "Filter Groups",
    ["All", "🚨 Critical Only", "⚠️ Warning Only"]
)

uploaded_files = st.file_uploader(
    "Upload Excel files",
    type=["xlsx"],
    accept_multiple_files=True
)

# --- Helpers ---
def normalize(col):
    if not col:
        return ""
    return (
        str(col)
        .strip()
        .lower()
        .replace(" ", "")
        .replace("_", "")
        .replace("\xa0", "")   # ✅ handles hidden Excel spaces
    )

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

def is_highlighted(row):
    return any(cell.fill and cell.fill.fill_type for cell in row)

def is_sold_to(val):
    return val and "sold to" in str(val).lower()

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

# ✅ caching
@st.cache_data
def process_file(file_bytes, filename):

    skipped = False

    wb = load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows())

    if not rows:
        return [], True

    raw_headers = [cell.value for cell in rows[0]]
    headers = clean_headers(raw_headers)

    normalized_headers = [normalize(h) for h in raw_headers]

    # ✅ Robust detection
    acc_idx = None
    for i, col in enumerate(normalized_headers):
        if col == "accountgroup" or ("account" in col and "group" in col):
            acc_idx = i
            break

    if acc_idx is None:
        skipped = True
        return [], True

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

    output = []

    for group in groups:
        sold_to_count = sum(
            1 for r in group if is_sold_to(r[acc_idx].value)
        )

        if sold_to_count >= 2:
            output.append((filename, group, headers, acc_idx, sold_to_count))

    return output, False


# --- Session storage ---
if "results" not in st.session_state:
    st.session_state.results = None

if "export_file" not in st.session_state:
    st.session_state.export_file = None


# --- Run Audit ---
if uploaded_files:
    if st.button("🚀 Run Audit"):

        with st.spinner("🔄 Processing..."):

            all_results = []
            skipped_files = []

            # ✅ FIRST PASS
            for file in uploaded_files:
                file_bytes = file.read()
                result, skipped = process_file(file_bytes, file.name)

                if skipped:
                    skipped_files.append((file_bytes, file.name))
                elif result:
                    all_results.extend(result)

            # ✅ SECOND PASS (retry skipped)
            if skipped_files:
                st.warning(f"Retrying {len(skipped_files)} skipped files...")

                for file_bytes, filename in skipped_files:
                    result, skipped = process_file(file_bytes, filename)

                    if result:
                        all_results.extend(result)
                    else:
                        st.error(f"❌ Could not process {filename}")

            all_results = sorted(all_results, key=lambda x: x[4], reverse=True)

            st.session_state.results = all_results

            # ✅ BUILD EXPORT FILE IMMEDIATELY
            out_wb = Workbook()
            out_ws = out_wb.active

            if all_results:
                first_headers = all_results[0][2]
                out_ws.append(["Source File"] + first_headers)

                for file_name, group, headers, acc_idx, sold_to_count in all_results:

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
            else:
                out_ws.append(["No duplicate groups found with multiple Sold To"])

            buffer = io.BytesIO()
            out_wb.save(buffer)
            buffer.seek(0)

            st.session_state.export_file = buffer

            st.success("✅ Audit complete!")


# ✅ ONE-CLICK DOWNLOAD (always visible after run)
if st.session_state.export_file:
    st.subheader("📥 Download Results")

    st.download_button(
        "⬇️ Download Multiple Sold To Duplicates.xlsx",
        st.session_state.export_file,
        file_name="Multiple Sold To Duplicates.xlsx",
        use_container_width=True
    )


# --- Display Results ---
if st.session_state.results:

    all_results = st.session_state.results

    # ✅ Filtering
    filtered = all_results
    if group_filter == "🚨 Critical Only":
        filtered = [x for x in all_results if x[4] >= 3]
    elif group_filter == "⚠️ Warning Only":
        filtered = [x for x in all_results if x[4] == 2]

    st.subheader("🔍 Preview of Flagged Groups")

    group_counter = {}

    for file_name, group, headers, acc_idx, sold_to_count in filtered:

        group_counter[file_name] = group_counter.get(file_name, 0) + 1
        group_num = group_counter[file_name]

        severity = "🚨 Critical" if sold_to_count >= 3 else "⚠️ Warning"

        with st.expander(f"{file_name} — Group {group_num} | {severity}"):

            group_data = [[cell.value for cell in row] for row in group]

            df = pd.DataFrame(group_data[:50], columns=headers)
            st.dataframe(df, width="stretch")

            st.markdown(f"**✅ Sold To Count: {sold_to_count}**")
