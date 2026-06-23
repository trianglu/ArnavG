import streamlit as st
from openpyxl import load_workbook, Workbook
import io
import time
from concurrent.futures import ThreadPoolExecutor

st.set_page_config(page_title="Duplicate Row Auditor", layout="wide")
st.title("📊 Duplicate Row Auditor")

uploaded_files = st.file_uploader(
    "Upload Excel files",
    type=["xlsx"],
    accept_multiple_files=True
)

# -------------------------------
# ✅ HELPERS
# -------------------------------

def normalize(col):
    if not col:
        return ""
    return str(col).strip().lower().replace(" ", "").replace("_", "")

def smart_column_map(headers):
    norm_headers = {normalize(h): i for i, h in enumerate(headers)}

    def find_column(possible_names):
        for name in possible_names:
            if name in norm_headers:
                return norm_headers[name]
        return None

    return {
        "group_id": find_column([
            "groupid", "group_id", "group", "groupkey"
        ]),
        "account_group": find_column([
            "accountgroup", "account_group", "type"
        ])
    }

def is_sold_to(val):
    return val and "sold to" in str(val).lower()


# ✅ Detect highlighted row (STRICT audit mode)
def is_highlighted(row):
    for cell in row:
        if cell.fill and cell.fill.fill_type is not None:
            return True
    return False


# -------------------------------
# ✅ SINGLE FILE PROCESSING
# -------------------------------

def process_single_file(file_dict):
    file_bytes = file_dict["bytes"]
    filename = file_dict["name"]

    wb = load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = wb.active

    headers = [cell.value for cell in ws[1]]

    col_map = smart_column_map(headers)

    if col_map["group_id"] is None or col_map["account_group"] is None:
        raise ValueError(f"Missing required columns in {filename}")

    group_id_idx = col_map["group_id"]
    account_group_idx = col_map["account_group"]

    groups = {}

    # ✅ ONLY process highlighted rows
    for row in ws.iter_rows(min_row=2):

        if not is_highlighted(row):
            continue

        group_id = row[group_id_idx].value
        unique_group_key = f"{filename}__{group_id}"

        if unique_group_key not in groups:
            groups[unique_group_key] = []

        groups[unique_group_key].append(row)

    # ---- classify groups
    groups_0, groups_1, groups_2_plus = [], [], []

    for group_key, rows in groups.items():
        sold_to_count = sum(
            1 for r in rows if is_sold_to(r[account_group_idx].value)
        )

        if sold_to_count == 0:
            groups_0.append(rows)
        elif sold_to_count == 1:
            groups_1.append(rows)
        else:
            groups_2_plus.append(rows)

    stats = {
        "file": filename,
        "total_groups": len(groups),
        "0": len(groups_0),
        "1": len(groups_1),
        "2plus": len(groups_2_plus)
    }

    return headers, groups_0, groups_1, groups_2_plus, stats


# -------------------------------
# ✅ MAIN PARALLEL ENGINE
# -------------------------------

@st.cache_data(show_spinner=False)
def process_files_parallel(file_dict_list):
    start_time = time.time()

    all_0, all_1, all_2 = [], [], []
    summary_stats = []
    headers = None

    # ✅ PARALLEL PROCESSING
    with ThreadPoolExecutor() as executor:
        results = list(executor.map(process_single_file, file_dict_list))

    for h, g0, g1, g2, stats in results:

        if headers is None:
            headers = ["Source File"] + h

        filename = stats["file"]

        def attach_source(groups):
            new_groups = []
            for group in groups:
                new_group = []
                for row in group:
                    new_group.append((filename, row))
                new_groups.append(new_group)
            return new_groups

        all_0.extend(attach_source(g0))
        all_1.extend(attach_source(g1))
        all_2.extend(attach_source(g2))

        summary_stats.append(stats)

    # -------------------------------
    # ✅ BUILD OUTPUT WORKBOOK
    # -------------------------------

    out_wb = Workbook()
    out_wb.remove(out_wb.active)

    def write_sheet(name, grouped_rows):
        ws_out = out_wb.create_sheet(name)

        # write headers
        for col_idx, header in enumerate(headers, start=1):
            ws_out.cell(row=1, column=col_idx, value=header)

        row_cursor = 2

        for group in grouped_rows:
            for (filename, row) in group:

                ws_out.cell(row=row_cursor, column=1, value=filename)

                for col_idx, cell in enumerate(row, start=2):
                    new_cell = ws_out.cell(
                        row=row_cursor,
                        column=col_idx,
                        value=cell.value
                    )

                    # ✅ PRESERVE COLORS
                    if cell.fill:
                        new_cell.fill = cell.fill

                row_cursor += 1

            row_cursor += 1

    write_sheet("0 SoldTo Accounts", all_0)
    write_sheet("1 SoldTo Accounts", all_1)
    write_sheet("2+ SoldTo Accounts", all_2)

    # -------------------------------
    # ✅ SUMMARY SHEET
    # -------------------------------

    ws_summary = out_wb.create_sheet("Summary")

    ws_summary.append([
        "File", "Total Groups", "0 SoldTo", "1 SoldTo", "2+ SoldTo"
    ])

    for stat in summary_stats:
        ws_summary.append([
            stat["file"],
            stat["total_groups"],
            stat["0"],
            stat["1"],
            stat["2plus"]
        ])

    # TOTAL ROW
    ws_summary.append([])
    ws_summary.append([
        "TOTAL",
        sum(s["total_groups"] for s in summary_stats),
        sum(s["0"] for s in summary_stats),
        sum(s["1"] for s in summary_stats),
        sum(s["2plus"] for s in summary_stats),
    ])

    # -------------------------------
    # ✅ SAVE
    # -------------------------------

    output = io.BytesIO()
    out_wb.save(output)
    output.seek(0)

    end_time = time.time()

    stats = {
        "total_files": len(file_dict_list),
        "time": round(end_time - start_time, 2)
    }

    return output, stats


# -------------------------------
# ✅ STREAMLIT UI
# -------------------------------

if uploaded_files:

    if st.button("⚡ Generate Merged Excel"):

        progress_bar = st.progress(0)
        status_text = st.empty()

        file_dict_list = []
        total_files = len(uploaded_files)

        # STEP 1: Load files
        for i, file in enumerate(uploaded_files):
            status_text.text(f"Loading {file.name}...")
            file_dict_list.append({
                "name": file.name,
                "bytes": file.getvalue()
            })
            progress_bar.progress((i + 1) / total_files * 0.4)

        # STEP 2: Process
        status_text.text("Processing highlighted duplicate groups (parallel)...")
        output_file, stats = process_files_parallel(file_dict_list)
        progress_bar.progress(0.85)

        # STEP 3: Finish
        status_text.text("Finalizing Excel...")
        progress_bar.progress(1.0)

        st.success("✅ Excel generated successfully!")

        # Metrics
        col1, col2 = st.columns(2)
        col1.metric("Files Processed", stats["total_files"])
        col2.metric("Processing Time (sec)", stats["time"])

        st.download_button(
            label="⬇️ Download Merged Excel",
            data=output_file,
            file_name="Multiple Sold To Duplicates.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
