import streamlit as st
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill
import io

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
    return (
        str(col).strip().lower().replace(" ", "").replace("_", "").replace("\xa0", "")
    )

def smart_column_map(headers):
    normalized = [normalize(h) for h in headers]

    def find_best(keywords):
        for i, col in enumerate(normalized):
            for k in keywords:
                if k in col:
                    return i
        return None

    return {
        "group_id": find_best(["groupid", "group"]),
        "account_group": find_best(["accountgroup"])
    }

def is_sold_to(val):
    return val and "sold to" in str(val).lower()

def is_highlighted(row):
    return any(cell.fill and cell.fill.fill_type for cell in row)

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


# -------------------------------
# ✅ MAIN PROCESS
# -------------------------------

def process_files(file_dict_list):

    all_0, all_1, all_2 = [], [], []
    summary_stats = []
    error_files = []
    headers = None

    for file_dict in file_dict_list:

        filename = file_dict["name"]
        file_bytes = file_dict["bytes"]

        try:
            wb = load_workbook(io.BytesIO(file_bytes), data_only=True)
            ws = wb.active
        except:
            error_files.append(filename)
            continue

        headers_raw = [cell.value for cell in ws[1]]
        col_map = smart_column_map(headers_raw)

        if col_map["group_id"] is None or col_map["account_group"] is None:
            error_files.append(filename)
            continue

        if headers is None:
            headers = ["Source File"] + headers_raw

        group_idx = col_map["group_id"]
        acc_idx = col_map["account_group"]

        groups = {}

        for row in ws.iter_rows(min_row=2):
            if not is_highlighted(row):
                continue

            key = f"{filename}__{row[group_idx].value}"
            groups.setdefault(key, []).append(row)

        g0 = g1 = g2 = 0

        for group_rows in groups.values():

            sold_to_count = sum(
                1 for r in group_rows if is_sold_to(r[acc_idx].value)
            )

            if sold_to_count >= 2:
                all_2.append([(filename, r) for r in group_rows])
                g2 += 1
            elif sold_to_count == 1:
                all_1.append([(filename, r) for r in group_rows])
                g1 += 1
            else:
                all_0.append([(filename, r) for r in group_rows])
                g0 += 1

        summary_stats.append({
            "file": filename,
            "total": len(groups),
            "0": g0,
            "1": g1,
            "2": g2
        })

    return headers, all_0, all_1, all_2, summary_stats, error_files


# -------------------------------
# ✅ UI
# -------------------------------

if uploaded_files:

    if st.button("⚡ Generate Merged Excel"):

        progress_bar = st.progress(0)
        file_dict_list = []

        for i, file in enumerate(uploaded_files):
            file_dict_list.append({
                "name": file.name,
                "bytes": file.getvalue()
            })
            progress_bar.progress((i + 1) / len(uploaded_files) * 0.3)

        headers, all_0, all_1, all_2, summary, errors = process_files(file_dict_list)

        wb = Workbook()
        wb.remove(wb.active)

        # ✅ COLOR DEFINITIONS
        red_fill = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
        yellow_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

        def write_sheet(name, groups, severity_fill):
            ws = wb.create_sheet(name)

            if headers:
                for i, h in enumerate(headers, start=1):
                    ws.cell(1, i, h)

            row_cursor = 2

            for group in groups:

                for filename, row in group:

                    ws.cell(row_cursor, 1, filename)

                    for j, cell in enumerate(row, start=2):
                        new_cell = ws.cell(row_cursor, j, cell.value)

                        # ✅ ORIGINAL HIGHLIGHT PRESERVED
                        original_fill = safe_copy_fill(cell)
                        if original_fill:
                            new_cell.fill = original_fill

                        # ✅ OVERLAY WITH SEVERITY COLOR
                        if severity_fill:
                            new_cell.fill = severity_fill

                    row_cursor += 1

                row_cursor += 1

        write_sheet("2+ SoldTo Accounts", all_2, red_fill)
        write_sheet("1 SoldTo Accounts", all_1, yellow_fill)
        write_sheet("0 SoldTo Accounts", all_0, None)

        # ✅ SUMMARY
        ws_summary = wb.create_sheet("Summary")
        ws_summary.append(["File", "Total Groups", "0", "1", "2+"])

        for s in summary:
            ws_summary.append([
                s["file"], s["total"], s["0"], s["1"], s["2"]
            ])

        output = io.BytesIO()
        wb.save(output)
        output.seek(0)

        progress_bar.progress(1.0)

        if errors:
            st.warning("⚠️ Some files skipped:")
            for e in errors:
                st.write(f"❌ {e}")

        st.success("✅ Excel generated successfully!")

        st.download_button(
            "⬇️ Download Merged Excel",
            data=output,
            file_name="Multiple Sold To Duplicates.xlsx"
        )
