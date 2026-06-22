import streamlit as st
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill
import io
import pandas as pd

st.set_page_config(page_title="Duplicate Row Auditor", layout="wide")

st.title("📊 Duplicate Row Auditor")
st.write("One-click audit + export enabled.")

uploaded_files = st.file_uploader(
    "Upload Excel files",
    type=["xlsx"],
    accept_multiple_files=True
)

# ---------- SESSION ----------
if "results" not in st.session_state:
    st.session_state.results = []

if "summary" not in st.session_state:
    st.session_state.summary = {}

if "export_file" not in st.session_state:
    st.session_state.export_file = None

# ---------- BUTTON ----------
run_clicked = st.button("🚀 Run Audit & Generate Export")

# ---------- HELPERS ----------
def normalize(col):
    return str(col).strip().lower().replace(" ", "").replace("_", "") if col else ""

def clean_headers(headers):
    cleaned, seen = [], {}
    for i, h in enumerate(headers):
        h = f"Column_{i+1}" if not h else str(h).strip()
        if h in seen:
            seen[h] += 1
            h = f"{h}_{seen[h]}"
        else:
            seen[h] = 0
        cleaned.append(h)
    return cleaned

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

# ---------- PROCESS ----------
if run_clicked and uploaded_files:

    st.session_state.results = []
    st.session_state.summary = {}

    progress = st.progress(0)
    status = st.empty()

    master_headers = None
    master_acc_idx = None

    for i, file in enumerate(uploaded_files):

        file_critical = 0
        file_warning = 0

        try:
            status.text(f"Processing {file.name} ({i+1}/{len(uploaded_files)})")

            wb = load_workbook(file, data_only=True)
            ws = wb.active

            rows = list(ws.iter_rows())
            if not rows:
                continue

            raw_headers = [cell.value for cell in rows[0]]

            if master_headers is None:
                master_headers = clean_headers(raw_headers)
                normalized = [normalize(h) for h in raw_headers]

                if "accountgroup" not in normalized:
                    st.error("AccountGroup column missing.")
                    st.stop()

                master_acc_idx = normalized.index("accountgroup")

            headers = master_headers
            acc_idx = master_acc_idx

            current_group = []
            groups = []

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
                    severity = "CRITICAL" if sold_to_count >= 3 else "WARNING"

                    if severity == "CRITICAL":
                        file_critical += 1
                    else:
                        file_warning += 1

                    st.session_state.results.append(
                        (severity, file.name, group, headers, acc_idx, sold_to_count)
                    )

        except Exception as e:
            st.warning(f"⚠️ {file.name} failed: {str(e)}")

        st.session_state.summary[file.name] = {
            "CRITICAL": file_critical,
            "WARNING": file_warning,
            "TOTAL": file_critical + file_warning
        }

        progress.progress((i + 1) / len(uploaded_files))

    status.text("✅ Processing complete")

    # ✅ BUILD EXPORT FILE ONCE
    wb = Workbook()
    wb.remove(wb.active)

    # SUMMARY SHEET
    summary_ws = wb.create_sheet("SUMMARY")
    summary_ws.append(["File", "Critical", "Warning", "Total"])

    for fname, vals in st.session_state.summary.items():
        summary_ws.append([
            fname,
            vals["CRITICAL"],
            vals["WARNING"],
            vals["TOTAL"]
        ])

    # DATA SHEETS
    sheets = {
        "CRITICAL": wb.create_sheet("CRITICAL"),
        "WARNING": wb.create_sheet("WARNING")
    }

    headers = st.session_state.results[0][3] if st.session_state.results else []

    for name, ws in sheets.items():
        ws.append(["Priority", "Source File"] + headers)

    for severity, file_name, group, headers, acc_idx, sold_to_count in st.session_state.results:

        ws = sheets[severity]

        for row in group:
            vals = [c.value for c in row]
            ws.append([severity, file_name] + vals)

            for col_idx, cell in enumerate(row):
                out_cell = ws.cell(row=ws.max_row, column=col_idx + 3)
                fill = safe_copy_fill(cell)
                if fill:
                    try:
                        out_cell.fill = fill
                    except:
                        pass

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    st.session_state.export_file = buffer

# ---------- ONE-CLICK DOWNLOAD (TOP) ----------
if st.session_state.export_file:

    st.subheader("📥 One-Click Export")

    st.download_button(
        "⬇️ Download Compiled Excel Report",
        data=st.session_state.export_file,
        file_name="Multiple Sold To Duplicates.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# ---------- PREVIEW ----------
if st.session_state.results:

    st.subheader("🔍 Preview")

    for severity, file_name, group, headers, acc_idx, sold_to_count in st.session_state.results:

        label = "🚨 Critical" if severity == "CRITICAL" else "⚠️ Warning"

        with st.expander(f"{label} — {file_name}"):

            df = pd.DataFrame(
                [[c.value for c in row] for row in group],
                columns=headers
            )

            st.dataframe(df, width="stretch")

# ---------- BACKUP DOWNLOAD (BOTTOM) ----------
if st.session_state.export_file:

    st.download_button(
        "⬇️ Download Again",
        data=st.session_state.export_file,
        file_name="Multiple Sold To Duplicates.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
