import streamlit as st
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill
import io
import pandas as pd

st.set_page_config(page_title="Duplicate Row Auditor", layout="wide")

st.title("📊 Duplicate Row Auditor")
st.write("Optimized for large uploads (40–50 files).")

uploaded_files = st.file_uploader(
    "Upload Excel files",
    type=["xlsx"],
    accept_multiple_files=True
)

# ---------- SESSION STATE ----------
if "results" not in st.session_state:
    st.session_state.results = []

if "processed" not in st.session_state:
    st.session_state.processed = 0

if "run" not in st.session_state:
    st.session_state.run = False

# ---------- BUTTON ----------
if st.button("Run Audit"):
    st.session_state.run = True
    st.session_state.results = []
    st.session_state.processed = 0

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

# ---------- PROCESSING ----------
if uploaded_files and st.session_state.run:

    progress_bar = st.progress(0)
    status = st.empty()

    master_headers = None
    master_acc_idx = None

    for i, file in enumerate(uploaded_files):

        try:
            status.text(f"Processing {file.name} ({i+1}/{len(uploaded_files)})")

            wb = load_workbook(file, data_only=True)
            ws = wb.active

            rows = list(ws.iter_rows())
            if not rows:
                continue

            raw_headers = [cell.value for cell in rows[0]]

            # ✅ Detect schema ONCE
            if master_headers is None:
                master_headers = clean_headers(raw_headers)

                normalized_headers = [normalize(h) for h in raw_headers]
                if "accountgroup" not in normalized_headers:
                    st.error("AccountGroup column missing in first file.")
                    st.stop()

                master_acc_idx = normalized_headers.index("accountgroup")

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

                    st.session_state.results.append(
                        (severity, file.name, group, headers, acc_idx, sold_to_count)
                    )

        except Exception as e:
            st.warning(f"⚠️ Error processing {file.name}: {str(e)}")

        # ✅ cleanup memory
        try:
            del wb
            del ws
        except:
            pass

        st.session_state.processed += 1
        progress_bar.progress((i + 1) / len(uploaded_files))

    status.text("✅ Processing complete!")
    st.session_state.run = False  # ✅ prevent rerun loops

# ---------- DISPLAY ----------
if st.session_state.results:

    priority_order = {"CRITICAL": 0, "WARNING": 1}
    results = sorted(st.session_state.results, key=lambda x: priority_order[x[0]])

    st.subheader("🔍 Preview of Flagged Duplicate Groups")

    group_counter = {}

    for severity, file_name, group, headers, acc_idx, sold_to_count in results:

        group_counter[file_name] = group_counter.get(file_name, 0) + 1
        group_num = group_counter[file_name]

        label = "🚨 Critical" if severity == "CRITICAL" else "⚠️ Warning"
        color = "#FFCDD2" if severity == "CRITICAL" else "#FFF9C4"

        with st.expander(f"{label} — 📁 {file_name} — Group {group_num}"):

            st.markdown(f"""
            <div style="padding:10px; border-radius:8px; background-color:{color};">
            <b>File:</b> {file_name}<br>
            <b>Group #:</b> {group_num}<br>
            <b>Sold To Count:</b> {sold_to_count}<br>
            <b>Status:</b> {label}
            </div>
            """, unsafe_allow_html=True)

            df = pd.DataFrame(
                [[cell.value for cell in row] for row in group],
                columns=headers
            )

            def highlight_soldto(row):
                val = str(row.iloc[acc_idx]).strip().lower()
                return ["background-color: #FFF59D"] * len(row) if val == "sold to" else [""] * len(row)

            st.dataframe(df.style.apply(highlight_soldto, axis=1), width="stretch")

    # ---------- EXPORT (FIXED) ----------
    out_wb = Workbook()
    out_wb.remove(out_wb.active)

    sheets = {
        "CRITICAL": out_wb.create_sheet("CRITICAL"),
        "WARNING": out_wb.create_sheet("WARNING")
    }

    first_headers = results[0][3]

    for name, ws in sheets.items():
        ws.append(["Priority", "Source File"] + first_headers)

    for severity, file_name, group, headers, acc_idx, sold_to_count in results:
        ws = sheets[severity]

        for row in group:
            values = [cell.value for cell in row]
            ws.append([severity, file_name] + values)

            for col_idx, cell in enumerate(row):
                out_cell = ws.cell(row=ws.max_row, column=col_idx + 3)

                fill = safe_copy_fill(cell)
                if fill:
                    try:
                        out_cell.fill = fill
                    except:
                        pass

    buffer = io.BytesIO()
    out_wb.save(buffer)
    buffer.seek(0)

    st.success(f"✅ Processed {st.session_state.processed} files")

    # ✅ ALWAYS VISIBLE NOW
    st.download_button(
        label="📥 Download Multiple Sold To Duplicates.xlsx",
        data=buffer,
        file_name="Multiple Sold To Duplicates.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
