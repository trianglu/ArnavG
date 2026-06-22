import streamlit as st
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill
import io
import pandas as pd

st.set_page_config(page_title="Duplicate Row Auditor", layout="wide")

st.title("📊 Duplicate Row Auditor")
st.write("Upload up to 50 Excel files. Highlighted rows are treated as duplicate groups.")

uploaded_files = st.file_uploader(
    "Upload Excel files",
    type=["xlsx"],
    accept_multiple_files=True
)

# ---------- SESSION STATE FIX (CRITICAL) ----------
if "run_audit" not in st.session_state:
    st.session_state.run_audit = False

if st.button("Run Audit"):
    st.session_state.run_audit = True

# ---------- HELPERS ----------
def normalize(col):
    if not col:
        return ""
    return str(col).strip().lower().replace(" ", "").replace("_", "")

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

# ---------- MAIN LOGIC ----------
if uploaded_files and st.session_state.run_audit:

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

        # --- GROUP EXTRACTION ---
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

                # Severity logic
                if sold_to_count >= 3:
                    severity = "CRITICAL"
                else:
                    severity = "WARNING"

                output_rows.append(
                    (severity, file.name, group, headers, acc_idx, sold_to_count)
                )

        summary.append((file.name, total_groups, flagged_groups))

    # ---------- SORT ----------
    priority_order = {"CRITICAL": 0, "WARNING": 1}
    output_rows.sort(key=lambda x: priority_order[x[0]])

    # ---------- PREVIEW ----------
    if output_rows:
        st.subheader("🔍 Preview of Flagged Duplicate Groups")

        group_counter = {}

        for severity, file_name, group, headers, acc_idx, sold_to_count in output_rows:

            group_counter[file_name] = group_counter.get(file_name, 0) + 1
            group_num = group_counter[file_name]

            color = "#FFCDD2" if severity == "CRITICAL" else "#FFF9C4"
            label = "🚨 Critical" if severity == "CRITICAL" else "⚠️ Warning"

            with st.expander(f"{label} — 📁 {file_name} — Group {group_num}"):

                st.markdown(f"""
                <div style="padding:10px; border-radius:8px; background-color:{color};">
                <b>File:</b> {file_name}<br>
                <b>Group #:</b> {group_num}<br>
                <b>Sold To Count:</b> {sold_to_count}<br>
                <b>Status:</b> {label}
                </div>
                """, unsafe_allow_html=True)

                group_data = [[cell.value for cell in row] for row in group]
                df = pd.DataFrame(group_data, columns=headers)

                def highlight_soldto(row):
                    val = str(row.iloc[acc_idx]).strip().lower()
                    if val == "sold to":
                        return ["background-color: #FFF59D"] * len(row)
                    return [""] * len(row)

                st.dataframe(df.style.apply(highlight_soldto, axis=1), width="stretch")

    else:
        st.warning("⚠️ No groups found with multiple 'Sold To' rows.")

    # ---------- EXPORT (FIXED) ----------
    if output_rows:

        out_wb = Workbook()
        out_wb.remove(out_wb.active)

        sheets = {
            "CRITICAL": out_wb.create_sheet("CRITICAL"),
            "WARNING": out_wb.create_sheet("WARNING")
        }

        first_headers = output_rows[0][3]

        for name, ws in sheets.items():
            ws.append(["Priority", "Source File"] + first_headers)

        for severity, file_name, group, headers, acc_idx, sold_to_count in output_rows:

            ws = sheets[severity]

            ws.append([f"--- {file_name} ---"] + [""] * (len(headers) + 1))

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

        st.success("✅ Audit complete!")

        # ✅ THIS WAS YOUR MISSING PIECE — NOW GUARANTEED VISIBLE
        st.download_button(
            label="📥 Download Multiple Sold To Duplicates.xlsx",
            data=buffer,
            file_name="Multiple Sold To Duplicates.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    # ---------- SUMMARY ----------
    st.subheader("📋 Summary")
    for name, total, flagged in summary:
        st.write(f"**{name}** → Groups Reviewed: {total}, Flagged: {flagged}")
