import streamlit as st
import pandas as pd
import re
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import PatternFill
from rapidfuzz import fuzz

# ------------------ UI CONFIG ------------------
st.set_page_config(page_title="Dedup Pipeline Tool", layout="wide")

st.title("🔍 Excel Deduplication Tool")
st.markdown("### Upload → Process → Download (CA-style output)")

with st.expander("ℹ️ How this works"):
    st.write("""
    - Removes metadata rows + first column
    - Normalizes names & addresses
    - Uses fuzzy matching to group duplicates
    - Outputs:
        - Highlighted main sheet
        - Duplicate groups
        - Merge instructions
        - Dashboard
    """)

uploaded_file = st.file_uploader("📂 Upload Excel file", type=["xlsx"])

# ------------------ NORMALIZATION ------------------
def normalize(s):
    if pd.isna(s):
        return ""
    s = str(s).lower()
    s = re.sub(r'[^a-z0-9 ]', ' ', s)
    s = re.sub(r'\s+', ' ', s).strip()
    return s

# ------------------ FUZZY MATCH ------------------
def fuzzy_match(row1, row2):
    name_score = fuzz.token_sort_ratio(row1["norm_name"], row2["norm_name"])
    addr_score = fuzz.token_sort_ratio(row1["norm_address"], row2["norm_address"])

    return name_score >= 90 and addr_score >= 90

# ------------------ PIPELINE ------------------
def run_pipeline(df):
    # Remove first column
    df = df.iloc[:, 1:]

    # Normalize
    df["norm_name"] = df["Name"].apply(normalize)
    df["norm_address"] = df["Address"].apply(normalize)

    df["group_id"] = None
    group_counter = 1

    used = set()

    for i in range(len(df)):
        if i in used:
            continue

        current_group = [i]
        used.add(i)

        for j in range(i + 1, len(df)):
            if j in used:
                continue

            if fuzzy_match(df.iloc[i], df.iloc[j]):
                current_group.append(j)
                used.add(j)

        if len(current_group) > 1:
            for idx in current_group:
                df.loc[idx, "group_id"] = group_counter
            group_counter += 1

    df["dup_count"] = df["group_id"].map(df["group_id"].value_counts())

    # Actions
    def assign(group):
        if pd.isna(group["group_id"].iloc[0]):
            return ["UNIQUE"] * len(group)

        idxs = list(group.index)
        return ["KEEP (Master)"] + [f"MERGE into {idxs[0]}"] * (len(group) - 1)

    df["Action"] = df.groupby("group_id", dropna=False).apply(
        lambda g: pd.Series(assign(g), index=g.index)
    ).reset_index(level=0, drop=True)

    # Sorting
    dups = df[df["group_id"].notna()].sort_values(["dup_count"], ascending=False)
    singles = df[df["group_id"].isna()]

    return pd.concat([dups, singles])

# ------------------ EXCEL OUTPUT ------------------
def build_file(df):
    output = BytesIO()
    wb = Workbook()

    ws_main = wb.active
    ws_main.title = "Main"

    ws_main.append(list(df.columns))
    for row in df.itertuples(index=False):
        ws_main.append(list(row))

    # Colors
    colors = ["FFFF99","99FF99","99CCFF","FF9999","CC99FF"]
    color_map = {}
    idx = 0

    for gid in df["group_id"].dropna().unique():
        color_map[gid] = colors[idx % len(colors)]
        idx += 1

    for i, row in enumerate(df.itertuples(index=False), start=2):
        gid = row.group_id
        if pd.notna(gid):
            fill = PatternFill(start_color=color_map[gid],
                               end_color=color_map[gid],
                               fill_type="solid")
            for c in range(1, len(df.columns)+1):
                ws_main.cell(row=i, column=c).fill = fill

    # Duplicate Groups
    ws_groups = wb.create_sheet("Duplicate Groups")

    for gid, group in df[df["group_id"].notna()].groupby("group_id"):
        ws_groups.append([f"==== GROUP {int(gid)} ===="])
        ws_groups.append(list(df.columns))
        for r in group.itertuples(index=False):
            ws_groups.append(list(r))
        ws_groups.append([])

    # Merge Sheet
    ws_merge = wb.create_sheet("Merge Instructions")
    ws_merge.append(["group_id","Action"])

    for r in df[df["Action"].str.contains("MERGE", na=False)].itertuples(index=False):
        ws_merge.append([r.group_id, r.Action])

    # Summary
    ws_summary = wb.create_sheet("Summary Dashboard")
    ws_summary.append(["Group ID","Count"])

    summary = df[df["group_id"].notna()].groupby("group_id").size().sort_values(ascending=False)

    for gid, count in summary.items():
        ws_summary.append([gid, count])

    wb.save(output)
    return output

# ------------------ RUN ------------------
if uploaded_file:
    st.success("✅ File uploaded")

    df_raw = pd.read_excel(uploaded_file, skiprows=4)

    st.write("Preview:")
    st.dataframe(df_raw.head())

    if st.button("🚀 Run Pipeline"):
        with st.spinner("Processing..."):
            result = run_pipeline(df_raw)

            excel_file = build_file(result)

        st.success("✅ Pipeline Complete!")

        st.download_button(
            "⬇ Download Output",
            excel_file.getvalue(),
            file_name="pipeline_output.xlsx"
        )
