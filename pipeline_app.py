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
    - Uses fuzzy matching + blocking
    - Outputs:
        - Highlighted main sheet
        - Duplicate groups
        - Merge instructions
        - Dashboard
    """)

# ------------------ HELPERS ------------------

def safe_read_excel(file):
    try:
        # Attempt normal read first
        df = pd.read_excel(file, skiprows=4, engine="openpyxl")

    except Exception:

        st.warning("⚠️ File appears corrupted. Using recovery mode...")

        try:
            file.seek(0)

            # Read raw data safely (no structure assumptions)
            df = pd.read_excel(file, header=None, engine="openpyxl")

            # Drop completely empty rows
            df = df.dropna(how="all")

            # Drop empty columns
            df = df.dropna(axis=1, how="all")

            # Use first valid row as header
            df.columns = df.iloc[0].astype(str)
            df = df[1:].reset_index(drop=True)

        except Exception:
            st.error("❌ This file cannot be read due to severe formatting issues.\n\n✅ Fix: Open the file in Excel → Copy all → Paste into a new file → Upload again.")
            st.stop()

    # Final cleanup
    df.columns = [str(col).strip() for col in df.columns]

    return df


def find_column(df, possible_names):
    for col in df.columns:
        if any(name.lower() in col.lower() for name in possible_names):
            return col
    return None


def normalize(s):
    if pd.isna(s):
        return ""
    s = str(s).lower()
    s = re.sub(r'[^a-z0-9 ]', ' ', s)
    s = re.sub(r'\s+', ' ', s).strip()
    return s

# ------------------ PIPELINE ------------------

def run_pipeline(df, name_col, address_col):
    df = df.reset_index(drop=True)

    df["norm_name"] = df[name_col].apply(normalize)
    df["norm_address"] = df[address_col].apply(normalize)

    # Blocking (speed optimization)
    df["block"] = df["norm_name"].str[:3] + "_" + df["norm_address"].str[:5]

    df["group_id"] = None
    df["match_score"] = None

    group_counter = 1

    for block, subset in df.groupby("block"):
        idxs = subset.index.tolist()

        for i_idx in range(len(idxs)):
            i = idxs[i_idx]

            if pd.notna(df.loc[i, "group_id"]):
                continue

            group = [i]

            for j_idx in range(i_idx + 1, len(idxs)):
                j = idxs[j_idx]

                if pd.notna(df.loc[j, "group_id"]):
                    continue

                name_score = fuzz.token_sort_ratio(
                    df.loc[i, "norm_name"],
                    df.loc[j, "norm_name"]
                )

                addr_score = fuzz.token_sort_ratio(
                    df.loc[i, "norm_address"],
                    df.loc[j, "norm_address"]
                )

                final_score = 0.6 * name_score + 0.4 * addr_score

                if final_score >= 92:
                    group.append(j)
                    df.loc[j, "match_score"] = final_score

            if len(group) > 1:
                for idx in group:
                    df.loc[idx, "group_id"] = group_counter
                    if pd.isna(df.loc[idx, "match_score"]):
                        df.loc[idx, "match_score"] = 100

                group_counter += 1

    # Duplicate count
    df["dup_count"] = df["group_id"].map(df["group_id"].value_counts())

    # Actions
    actions = []
    for gid, group in df.groupby("group_id", dropna=False):
        if pd.isna(gid):
            for idx in group.index:
                actions.append((idx, "UNIQUE"))
        else:
            master_idx = group[name_col].astype(str).str.len().idxmax()

            for idx in group.index:
                if idx == master_idx:
                    actions.append((idx, "KEEP (Master)"))
                else:
                    actions.append((idx, f"MERGE into {master_idx}"))

    df["Action"] = pd.Series(dict(actions))

    # Sort
    dups = df[df["group_id"].notna()].sort_values(
        ["dup_count", "group_id"], ascending=[False, True]
    )
    singles = df[df["group_id"].isna()]

    return pd.concat([dups, singles])


# ------------------ EXCEL OUTPUT ------------------

def build_file(df):
    output = BytesIO()
    wb = Workbook()

    # -------- MAIN SHEET --------
    ws_main = wb.active
    ws_main.title = "Main"

    ws_main.append(list(df.columns))
    for row in df.itertuples(index=False):
        ws_main.append(list(row))

    # -------- COLOR HIGHLIGHTING --------
    colors = ["FFFF99","99FF99","99CCFF","FF9999","CC99FF"]
    color_map = {}
    idx = 0

    for gid in df["group_id"].dropna().unique():
        color_map[gid] = colors[idx % len(colors)]
        idx += 1

    for i, row in enumerate(df.itertuples(index=False), start=2):
        if pd.notna(row.group_id):
            fill = PatternFill(start_color=color_map[row.group_id],
                               end_color=color_map[row.group_id],
                               fill_type="solid")
            for col in range(1, len(df.columns)+1):
                ws_main.cell(row=i, column=col).fill = fill

    # -------- DUPLICATE GROUPS SHEET --------
    ws_groups = wb.create_sheet("Duplicate Groups")

    for gid, group in df[df["group_id"].notna()].groupby("group_id"):
        ws_groups.append([f"==== GROUP {int(gid)} ===="])
        ws_groups.append(list(df.columns))

        for r in group.itertuples(index=False):
            ws_groups.append(list(r))

        ws_groups.append([])

    # -------- MERGE INSTRUCTIONS SHEET --------
    ws_merge = wb.create_sheet("Merge Instructions")
    ws_merge.append(["group_id", "Action"])

    for r in df[df["Action"].str.contains("MERGE", na=False)].itertuples(index=False):
        ws_merge.append([r.group_id, r.Action])

    # -------- SUMMARY DASHBOARD --------
    ws_summary = wb.create_sheet("Summary Dashboard")
    ws_summary.append(["Group ID", "Duplicate Count"])

    summary = (
        df[df["group_id"].notna()]
        .groupby("group_id")
        .size()
        .sort_values(ascending=False)
    )

    for gid, count in summary.items():
        ws_summary.append([gid, count])

    wb.save(output)
    return output


# ------------------ RUN APP ------------------

uploaded_file = st.file_uploader("📂 Upload Excel file", type=["xlsx"])

if uploaded_file:

    st.success("✅ File uploaded")

    df_raw = safe_read_excel(uploaded_file)

    name_col = find_column(df_raw, ["name"])
    address_col = find_column(df_raw, ["address"])
    state_col = find_column(df_raw, ["state"])

    if not name_col or not address_col:
        st.error("❌ Could not detect Name or Address columns.")
        st.stop()

    st.success(f"✅ Using columns: {name_col}, {address_col}")
    st.dataframe(df_raw.head())

    if st.button("🚀 Run Pipeline"):

        with st.spinner("Processing..."):
            df_clean = df_raw.iloc[:, 1:].copy()
            
            # Detect columns AFTER removing junk column
            name_col = find_column(df_clean, ["name"])
            address_col = find_column(df_clean, ["address"])
            state_col = find_column(df_clean, ["state"])
            
            result = run_pipeline(df_clean, name_col, address_col)
            excel_file = build_file(result)

        st.success("✅ Pipeline Complete!")

        # Dynamic filename
        if state_col:
            state_name = result[state_col].dropna().astype(str).str.title().mode()[0]
        else:
            state_name = "output"

        filename = f"{state_name.replace(' ','_')}_final_processed.xlsx"

        st.download_button(
            "⬇ Download Output",
            excel_file.getvalue(),
            file_name=filename
        )
