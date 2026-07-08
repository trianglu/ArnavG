import streamlit as st
import pandas as pd
from io import BytesIO

# ============================================================
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="Atlas Copco Excel Merger",
    layout="wide"
)

st.title("Atlas Copco Excel Merger")

st.write("""
Upload multiple Registered Product reports.

The app will:

✅ Automatically find the header row

✅ Remove report metadata

✅ Merge all files

✅ Remove duplicate IDs (optional)

✅ Generate a downloadable Excel file
""")

# ============================================================
# FUNCTIONS
# ============================================================

def load_excel(uploaded_file):
    """
    Load workbook safely from uploaded file.
    """
    uploaded_file.seek(0)

    file_bytes = BytesIO(uploaded_file.read())

    workbook = pd.read_excel(
        file_bytes,
        sheet_name=None,
        header=None,
        engine="openpyxl"
    )

    return workbook


def clean_sheet(df):
    """
    Find header row containing 'ID'
    and convert sheet into clean table.
    """

    header_row = None

    for i in range(len(df)):

        values = [
            str(v).strip()
            for v in df.iloc[i].fillna("")
        ]

        if "ID" in values:
            header_row = i
            break

    if header_row is None:
        return None

    headers = df.iloc[header_row]

    cleaned_df = df.iloc[header_row + 1:].copy()

    cleaned_df.columns = headers

    cleaned_df = cleaned_df.reset_index(drop=True)

    cleaned_df = cleaned_df.dropna(how="all")

    return cleaned_df


# ============================================================
# USER OPTIONS
# ============================================================

uploaded_files = st.file_uploader(
    "Upload Excel Files",
    type=["xlsx", "xls"],
    accept_multiple_files=True
)

remove_duplicates = st.checkbox(
    "Remove Duplicate IDs",
    value=True
)

show_preview = st.checkbox(
    "Show Preview",
    value=True
)

# ============================================================
# PROCESS FILES
# ============================================================

if st.button("Process Files"):

    if not uploaded_files:
        st.warning("Please upload at least one file.")

    else:

        merged_frames = []

        failed_files = []

        progress = st.progress(0)

        for idx, uploaded_file in enumerate(uploaded_files):

            try:

                workbook = load_excel(uploaded_file)

                for sheet_name, df in workbook.items():

                    try:

                        cleaned_df = clean_sheet(df)

                        if cleaned_df is None:
                            continue

                        cleaned_df["Source File"] = uploaded_file.name

                        merged_frames.append(cleaned_df)

                    except Exception:
                        pass

            except Exception as e:

                failed_files.append(
                    f"{uploaded_file.name} ({str(e)})"
                )

            progress.progress(
                (idx + 1) / len(uploaded_files)
            )

        if not merged_frames:

            st.error("No valid data found.")

        else:

            # ====================================================
            # MERGE
            # ====================================================

            final_df = pd.concat(
                merged_frames,
                ignore_index=True
            )

            # ====================================================
            # CLEAN COLUMN NAMES
            # ====================================================

            clean_columns = []

            for i, col in enumerate(final_df.columns):

                if pd.isna(col):
                    clean_columns.append(f"Column_{i}")
                else:
                    clean_columns.append(str(col))

            final_df.columns = clean_columns

            final_df = final_df.fillna("")

            # ====================================================
            # REMOVE DUPLICATES
            # ====================================================

            if remove_duplicates:

                if "ID" in final_df.columns:

                    before_count = len(final_df)

                    final_df = final_df.drop_duplicates(
                        subset=["ID"]
                    )

                    removed = before_count - len(final_df)

                    st.success(
                        f"Removed {removed:,} duplicate rows."
                    )

            # ====================================================
            # SUMMARY
            # ====================================================

            st.success(
                f"Successfully merged {len(uploaded_files)} files."
            )

            st.success(
                f"Final row count: {len(final_df):,}"
            )

            if failed_files:

                st.warning(
                    f"{len(failed_files)} file(s) could not be read."
                )

                with st.expander("View Failed Files"):

                    for file in failed_files:
                        st.write(file)

            # ====================================================
            # PREVIEW
            # ====================================================

            if show_preview:

                st.subheader("Preview")

                preview_df = (
                    final_df
                    .head(100)
                    .fillna("")
                    .astype(str)
                )

                st.write(
                    f"Showing first {len(preview_df)} rows "
                    f"of {len(final_df):,} total rows."
                )

                st.dataframe(
                    preview_df,
                    use_container_width=True
                )

            # ====================================================
            # CREATE EXCEL FILE
            # ====================================================

            output = BytesIO()

            with pd.ExcelWriter(
                output,
                engine="openpyxl"
            ) as writer:

                final_df.to_excel(
                    writer,
                    sheet_name="Merged_Data",
                    index=False
                )

            output.seek(0)

            # ====================================================
            # DOWNLOAD BUTTON
            # ====================================================

            st.download_button(
                label="📥 Download Merged Excel File",
                data=output.getvalue(),
                file_name="Merged_Registered_Products.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

st.markdown('### GitHub Requirements')
st.code('''
streamlit
pandas
openpyxl
xlrd
''', language='text')
