import streamlit as st
import pandas as pd
from io import BytesIO

# --------------------------------------------------
# PAGE CONFIG
# --------------------------------------------------

st.set_page_config(
    page_title="Atlas Copco Excel Merger",
    layout="wide"
)

st.title("Atlas Copco Excel Merger")

st.write(
    """
    Upload one or more Registered Product Excel reports.
    The app will:

    • Automatically locate the header row
    • Remove report metadata
    • Merge all uploads
    • Remove duplicate IDs (optional)
    • Generate one clean Excel file
    """
)

# --------------------------------------------------
# FUNCTIONS
# --------------------------------------------------

def load_excel(uploaded_file):
    """
    Load workbook into memory safely.
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
    Find row that contains the column header 'ID'
    and use that row as the header.
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

    cleaned = df.iloc[header_row + 1:].copy()

    cleaned.columns = headers

    cleaned = cleaned.reset_index(drop=True)

    cleaned = cleaned.dropna(how="all")

    return cleaned


# --------------------------------------------------
# UI
# --------------------------------------------------

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
    "Preview First 100 Rows",
    value=True
)

# --------------------------------------------------
# PROCESS
# --------------------------------------------------

if st.button("Process Files"):

    if not uploaded_files:
        st.warning("Please upload at least one file.")

    else:

        merged_frames = []

        progress = st.progress(0)

        files_processed = 0

        for index, uploaded_file in enumerate(uploaded_files):

            try:

                workbook = load_excel(uploaded_file)

                for sheet_name, df in workbook.items():

                    try:

                        cleaned_df = clean_sheet(df)

                        if cleaned_df is None:
                            continue

                        cleaned_df["Source File"] = uploaded_file.name

                        merged_frames.append(cleaned_df)

                        files_processed += 1

                    except Exception as e:

                        st.warning(
                            f"Sheet '{sheet_name}' in "
                            f"{uploaded_file.name} skipped: {e}"
                        )

            except Exception as e:

                st.error(
                    f"Unable to read "
                    f"{uploaded_file.name}: {e}"
                )

            progress.progress(
                (index + 1) / len(uploaded_files)
            )

        if len(merged_frames) == 0:

            st.error(
                "No valid data was found in any uploaded file."
            )

        else:

            final_df = pd.concat(
                merged_frames,
                ignore_index=True
            )

            if remove_duplicates:

                if "ID" in final_df.columns:

                    before_count = len(final_df)

                    final_df = final_df.drop_duplicates(
                        subset=["ID"]
                    )

                    duplicates_removed = (
                        before_count - len(final_df)
                    )

                    st.success(
                        f"Removed {duplicates_removed:,} duplicate rows."
                    )

            st.success(
                f"Successfully merged "
                f"{len(uploaded_files)} files."
            )

            st.success(
                f"Final row count: {len(final_df):,}"
            )

            if show_preview:

                st.subheader("Preview")

                st.dataframe(
                    final_df.head(100),
                    use_container_width=True
                )

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

            st.download_button(
                label="📥 Download Merged Excel File",
                data=output,
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
