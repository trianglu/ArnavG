
import streamlit as st
import pandas as pd
from io import BytesIO
st.set_page_config(page_title='Excel Merge & Trim Tool', layout='wide')
st.title('Excel Merge & Trim Tool')
st.write('Upload multiple Excel files, remove rows and columns, and download the processed workbook.')
uploaded_files = st.file_uploader(
    'Upload Excel Files',
    type=['xlsx', 'xls'],
    accept_multiple_files=True
)
rows_to_remove = st.number_input('Rows to Remove', min_value=0, value=0)
cols_to_remove = st.number_input('Columns to Remove', min_value=0, value=0)
merge_mode = st.radio(
    'Output Mode',
    ['Combine All Sheets Into One Sheet', 'Keep Sheets Separate']
)
if st.button('Process Files'):
    if not uploaded_files:
        st.warning('Please upload at least one Excel file.')
    else:
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            combined_data = []
            sheet_counter = 1
            for uploaded_file in uploaded_files:
                workbook = pd.read_excel(
                    uploaded_file,
                    sheet_name=None,
                    header=None
                )
                for sheet_name, df in workbook.items():
                    trimmed_df = df.iloc[
                        rows_to_remove:,
                        cols_to_remove:
                    ].reset_index(drop=True)

                    if merge_mode == 'Combine All Sheets Into One Sheet':
                        combined_data.append(trimmed_df)
                    else:
                        safe_sheet_name = f'Sheet_{sheet_counter}'
                        trimmed_df.to_excel(
                            writer,
                            sheet_name=safe_sheet_name,
                            index=False,
                            header=False
                        )
                        sheet_counter += 1
            if merge_mode == 'Combine All Sheets Into One Sheet':
                final_df = pd.concat(combined_data, ignore_index=True)
                final_df.to_excel(
                    writer,
                    sheet_name='Merged_Data',
                    index=False,
                    header=False
                )

        output.seek(0)

        st.success('Processing complete!')


        st.download_button(
            label='Download Processed Excel File',
            data=output,
            file_name='processed_excel.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

st.markdown('### GitHub Requirements')
st.code('''
streamlit
pandas
openpyxl
xlrd
''', language='text')
