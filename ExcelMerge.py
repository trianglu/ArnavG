

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
        combined_data = []
        processed_any_sheet = False
        sheet_counter = 1
        try:
            with pd.ExcelWriter(output, engine='openpyxl') as writer:


                for uploaded_file in uploaded_files:
                    try:
                        workbook = pd.read_excel(
                            uploaded_file,
                            sheet_name=None,
                            header=None,
                            engine='openpyxl'
                        )
                    except Exception as e:
                        st.error(f'Unable to read {uploaded_file.name}: {e}')
                        continue


                    for sheet_name, df in workbook.items():
                        try:
                            trimmed_df = df.iloc[
                                rows_to_remove:,
                                cols_to_remove:
                            ].reset_index(drop=True)


                            if trimmed_df.empty:
                                continue


                            processed_any_sheet = True


                            if merge_mode == 'Combine All Sheets Into One Sheet':
                                combined_data.append(trimmed_df)
                            else:
                                trimmed_df.to_excel(
                                    writer,
                                    sheet_name=f'Sheet_{sheet_counter}',
                                    index=False,
                                    header=False
                                )
                                sheet_counter += 1


                        except Exception as e:
                            st.warning(f"Error in sheet '{sheet_name}': {e}")


                if merge_mode == 'Combine All Sheets Into One Sheet':
                    if combined_data:
                        final_df = pd.concat(combined_data, ignore_index=True)
                        final_df.to_excel(
                            writer,
                            sheet_name='Merged_Data',
                            index=False,
                            header=False
                        )
                    else:
                        pd.DataFrame([['No valid data found']]).to_excel(
                            writer,
                            sheet_name='Merged_Data',
                            index=False,
                            header=False
                        )

                elif not processed_any_sheet:
                    pd.DataFrame([['No valid data found']]).to_excel(
                        writer,
                        sheet_name='Output',
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
        except Exception as e:
            st.error(f'Unexpected error: {e}')


st.markdown('### GitHub Requirements')
st.code('''
streamlit
pandas
openpyxl
xlrd
''', language='text')
