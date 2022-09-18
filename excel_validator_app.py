# Packages
import streamlit as st
import pandas as pd
import numpy as np
import warnings
import openpyxl as op
from io import BytesIO
from pyxlsb import open_workbook as open_xlsb

warnings.filterwarnings("ignore")

# App
def main():
    st.title('Excel Validator')

    # Get excel files
    if st.checkbox('Select Excel Files', key=1):
        col1, col2 = st.columns(2)

        with col1:
            st.subheader('HAMS File')
            file_1 = st.file_uploader("Select HAMS file", type=['xlsx'])

            if file_1:
                comment_details='''file_details = {
                    "Filename": file_1.name,
                    "FileType": file_1.type,
                    "FileSize": file_1.size}'''

                st.write("##")
                wb = op.load_workbook(file_1)

                # Show Excel file
                #st.subheader("File details:")
                #st.json(file_details, expanded=False)
                st.markdown("----")

                # Select sheet
                sheet_selector = st.selectbox("Select sheet:", wb.sheetnames, key='a')
                df_1 = pd.read_excel(file_1, sheet_selector)
                st.write("##")
                st.markdown(f"### Currently Selected: `{sheet_selector}`")
                st.write(df_1)

                # SUM
                st.write("##")
                #st.subheader('SUM of columns')
                df_1_sum = df_1.sum(numeric_only=True)
                #st.dataframe(df_1_sum)

                # Count
                st.write("##")
                #st.subheader('Count of columns')
                df_1_count = df_1.count(numeric_only=True)
                #st.dataframe(df_1_count)

        with col2:
            st.subheader('DWH File')
            file_2 = st.file_uploader("Select DWH file", type=['xlsx'])

            if file_2:
                comment_details='''file_details = {
                    "Filename": file_2.name,
                    "FileType": file_2.type,
                    "FileSize": file_2.size}'''

                st.write("##")
                wb = op.load_workbook(file_2)

                # Show Excel file
                #st.subheader("File details:")
                #st.json(file_details, expanded=False)
                st.markdown("----")

                # Select sheet
                sheet_selector = st.selectbox("Select sheet:", wb.sheetnames, key='b')
                df_2 = pd.read_excel(file_2, sheet_selector)
                st.write("##")
                st.markdown(f"### Currently Selected: `{sheet_selector}`")
                st.write(df_2)

                # SUM
                st.write("##")
                #st.subheader('SUM of columns')
                df_2_sum = df_2.sum(numeric_only=True)
                #st.dataframe(df_2_sum)

                # Count
                st.write("##")
                #st.subheader('Count of columns')
                df_2_count = df_2.count(numeric_only=True)
                #st.dataframe(df_2_count)

            st.write("##")

    ## Comparison
    # SUM
    if st.checkbox('Compare Excels', key=2):
        sum = 'SUM'
        st.subheader(f"### Validation: `{sum}`")


        # Series to Dataframe
        try:
            df_1_sum = df_1_sum.to_frame()
            df_2_sum = df_2_sum.to_frame()

            # Compare
            df_compare_sum =df_1_sum.compare(df_2_sum, align_axis=0)
            st.write('*self -> HAMS; other -> DWH*')
            st.dataframe(df_compare_sum)

            # Dif between sheets
            st.markdown('#### Diference between sheets')
            df_dif_sum = df_compare_sum.diff()
            st.dataframe(df_dif_sum)

            st.write("##")
        except:
            st.error('Sheets with different formats')

        # Count
        count = 'COUNT'
        st.subheader(f"### Validation: `{count}`")

        # Series to Dataframe
        try:
            df_1_count = df_1_count.to_frame()
            df_2_count = df_2_count.to_frame()

            # Compare
            df_compare_count = df_1_count.compare(df_2_count, align_axis=0)
            st.write('*self -> HAMS; other -> DWH*')
            st.dataframe(df_compare_count)

            # Dif between sheets
            st.markdown('#### Diference between sheets')
            df_dif_count = df_compare_count.diff()
            st.dataframe(df_dif_count)

        except:
            st.error('Sheets with different formats')


        ## Final Validation
        st.write("##")

        # Final Dataframe
        #df_dif_sum = df_dif_sum.rename(index={2: "SUM"})
        #st.dataframe(df_dif_sum)
        #df_dif_count = df_dif_count.rename(index={0: "COUNT"})

        #df_validation = pd.append([df_dif_sum, df_dif_count], axis=1)
        df_validation = df_dif_sum.append(df_dif_count)
        #df_validation.rename(index={0: "SUM", 1: "COUNT"})
        st.dataframe(df_validation)

        st.info('**Validation Status**')
        if df_dif_sum.empty & df_dif_count.empty:
            st.success('**Sheet Validated!**')
        else:
            st.error('**Sheet Not Validated!**')

        st.write("##")
        st.subheader('Download validation results')
        st.download_button(label='📥 Download Excel', data=df_validation.to_csv(), mime='text/csv')

        comment_download_excel='''def to_excel(df):
            output = BytesIO()
            writer = pd.ExcelWriter(output, engine='openpyxl')
            df.to_excel(writer, index=False, sheet_name='Validation')
            workbook = writer.book
            worksheet = writer.sheets['Validation']
            writer.save()
            processed_data = output.getvalue()
            return processed_data

        df_xlsx = to_excel(df_validation)
        st.download_button(label='📥 Download Excel',
                           data=df_xlsx,
                           file_name='df_test.xlsx')'''

    # warning caso formato seja diferente -> ta
    # mensagem sucesso -> ta
    # download file with 2 validation dataframes -> ta

if __name__ == '__main__':
    main()