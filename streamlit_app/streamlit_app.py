import streamlit as st
import pandas as pd
import zipfile
from io import BytesIO
import base64
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.generator import BytesGenerator
import sys
import smtplib

st.set_page_config(page_title="Excel/CSV File Splitter & Merger", page_icon="📊", layout="wide")

def split_excel_and_zip(df, column_name, file_extension):
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for value, group in df.groupby(column_name):
            file_buffer = BytesIO()
            if file_extension == 'xlsx':
                with pd.ExcelWriter(file_buffer, engine='openpyxl') as writer:
                    group.to_excel(writer, index=False, sheet_name='Sheet1')
            else:  # csv
                group.to_csv(file_buffer, index=False)
            file_buffer.seek(0)
            zipf.writestr(f"{value}.{file_extension}", file_buffer.getvalue())
    zip_buffer.seek(0)
    return zip_buffer

def merge_excel_files(uploaded_files):
    dfs = []
    for file in uploaded_files:
        if file.name.endswith('.xlsx'):
            df = pd.read_excel(file)
        else:  # csv
            df = pd.read_csv(file)
        dfs.append(df)
    merged_df = pd.concat(dfs, ignore_index=True)
    return merged_df

st.title('📊 Excel/CSV File Splitter & Merger')

operation = st.radio("Choose operation:", ("Split Excel/CSV", "Merge Excel/CSV"))

if operation == "Split Excel/CSV":
    st.header("Excel/CSV File Splitter")
    st.write("Upload an Excel or CSV file to split it by a specific column.")
    
    uploaded_file = st.file_uploader("Choose an Excel or CSV file", type=["xlsx", "csv"])

    if uploaded_file is not None:
        file_extension = uploaded_file.name.split('.')[-1].lower()
        
        with st.spinner("Loading data..."):
            if file_extension == 'xlsx':
                df = pd.read_excel(uploaded_file)
            else:  # csv
                df = pd.read_csv(uploaded_file)
        
        st.success("File uploaded successfully!")
        
        st.header("Step 1: Data Preview")
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("Data Preview")
            st.dataframe(df.head())
        
        with col2:
            st.subheader("File Information")
            st.write(f"Rows: {df.shape[0]}")
            st.write(f"Columns: {df.shape[1]}")
            st.write(f"File type: {file_extension.upper()}")

        st.header("Step 2: Select Column to Split By")
        column_name = st.selectbox("Select the column to split by:", df.columns)

        # Display the number of unique values in the selected column
        unique_values_count = df[column_name].nunique()
        st.write(f"Number of unique values in '{column_name}': {unique_values_count}")

        st.header("Step 3: Process and Download")
        if st.button("Process", key="process_button"):
            with st.spinner("Processing..."):
                zip_buffer = split_excel_and_zip(df, column_name, file_extension)
                st.success("ZIP file created successfully!")
                st.download_button(
                    label="📥 Download ZIP file",
                    data=zip_buffer,
                    file_name="split_files.zip",
                    mime="application/zip"
                )

    else:
        st.info("Please upload a file to get started.")
        st.write("Once you've uploaded a file, you'll be able to:")
        st.write("1. Preview your data")
        st.write("2. Select a column to split by")
        st.write("3. Process and download your split files")

else:  # Merge Excel/CSV
    st.header("Excel/CSV File Merger")
    st.write("Upload multiple Excel or CSV files to merge them into one.")
    
    uploaded_files = st.file_uploader("Choose Excel or CSV files", type=["xlsx", "csv"], accept_multiple_files=True)
    
    if uploaded_files:
        st.write(f"Uploaded {len(uploaded_files)} files.")
        if st.button("Merge Files"):
            with st.spinner("Merging files..."):
                merged_df = merge_excel_files(uploaded_files)
            
            st.success("Files merged successfully!")
            st.write("Preview of merged data:")
            st.dataframe(merged_df.head())
            
            # Offer download options
            csv_buffer = BytesIO()
            merged_df.to_csv(csv_buffer, index=False)
            csv_buffer.seek(0)
            
            excel_buffer = BytesIO()
            with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                merged_df.to_excel(writer, index=False, sheet_name='Sheet1')
            excel_buffer.seek(0)
            
            col1, col2 = st.columns(2)
            with col1:
                st.download_button(
                    label="📥 Download Merged CSV",
                    data=csv_buffer,
                    file_name="merged_files.csv",
                    mime="text/csv"
                )
            with col2:
                st.download_button(
                    label="📥 Download Merged Excel",
                    data=excel_buffer,
                    file_name="merged_files.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
    else:
        st.info("Please upload at least two Excel or CSV files to merge.")