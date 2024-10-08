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

st.set_page_config(page_title="Excel/CSV File Splitter", page_icon="📊", layout="wide")

# Print the Python interpreter that Streamlit is running on
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

def create_email_drafts(df, column_name, email_df, file_extension):
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

            # Get the corresponding row from email_df
            email_row = email_df[email_df.iloc[:, 0] == value]
            if not email_row.empty:
                # Get all non-null email addresses for this value
                email_list = email_row.iloc[0].dropna().tolist()[1:]  # Exclude the first column (Supplier Name)
            else:
                email_list = []

            if email_list:
                msg = MIMEMultipart()
                msg['Subject'] = f'Split File - {value}'
                msg['To'] = '; '.join(email_list)
                msg['From'] = 'your_email@example.com'
                msg.attach(MIMEText(f'Please find attached the {file_extension.upper()} file for {value}.'))

                part = MIMEBase('application', 'octet-stream')
                part.set_payload(file_buffer.getvalue())
                encoders.encode_base64(part)
                part.add_header('Content-Disposition', f'attachment; filename="{value}.{file_extension}"')
                msg.attach(part)

                eml_file = BytesIO()
                generator = BytesGenerator(eml_file)
                generator.flatten(msg)
                eml_file.seek(0)
                zipf.writestr(f"{value}.eml", eml_file.getvalue())
            else:
                st.warning(f"No email addresses found for {value}. Skipping draft creation.")

    zip_buffer.seek(0)
    return zip_buffer

st.title('📊 Excel/CSV File Splitter')

with st.sidebar:
    st.header("Instructions")
    st.write("1. Upload an Excel or CSV file")
    st.write("2. Select the column to split by")
    st.write("3. Choose output format (ZIP or Email Drafts)")
    st.write("4. Process and download result")
    
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

    st.header("Step 3: Choose Output Format")
    output_format = st.radio("Choose output format:", ("ZIP", "Email Drafts"))
    if output_format == "Email Drafts":
        st.write("Please upload a CSV file with email addresses. The file should have the following format:")
        st.code("Supplier Name,Email 1,Email 2,Email 3\nSupplier A,email1@example.com,email2@example.com,email3@example.com")
        
        email_file = st.file_uploader("Upload email list (CSV or Excel)", type=["csv", "xlsx"])
        if email_file is not None:
            if email_file.name.endswith('.xlsx'):
                email_df = pd.read_excel(email_file)
            else:
                email_df = pd.read_csv(email_file)
            st.dataframe(email_df.head())
            email_list = email_df.iloc[:, 1:].values.flatten().tolist()
            email_list = [email for email in email_list if isinstance(email, str) and '@' in email]
            st.write(f"Found {len(email_list)} valid email addresses.")
        else:
            email_list = []
            st.error("Please upload a CSV file with email addresses.")

    st.header("Step 4: Process and Download")
    if st.button("Process", key="process_button"):
        with st.spinner("Processing..."):
            if output_format == "ZIP":
                zip_buffer = split_excel_and_zip(df, column_name, file_extension)
                st.success("ZIP file created successfully!")
                st.download_button(
                    label="📥 Download ZIP file",
                    data=zip_buffer,
                    file_name="split_files.zip",
                    mime="application/zip"
                )
            else:  # Email Drafts format
                if email_df is not None:                    
                    zip_buffer = create_email_drafts(df, column_name, email_df, file_extension)
                    st.success("Email draft files created successfully!")
                    st.download_button(
                        label="📥 Download Email Drafts (ZIP)",
                        data=zip_buffer,
                        file_name="email_drafts.zip",
                        mime="application/zip"
                    )
                else:
                    st.error("Please upload a CSV file with email addresses.")

else:
    st.info("Please upload a file in the sidebar to get started.")
    st.header("Welcome to Excel/CSV File Splitter")
    st.write("To begin, please upload an Excel (.xlsx) or CSV file using the file uploader in the sidebar.")
    st.write("Once you've uploaded a file, you'll be able to:")
    st.write("1. Preview your data")
    st.write("2. Select a column to split by")
    st.write("3. Choose your output format (ZIP or Email Drafts)")
    st.write("4. Process and download your split files or email drafts")