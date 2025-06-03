# Author: Raphaele Guillemot

import streamlit as st
from utils import convert_workbook_md, clean_output_dir

OUTPUT_DIR = "result"
OUTPUT_FILE_NAME = "result.zip"

st.title("➡️ Excel to Markdown")
st.markdown(
    """ A simple tool to turn Excel workbooks into Markdown Tables.

    Please be assured that no data will be stored or sent anywhere. The conversion happens entirely in your browser.

    **Note:** If your Excel file is protected, please remove the protection before uploading it here.
    """
)
st.divider()

uploaded_file = st.file_uploader("Choose an Excel file", type = 'xlsx')
if not uploaded_file:
    clean_output_dir(OUTPUT_DIR)
else:
    convert_workbook_md(OUTPUT_DIR, uploaded_file, OUTPUT_FILE_NAME)