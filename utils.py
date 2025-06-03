import pandas as pd
import streamlit as st
import zipfile
import os
import xlrd
import datetime

def load_excel_file(uploaded_file):
    try:
        data = pd.read_excel(uploaded_file, sheet_name=None)
        return data
    except xlrd.biffh.XLRDError as e:
        st.error("It seems like your Excel file is protected. Please remove the protection and try again.")
        return None

def clean_output_dir(output_dir):
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        return
    
    for file in os.listdir(output_dir):
        file_path = os.path.join(output_dir, file)
        try:
            if os.path.isfile(file_path):
                os.unlink(file_path)
            elif os.path.isdir(file_path):
                os.rmdir(file_path)
        except Exception as e:
            print(f"Error deleting file {file_path}: {e}")

def clean_md(markdown):
    return markdown.replace("nan", "").replace("Unnamed: ", "").strip()

def convert_xlsx_md(df):
    markdown = "| " + " | ".join(df.columns) + " |\n"
    markdown += "| " + " | ".join(["---"] * len(df.columns)) + " |\n"

    for _, row in df.iterrows():
        markdown += "| " + " | ".join(str(cell) for cell in row) + " |\n"
    
    return clean_md(markdown)

def save_md_file(path, content):
    with open(path, "w", encoding="utf-8") as f:
        f.write(content)

def zipify(path, output_file_name):
    compression = zipfile.ZIP_DEFLATED  # ZIP_DEFLATED for compression, zipfile.ZIP_STORED to just store the file
    zf = zipfile.ZipFile(output_file_name, mode="w")
    
    try:
        for file_name in os.listdir(path):
            zf.write(os.path.join(path, file_name), file_name, compress_type=compression)
    except FileNotFoundError as e:
        st.error(f"File not found. {e}")
    
    finally:
        zf.close()
        return zf

def download_zip_btn(output_file_name):
    now = datetime.datetime.now()

    with open(output_file_name, "rb") as fp: 
        btn = st.download_button(
            label="Download ZIP",
            data=fp,
            file_name=f"excel_to_markdown_{now.strftime('%Y-%m-%d_%H-%M-%S')}.zip",
            mime="application/zip"
        )

def show_mds(output_dir):
    files = os.listdir(output_dir)
    
    if not files:
        st.write("No Markdown files found.")
        return
    
    for file in files:
        file_path = os.path.join(output_dir, file)
        with open(file_path, "r", encoding="utf-8") as f:
            content = f.read()
            with st.expander(file):
                st.code(content, language="markdown")

def show_results(output_file_name, output_dir):
    st.divider()
    st.markdown("## Complete!\n\nYou may directly copy the contents of individual Markdown files below by clicking the icon at the top-right corner\n\nYou can also download all files here:")
    download_zip_btn(output_file_name)
    show_mds(output_dir)

def convert_workbook_md(output_dir, uploaded_file, output_file_name):
    data = load_excel_file(uploaded_file)
    
    if data is None:
        return
    
    for sheet_name, df in data.items():
        md = convert_xlsx_md(df)
        path = os.path.join(output_dir, f"{sheet_name}.md")
        save_md_file(path, md)
    
    zf = zipify(output_dir, "result.zip")

    if not zf:
        st.error("An error occurred while creating the ZIP file.")
        return
    
    show_results(output_file_name, output_dir)
    