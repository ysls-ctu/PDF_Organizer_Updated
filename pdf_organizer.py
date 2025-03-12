import streamlit as st
import os
import re
import pandas as pd
import PyPDF2
import zipfile
from collections import defaultdict

# Page Configuration
st.set_page_config(page_title="PDF Organizer", page_icon="📄", layout="centered")

if "page" not in st.session_state:
    st.session_state.page = "home"

def navigate(page):
    st.session_state.page = page
    st.rerun()

button_style = """
    <style>
        .stButton>button {
            width: 100% !important;  
            height: 70px !important;  
            background-color: #ffac2c !important;  /* Kwik color */
            color: black !important;  
            border-radius: 10px !important; 
            border: none !important;
            transition: 0.3s;
        }
        .stButton>button:hover {
            background-color: #ffb84a !important; 
        }
    </style>
"""
st.markdown(button_style, unsafe_allow_html=True)

# Load SKU Mapping from Excel
def load_sku_mapping(excel_file):
    df = pd.read_excel(excel_file, skiprows=8)
    return dict(zip(df.iloc[:, 1], df.iloc[:, 2]))

# Extract Text from PDF Page
def extract_text_from_page(pdf_reader, page_num):
    return pdf_reader.pages[page_num].extract_text()

# Extract SKU from Text
def get_sku_from_text(text):
    match = re.search(r'(\w{2,3}-\w{4,5}-\w{3,4})', text)
    return match.group(1) if match else None

# Split and Group PDF by Model Number
def split_and_group_pdf(pdf_file, sku_mapping):
    pdf_reader = PyPDF2.PdfReader(pdf_file)
    pdf_writer_dict = defaultdict(PyPDF2.PdfWriter)
    total_pages = len(pdf_reader.pages)
    
    for i in range(0, total_pages, 2):
        if i + 1 >= total_pages:
            continue
        
        text = extract_text_from_page(pdf_reader, i)
        sku = get_sku_from_text(text)
        model_number = sku_mapping.get(sku, sku) if sku else "Unknown"
        pdf_writer = pdf_writer_dict[model_number]
        pdf_writer.add_page(pdf_reader.pages[i])
        pdf_writer.add_page(pdf_reader.pages[i + 1])
    
    return pdf_writer_dict

# Save PDFs to Directory
def save_pdfs(pdf_writer_dict, output_dir):
    os.makedirs(output_dir, exist_ok=True)
    output_files = []
    
    for model_number, writer in pdf_writer_dict.items():
        output_pdf_path = os.path.join(output_dir, f"{model_number}.pdf")
        with open(output_pdf_path, "wb") as output_pdf:
            writer.write(output_pdf)
        output_files.append(output_pdf_path)
    
    return output_files

# Zip Processed PDFs
def zip_files(files, zip_name):
    with zipfile.ZipFile(zip_name, 'w') as zipf:
        for file in files:
            zipf.write(file, os.path.basename(file))
    return zip_name

# Merge Multiple PDFs
def merge_pdfs(pdf_files):
    pdf_writer = PyPDF2.PdfWriter()
    for pdf_file in pdf_files:
        pdf_reader = PyPDF2.PdfReader(pdf_file)
        for page in pdf_reader.pages:
            pdf_writer.add_page(page)
    merged_pdf_path = "merged.pdf"
    with open(merged_pdf_path, "wb") as output_pdf:
        pdf_writer.write(output_pdf)
    return merged_pdf_path

# Home Page
if st.session_state.page == "home":
    st.title("PDF Organizer")
    st.write("Split and organize labels by model number for easy processing. Made for KwikSafety Logistics Department.")
    st.markdown("### Select an option to proceed:")
    col1, col2 = st.columns(2)
    with col1:
        if st.button("One PDF"):
            navigate("split_1pdf")
    with col2:
        if st.button("Multiple PDFs"):
            navigate("split_mpdf")

# Single PDF Processing
elif st.session_state.page == "split_1pdf":
    st.title("Split One PDF")
    st.markdown("Upload an Excel file and one PDF file to split and group labels by Model Number. <br>", unsafe_allow_html=True)
    if st.button("Back to Home"):
        navigate("home")
    
    excel_file = st.file_uploader("Upload Excel File", type=["xlsx"])
    pdf_file = st.file_uploader("Upload PDF File", type=["pdf"])
    
    if excel_file and pdf_file:
        with st.spinner("Processing..."):
            sku_mapping = load_sku_mapping(excel_file)
            pdf_writer_dict = split_and_group_pdf(pdf_file, sku_mapping)
            output_files = save_pdfs(pdf_writer_dict, "output_pdfs")
            zip_name = zip_files(output_files, "Processed_PDFs.zip")
        with open(zip_name, "rb") as f:
            st.download_button("Download Processed PDFs", f, file_name="Processed_PDFs.zip", mime="application/zip")

# Multiple PDF Processing
elif st.session_state.page == "split_mpdf":
    st.title("Split Multiple PDFs")
    st.markdown("Upload an Excel file and multiple PDF files to split and group labels by Model Number. <br>", unsafe_allow_html=True)
    if st.button("Back to Home"):
        navigate("home")
    
    excel_file = st.file_uploader("Upload Excel File", type=["xlsx"])
    pdf_files = st.file_uploader("Upload Multiple PDFs", type=["pdf"], accept_multiple_files=True)
    
    if excel_file and pdf_files:
        with st.spinner("Processing..."):
            sku_mapping = load_sku_mapping(excel_file)
            merged_pdf = merge_pdfs(pdf_files)
            pdf_writer_dict = split_and_group_pdf(merged_pdf, sku_mapping)
            output_files = save_pdfs(pdf_writer_dict, "output_pdfs")
            zip_name = zip_files(output_files, "Processed_PDFs.zip")
        with open(zip_name, "rb") as f:
            st.download_button("Download Processed PDFs", f, file_name="Processed_PDFs.zip", mime="application/zip")

st.markdown(
    """
    <style>
        .footer {
            position: fixed;
            left: 0;
            bottom: 0;
            width: 100%;
            text-align: center;
            padding: 5px;
            padding-top: 20px;
            font-size: 14px;
            color: white;
            background-color: #333;
        }
        a {
            text-decoration: none !important;
            color: white !important; 
        }
    </style>
    <div class="footer">
        <p>Created by <b><a href="https://github.com/ysls-ctu">YSLS</a></b></p>
    </div>
    """,
    unsafe_allow_html=True
)
