import streamlit as st
import os
import invoice_generator
import io
from PyPDF2 import PdfMerger
import logging

#path of current directory
cwd = os.getcwd()
#path for output pdf
output_pdf = os.path.join(cwd,'Merged_invoice.pdf')

#logfile path 
logfile = os.path.join(cwd,'invoice_data.log')
logging.basicConfig(filename=logfile,format="%(asctime)s %(levelname)s %(message)s",filemode='w')
logger = logging.getLogger()#creating logger object

#File upload function
def upload_file():
    uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")
    return uploaded_file

#File download function
def download_pdf(List_of_path,output_pdf):
    merger = PdfMerger()
    try:
        for path in List_of_path:
            if os.path.exists(path):
                merger.append(path)
                
        merger.write(output_pdf)
        merger.close()
        with open(output_pdf, "rb") as f:  # Open in binary read mode ("rb")
            pdf_bytes = f.read()
        filename = os.path.basename(output_pdf)
        st.download_button(
            label=filename,
            data = pdf_bytes,
            file_name = filename,
            mime='application/pdf',
            )
    except Exception as e:
        logger.error(str(e))

if __name__=='__main__':
    excel_file_path=upload_file()
    if excel_file_path:
        df=invoice_generator.read_excel(excel_file_path)
        List_of_path=invoice_generator.generate_pdf(df)
        download_pdf(List_of_path,output_pdf)
