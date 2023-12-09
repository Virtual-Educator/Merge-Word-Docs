import streamlit as st
from docx import Document
from zipfile import ZipFile
import os
import tempfile
from io import BytesIO

def combine_word_documents(docs):
    combined_doc = Document()
    for doc in docs:
        sub_doc = Document(BytesIO(doc))
        for element in sub_doc.element.body:
            combined_doc.element.body.append(element)
    return combined_doc

def process_zip_file(zip_file):
    with ZipFile(zip_file, 'r') as z:
        with tempfile.TemporaryDirectory() as tempdir:
            z.extractall(tempdir)
            word_docs = []
            error_occurred = False

            for folder in os.listdir(tempdir):
                folder_path = os.path.join(tempdir, folder)
                if os.path.isdir(folder_path):
                    docs_in_folder = [file for file in os.listdir(folder_path) if file.endswith('.docx')]
                    
                    if len(docs_in_folder) > 1:
                        st.error(f"More than one Word document found in the folder '{folder}'. Only the first document will be processed.")
                        error_occurred = True
                    
                    if docs_in_folder:
                        file_path = os.path.join(folder_path, docs_in_folder[0])
                        with open(file_path, 'rb') as f:
                            word_docs.append(f.read())

            return word_docs, error_occurred

def process_word_files(word_files):
    return [file.getvalue() for file in word_files]

st.title('Word Document Combiner')

upload_choice = st.radio("Choose your upload method", ('Zip File', 'Word Files'))
combined_document = None

if upload_choice == 'Zip File':
    uploaded_file = st.file_uploader("Upload ZIP file", type=['zip'])
    if st.button('Combine Documents from ZIP') and uploaded_file:
        word_docs, error_occurred = process_zip_file(uploaded_file)
        if word_docs and not error_occurred:
            combined_document = combine_word_documents(word_docs)

elif upload_choice == 'Word Files':
    uploaded_files = st.file_uploader("Upload Word files", accept_multiple_files=True, type=['docx'])
    if st.button('Combine Word Documents') and uploaded_files:
        word_docs = process_word_files(uploaded_files)
        combined_document = combine_word_documents(word_docs)

if combined_document:
    export_format = st.selectbox("Select export format", ("Word", "PDF", "Text"))

    if st.button('Export Combined Document'):
        file_stream = BytesIO()
        combined_document.save(file_stream)
        file_stream.seek(0)

        if export_format == "Word":
            st.download_button(label="Download Combined Document",
                               data=file_stream,
                               file_name="combined_document.docx",
                               mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        elif export_format == "PDF":
            # PDF conversion logic here (Windows only with docx2pdf)
            # For cross-platform, use an alternative method
            pass
        elif export_format == "Text":
            text_stream = BytesIO()
            for paragraph in combined_document.paragraphs:
                text_stream.write(paragraph.text.encode('utf-8') + b'\n')
            text_stream.seek(0)
            st.download_button(label="Download Combined Document",
                               data=text_stream,
                               file_name="combined_document.txt",
                               mime="text/plain")
