import streamlit as st
from docx import Document
from io import BytesIO

def combine_word_documents(files):
    combined_doc = Document()

    for file in files:
        sub_doc = Document(file)
        for element in sub_doc.element.body:
            combined_doc.element.body.append(element)

    return combined_doc

st.title('Word Document Combiner')

uploaded_files = st.file_uploader("Upload Word files", accept_multiple_files=True, type=['docx'])

if st.button('Combine Documents'):
    if uploaded_files:
        combined_document = combine_word_documents(uploaded_files)

        # To download the combined document
        file_stream = BytesIO()
        combined_document.save(file_stream)
        file_stream.seek(0)

        st.download_button(label="Download Combined Document",
                           data=file_stream,
                           file_name="combined_document.docx",
                           mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    else:
        st.write("Please upload at least one document.")
