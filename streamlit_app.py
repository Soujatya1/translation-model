import streamlit as st
from docx import Document
from deep_translator import GoogleTranslator

def translate_doc(doc, destination='hi'):
    """
    Translate a Word document while ensuring all text is properly replaced and appears in the output document.
    
    :param doc: Word doc object (from Document class)
    :param destination: Target language (default is Hindi 'hi')
    """
    translator = GoogleTranslator(source='auto', target=destination)  # Ensure correct source detection

    # Translate paragraphs
    for p in doc.paragraphs:
        if p.text.strip():  # Check if the paragraph is not empty
            try:
                translated_text = translator.translate(p.text.strip()) or p.text  # Fallback to original text
                p.text = translated_text  # Replace entire paragraph text
            except Exception as e:
                print(f"Error translating paragraph: {e}")

    # Translate table cells
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if cell.text.strip():  # Check if the cell is not empty
                    try:
                        full_text = "\n".join([para.text.strip() for para in cell.paragraphs])  # Collect all text
                        translated_text = translator.translate(full_text) or full_text  # Fallback to original text
                        
                        # Clear existing content properly
                        cell.text = translated_text  
                    except Exception as e:
                        print(f"Error translating cell text: {e}")

    return doc

def main():
    st.title("Word Document Translator")

    # Upload the document
    uploaded_file = st.file_uploader("Upload a Word Document", type=["docx"])
    
    if uploaded_file:
        # Load the document
        doc = Document(uploaded_file)

        # Dropdown for language selection
        language_options = {
            "Bengali": "bn", "Hindi": "hi", "Odia": "or", "Punjabi": "pa", 
            "Tamil": "ta", "Telegu": "te", "Gujarati": "gu", "Malayalam": "ml"
        }
        target_language = st.selectbox("Select Target Language", options=list(language_options.keys()))
        language_code = language_options[target_language]

        # Translate button
        if st.button("Translate Document"):
            with st.spinner('Translating...'):
                translated_doc = translate_doc(doc, language_code)

                # Save the translated document
                with open("translated_document.docx", "wb") as f:
                    translated_doc.save(f)

                # Provide the download button
                with open("translated_document.docx", "rb") as f:
                    st.download_button(
                        label="Download Translated Document",
                        data=f,
                        file_name="translated_document.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                st.success("Translation complete!")

if __name__ == '__main__':
    main()
