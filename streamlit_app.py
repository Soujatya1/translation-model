import streamlit as st
from docx import Document
from deep_translator import GoogleTranslator

def translate_doc(doc, destination='hi'):
    """
    Translate a Word document and save the result with the same format, ensuring all text is translated properly.
    
    :param doc: Word doc object (from Document class)
    :param destination: Target language (default is Hindi 'hi')
    """
    translator = GoogleTranslator(source='en', target=destination)  # Explicitly set source language

    # Translate paragraphs
    for p in doc.paragraphs:
        if p.text.strip():  # Check if the paragraph is not empty
            try:
                translated_text = translator.translate(p.text.strip())  # Translate whole paragraph
                p.clear()  # Clear paragraph to retain formatting
                p.add_run(translated_text)  # Add translated text back
            except Exception as e:
                print(f"Error translating paragraph: {e}")
                continue 

    # Translate table cells
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if cell.text.strip():  # Check if the cell is not empty
                    try:
                        full_text = "\n".join([para.text.strip() for para in cell.paragraphs])  # Get full text
                        translated_text = translator.translate(full_text)  # Translate entire cell
                        
                        cell.clear()  # Clear cell
                        cell.text = translated_text  # Assign translated text
                    except Exception as e:
                        print(f"Error translating cell text: {e}")
                        continue  

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
