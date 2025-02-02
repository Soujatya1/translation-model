import streamlit as st
from docx import Document
from deep_translator import GoogleTranslator

def translate_text(text, destination='hi'):
    """
    Translate text using Google Translator.
    :param text: Text to be translated
    :param destination: Target language code (e.g., 'hi' for Hindi)
    :return: Translated text
    """
    try:
        translator = GoogleTranslator(source='auto', target=destination)
        return translator.translate(text) if text.strip() else text
    except Exception as e:
        print(f"Error translating text: {e}")
        return text  # Return original text if translation fails

def translate_doc(doc, destination='hi'):
    """
    Translate a Word document while preserving format and alignment.
    :param doc: Word doc object (from `Document` class)
    :param destination: Target language (default is Hindi 'hi')
    :return: Translated Word doc
    """
    # Translate paragraphs as a whole to avoid breaking formatting
    for p in doc.paragraphs:
        if p.text.strip():
            translated_text = translate_text(p.text, destination)
            p.clear()  # Remove existing runs to avoid formatting issues
            p.add_run(translated_text)  # Add translated text while preserving alignment

    # Translate table cells
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if cell.text.strip():
                    translated_text = translate_text(cell.text, destination)
                    cell.paragraphs[0].clear()
                    cell.paragraphs[0].add_run(translated_text)

    return doc


def main():
    st.title("Word Document Translator")

    uploaded_file = st.file_uploader("Upload a Word Document", type=["docx"])
    
    if uploaded_file:
        doc = Document(uploaded_file)

        language_options = {
            "Bengali": "bn", "Hindi": "hi", "Odia": "or", "Punjabi": "pa", 
            "Tamil": "ta", "Telugu": "te", "Gujarati": "gu", "Malayalam": "ml"
        }
        target_language = st.selectbox("Select Target Language", options=list(language_options.keys()))
        language_code = language_options[target_language]

        if st.button("Translate Document"):
            with st.spinner('Translating...'):
                translated_doc = translate_doc(doc, language_code)

                # Save the translated document in memory
                from io import BytesIO
                translated_io = BytesIO()
                translated_doc.save(translated_io)
                translated_io.seek(0)

                # Provide the download button
                st.download_button(
                    label="Download Translated Document",
                    data=translated_io,
                    file_name="translated_document.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
                st.success("Translation complete!")

if __name__ == '__main__':
    main()
