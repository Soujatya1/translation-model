import streamlit as st
from docx import Document
from deep_translator import GoogleTranslator

def translate_doc(doc, destination='hi'):
    """
    Translate a Word document and save the result with the same format, showing only the translated text.
    
    :param doc: Word doc object (from Document class)
    :param destination: Target language (default is Hindi 'hi')
    """
    translator = GoogleTranslator(source='auto', target=destination)

    # Translate paragraphs
    for p in doc.paragraphs:
        if p.text.strip():  # Check if the paragraph is not empty
            for run in p.runs:  # Iterate over runs in the paragraph to preserve formatting
                if run.text.strip():  # Translate only non-empty runs
                    try:
                        translated_text = translator.translate(run.text)
                        run.text = translated_text  # Replace text while preserving formatting
                    except Exception as e:
                        print(f"Error translating paragraph: {e}")
                        continue  # Skip this run if there's an error

    # Translate table cells
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if cell.text.strip():  # Check if the cell is not empty
                    for run in cell.paragraphs[0].runs:  # Iterate over runs in the cell
                        if run.text.strip():  # Translate only non-empty runs
                            try:
                                translated_text = translator.translate(run.text)
                                run.text = translated_text  # Replace text while preserving formatting
                            except Exception as e:
                                print(f"Error translating cell text: {e}")
                                continue  # Skip this run if there's an error

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
