import streamlit as st
from docx import Document
from deep_translator import GoogleTranslator
from io import BytesIO

def translate_doc(doc, destination='hi'):
    """
    Translate a Word document while preserving formatting, including bold text.
    
    :param doc: Word document object.
    :param destination: Target language code.
    """
    translator = GoogleTranslator(source='auto', target=destination)

    # Translate paragraphs
    for p in doc.paragraphs:
        if p.text.strip():  # Check if the paragraph is not empty
            try:
                # We will translate each run and preserve the formatting
                for run in p.runs:
                    translated_text = translator.translate(run.text) if run.text.strip() else ""
                    run.text = translated_text  # Update the run's text
                    run.bold = run.bold  # Preserve bold formatting
                    run.italic = run.italic  # Preserve italic formatting
                    run.underline = run.underline  # Preserve underline formatting
            except Exception as e:
                print(f"Error translating paragraph: {e}")
    
    # Translate tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    if para.text.strip():  # Skip empty paragraphs
                        try:
                            for run in para.runs:
                                translated_text = translator.translate(run.text) if run.text.strip() else ""
                                run.text = translated_text  # Update the run's text
                                run.bold = run.bold  # Preserve bold formatting
                                run.italic = run.italic  # Preserve italic formatting
                                run.underline = run.underline  # Preserve underline formatting
                        except Exception as e:
                            print(f"Error translating cell text: {e}")

    return doc

def main():
    st.title("Word Document Translator")

    # Upload the document
    uploaded_file = st.file_uploader("Upload a Word Document", type=["docx"])
    
    if uploaded_file:
        doc = Document(uploaded_file)

        # Dropdown for language selection
        language_options = {
            "Bengali": "bn", "Hindi": "hi", "Odia": "or", "Punjabi": "pa", 
            "Tamil": "ta", "Telugu": "te", "Gujarati": "gu", "Malayalam": "ml"
        }
        target_language = st.selectbox("Select Target Language", options=list(language_options.keys()))
        language_code = language_options[target_language]

        if st.button("Translate Document"):
            with st.spinner('Translating... Please wait...'):
                translated_doc = translate_doc(doc, language_code)

                # Save the translated document in memory
                translated_file = BytesIO()
                translated_doc.save(translated_file)
                translated_file.seek(0)  # Reset the file pointer

                st.success("Translation complete!")

                # Provide the download button
                st.download_button(
                    label="Download Translated Document",
                    data=translated_file,
                    file_name=f"translated_document_{language_code}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )

if __name__ == '__main__':
    main()
