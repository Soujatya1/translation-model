import streamlit as st
from docx import Document
from docx.shared import Pt  # For setting font size
from deep_translator import GoogleTranslator

def translate_doc(doc, destination='hi'):
    """
    Translate a Word document while ensuring all text is properly replaced and appears in the output document.
    Also preserves basic formatting such as font size.
    :param doc: Word doc object (from Document class)
    :param destination: Target language (default is Hindi 'hi')
    """
    translator = GoogleTranslator(source='auto', target=destination)

    # Translate paragraphs
    for p in doc.paragraphs:
        if p.text.strip():
            try:
                translated_text = translator.translate(p.text.strip()) or p.text  # Fallback to original text
                p.clear()
                run = p.add_run(translated_text)
                run.font.size = Pt(8)  # Set font size to 8pt
            except Exception as e:
                print(f"Error translating paragraph: {e}")

    # Translate table cells
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if cell.text.strip():
                    try:
                        full_text = "\n".join([para.text.strip() for para in cell.paragraphs])
                        translated_text = translator.translate(full_text) or full_text  # Fallback to original text
                        cell.text = ""  # Clear cell text properly
                        p = cell.add_paragraph(translated_text)
                        run = p.runs[0]
                        run.font.size = Pt(8)  # Set font size to 8pt
                    except Exception as e:
                        print(f"Error translating cell text: {e}")
    
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
                translated_doc_path = "translated_document.docx"
                translated_doc.save(translated_doc_path)

                with open(translated_doc_path, "rb") as f:
                    st.download_button(
                        label="Download Translated Document",
                        data=f,
                        file_name="translated_document.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                st.success("Translation complete!")

if __name__ == '__main__':
    main()
