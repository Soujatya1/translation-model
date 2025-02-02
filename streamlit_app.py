import streamlit as st
from docx import Document
from deep_translator import GoogleTranslator
from concurrent.futures import ThreadPoolExecutor

def batch_translate(texts, destination='hi'):
    """
    Translate a batch of texts in a single API call.
    :param texts: List of texts to translate
    :param destination: Target language code
    :return: List of translated texts
    """
    if not texts:
        return []  # Return an empty list if input is empty

    translator = GoogleTranslator(source='auto', target=destination)
    try:
        translated_texts = translator.translate_batch(texts)
        
        # If API returns None, return original texts
        if translated_texts is None:
            print("Warning: translate_batch() returned None, using original texts.")
            return texts
        
        # Ensure the output list is the correct length
        if len(translated_texts) != len(texts):
            print("Warning: Translated text length mismatch, using original texts.")
            return texts
        
        return translated_texts
    except Exception as e:
        print(f"Error translating batch: {e}")
        return texts  # Return original texts if translation fails

def translate_doc(doc, destination='hi'):
    """
    Translate a Word document efficiently while preserving formatting.
    :param doc: Word doc object (from `Document` class)
    :param destination: Target language
    :return: Translated Word document
    """
    # Collect all text blocks
    paragraphs = [p for p in doc.paragraphs if p.text.strip()]
    cells = [cell for table in doc.tables for row in table.rows for cell in row.cells if cell.text.strip()]
    
    # Extract text for batch translation
    paragraph_texts = [p.text for p in paragraphs]
    cell_texts = [cell.text for cell in cells]

    # Use ThreadPoolExecutor for parallel processing
    with ThreadPoolExecutor() as executor:
        future_paragraphs = executor.submit(batch_translate, paragraph_texts, destination)
        future_cells = executor.submit(batch_translate, cell_texts, destination)

        translated_paragraphs = future_paragraphs.result() or paragraph_texts
        translated_cells = future_cells.result() or cell_texts

    # Ensure translated lists are the correct length
    if len(translated_paragraphs) == len(paragraphs):
        for p, translated_text in zip(paragraphs, translated_paragraphs):
            p.text = translated_text

    if len(translated_cells) == len(cells):
        for cell, translated_text in zip(cells, translated_cells):
            cell.text = translated_text

    return doc

def main():
    st.title("Word Document Translator")

    uploaded_file = st.file_uploader("Upload a Word Document", type=["docx"])
    
    if uploaded_file:
        doc = Document(uploaded_file)

        language_options = {
            "Bengali": "bn", "Hindi": "hi", "Odia": "or", "Punjabi": "pa",
            "Tamil": "ta", "Telegu": "te", "Gujarati": "gu", "Malayalam": "ml"
        }
        target_language = st.selectbox("Select Target Language", options=list(language_options.keys()))
        language_code = language_options[target_language]

        if st.button("Translate Document"):
            with st.spinner('Translating...'):
                translated_doc = translate_doc(doc, language_code)

                with open("translated_document.docx", "wb") as f:
                    translated_doc.save(f)

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
