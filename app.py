import io
import os
import tempfile
import streamlit as st
from google import genai
from google.genai import types
from docx import Document
from openpyxl import load_workbook
from pptx import Presentation
from pdf2docx import Converter

# --- Nastavenie stránky ---
st.set_page_config(
    page_title="AI Prekladač Dokumentov",
    page_icon="⚡",
    layout="centered"
)

st.title("⚡ AI Prekladač Dokumentov (Ultra-Rýchly)")
st.caption("Prekladaj DOCX, XLSX, PPTX a PDF so zachovaním pôvodného formátovania a tabuliek.")

# --- Správa API kľúča ---
api_key = None
if "GEMINI_API_KEY" in st.secrets:
    api_key = st.secrets["GEMINI_API_KEY"]
else:
    api_key = st.sidebar.text_input("Zadaj Gemini API Key", type="password")

if not api_key:
    st.info("🔑 Na pokračovanie vlož svoj Gemini API kľúč v bočnom paneli (alebo ho nastav v Secrets v Streamlit Cloude).")
    st.stop()

# Inicializácia Gemini klienta
client = genai.Client(api_key=api_key)

DELIMITER = "\n===ITEM_DELIMITER===\n"

def translate_batch(text_list: list[str], target_language: str) -> list[str]:
    """Preloží celý zoznam textov v jedinom API volaní."""
    if not text_list:
        return []
    
    combined_text = DELIMITER.join(text_list)

    prompt = f"""You are a professional context-aware translator.
Task: Translate each segment of the following text into target language: '{target_language}'.
Rules:
1. Automatically detect the source language.
2. Maintain the EXACT same layout, formatting variables, placeholders, numbers, and technical terminology.
3. The segments are separated by the exact delimiter string: '===ITEM_DELIMITER==='.
4. Return the translated segments separated by the EXACT SAME delimiter '===ITEM_DELIMITER==='.
5. Do NOT alter, add, or remove delimiters. Return ONLY the translated segments and delimiters.

Text to translate:
{combined_text}"""

    try:
        response = client.models.generate_content(
            model='gemini-2.5-flash',
            contents=prompt,
            config=types.GenerateContentConfig(
                temperature=0.1,
            )
        )
        if not response.text:
            return text_list
            
        translated_segments = response.text.strip().split("===ITEM_DELIMITER===")
        translated_segments = [s.strip() for s in translated_segments]
        
        # Ak model náhodou nesprávne vrátil počet častí, použije pôvodný text ako zálohu
        if len(translated_segments) == len(text_list):
            return translated_segments
        else:
            return text_list
    except Exception as e:
        st.error(f"Chyba pri batch preklade: {e}")
        return text_list

def process_in_batches(elements: list, get_text_fn, set_text_fn, target_lang: str, batch_size: int = 60):
    """Spracováva texty vo veľkých dávkach (60 položiek naraz) s progresom."""
    valid_elements = [el for el in elements if get_text_fn(el) and get_text_fn(el).strip()]
    total_elements = len(valid_elements)
    
    if total_elements == 0:
        return

    progress_bar = st.progress(0.0)
    status_text = st.empty()

    for i in range(0, total_elements, batch_size):
        batch = valid_elements[i:i + batch_size]
        texts = [get_text_fn(el) for el in batch]
        
        # Aktualizácia progresu
        current_progress = min((i + len(batch)) / total_elements, 1.0)
        progress_bar.progress(current_progress)
        status_text.text(f"⚡ Preložených {i + len(batch)} / {total_elements} blokov textu...")

        translated_texts = translate_batch(texts, target_lang)
        
        for el, trans_text in zip(batch, translated_texts):
            set_text_fn(el, trans_text)

    progress_bar.empty()
    status_text.empty()

# --- Funkcie pre spracovanie dokumentov ---

def process_docx(file_bytes: bytes, target_lang: str) -> bytes:
    doc = Document(io.BytesIO(file_bytes))
    
    # 1. Odseky
    paragraphs = [p for p in doc.paragraphs if p.text.strip()]
    if paragraphs:
        st.caption("📝 Prekladám odseky...")
        process_in_batches(paragraphs, lambda p: p.text, lambda p, t: setattr(p, 'text', t), target_lang)

    # 2. Tabuľky
    cells = []
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if cell.text.strip():
                    cells.append(cell)
    if cells:
        st.caption("📊 Prekladám tabuľky...")
        process_in_batches(cells, lambda c: c.text, lambda c, t: setattr(c, 'text', t), target_lang)

    output = io.BytesIO()
    doc.save(output)
    return output.getvalue()

def process_xlsx(file_bytes: bytes, target_lang: str) -> bytes:
    wb = load_workbook(filename=io.BytesIO(file_bytes))
    
    cells = []
    for sheet in wb.worksheets:
        for row in sheet.iter_rows():
            for cell in row:
                if cell.value and isinstance(cell.value, str) and not cell.value.startswith("="):
                    if cell.value.strip():
                        cells.append(cell)
                        
    if cells:
        st.caption("📊 Prekladám Excel hárky...")
        process_in_batches(cells, lambda c: str(c.value), lambda c, t: setattr(c, 'value', t), target_lang)

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

def process_pptx(file_bytes: bytes, target_lang: str) -> bytes:
    prs = Presentation(io.BytesIO(file_bytes))
    
    paragraphs = []
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                for p in shape.text_frame.paragraphs:
                    if p.text.strip():
                        paragraphs.append(p)
            if shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        if cell.text.strip():
                            paragraphs.append(cell)

    if paragraphs:
        st.caption("🖼️ Prekladám snímky prezentácie...")
        process_in_batches(paragraphs, lambda p: p.text, lambda p, t: setattr(p, 'text', t), target_lang)

    output = io.BytesIO()
    prs.save(output)
    return output.getvalue()

def process_pdf(file_bytes: bytes, target_lang: str) -> bytes:
    status_info = st.info("⚙️ Prebieha konverzia PDF rozloženia na upraviteľný DOCX formát...")
    
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_pdf:
        temp_pdf.write(file_bytes)
        temp_pdf_path = temp_pdf.name

    temp_docx_path = temp_pdf_path.replace(".pdf", ".docx")

    try:
        cv = Converter(temp_pdf_path)
        cv.convert(temp_docx_path)
        cv.close()

        status_info.empty()

        with open(temp_docx_path, "rb") as f:
            docx_bytes = f.read()

        translated_docx_bytes = process_docx(docx_bytes, target_lang)
        return translated_docx_bytes

    finally:
        if os.path.exists(temp_pdf_path):
            os.remove(temp_pdf_path)
        if os.path.exists(temp_docx_path):
            os.remove(temp_docx_path)

# --- Používateľské rozhranie ---

st.markdown("---")

uploaded_file = st.file_uploader(
    "Nahraj súbor (podporované: DOCX, XLSX, PPTX, PDF)", 
    type=["docx", "xlsx", "pptx", "pdf"]
)

target_language = st.selectbox(
    "Vyber cieľový jazyk prekladu",
    [
        "Slovenčina",
        "Čeština",
        "Angličtina",
        "Nemčina",
        "Poľština",
        "Francúzština",
        "Španielčina",
        "Taliančina",
        "Ukrajinčina"
    ]
)

if uploaded_file and st.button("🚀 Spustiť preklad", type="primary"):
    file_bytes = uploaded_file.read()
    file_ext = uploaded_file.name.split(".")[-1].lower()
    
    try:
        if file_ext == "docx":
            result_bytes = process_docx(file_bytes, target_language)
            mime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            out_ext = "docx"
        elif file_ext == "xlsx":
            result_bytes = process_xlsx(file_bytes, target_language)
            mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            out_ext = "xlsx"
        elif file_ext == "pptx":
            result_bytes = process_pptx(file_bytes, target_language)
            mime = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
            out_ext = "pptx"
        elif file_ext == "pdf":
            result_bytes = process_pdf(file_bytes, target_language)
            mime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            out_ext = "docx"
        else:
            st.error("Nepodporovaný formát súboru.")
            st.stop()
            
        st.success("✅ Preklad bol úspešne dokončený!")
        
        base_name = uploaded_file.name.rsplit('.', 1)[0]
        output_filename = f"prelozene_{base_name}.{out_ext}"
        
        st.download_button(
            label=f"📥 Stiahnuť preložený súbor ({output_filename})",
            data=result_bytes,
            file_name=output_filename,
            mime=mime
        )
    except Exception as e:
        st.error(f"Vyskytla sa chyba pri spracovaní súboru: {e}")
