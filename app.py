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
    page_icon="🌐",
    layout="centered"
)

st.title("🌐 AI Prekladač Dokumentov")
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

# --- Pomocná funkcia na preklad cez Gemini ---
def translate_text(text: str, target_language: str) -> str:
    """Preloží text pomocou Gemini 2.5 Flash. Automaticky deteguje zdrojový jazyk."""
    if not text or not text.strip():
        return text

    prompt = f"""You are a professional context-aware translator.
Task: Translate the following text into target language: '{target_language}'.
Rules:
1. Automatically detect the source language.
2. Preserve all original line breaks, formatting variables, placeholders, numbers, and technical terminology.
3. Return ONLY the translation, with no explanation or meta-text.

Text to translate:
{text}"""

    try:
        response = client.models.generate_content(
            model='gemini-2.5-flash',
            contents=prompt,
            config=types.GenerateContentConfig(
                temperature=0.1,  # Nízka teplota pre presný preklad
            )
        )
        return response.text.strip() if response.text else text
    except Exception as e:
        st.error(f"Chyba pri komunikácii s API: {e}")
        return text

# --- Funkcie pre spracovanie dokumentov ---

def process_docx(file_bytes: bytes, target_lang: str) -> bytes:
    """Spracuje DOCX súbor - preloží odseky a tabuľky, pričom zachová odsekové štýly a formátovanie."""
    doc = Document(io.BytesIO(file_bytes))
    
    status_text = st.empty()
    status_text.text("⏳ Prekladám odseky v dokumente...")
    
    # 1. Preklad obyčajných odsekov
    for p in doc.paragraphs:
        if p.text.strip():
            translated = translate_text(p.text, target_lang)
            p.text = translated

    # 2. Preklad tabuliek
    status_text.text("⏳ Prekladám tabuľky v dokumente...")
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if cell.text.strip():
                    translated = translate_text(cell.text, target_lang)
                    cell.text = translated

    status_text.empty()
    output = io.BytesIO()
    doc.save(output)
    return output.getvalue()

def process_xlsx(file_bytes: bytes, target_lang: str) -> bytes:
    """Spracuje XLSX súbor - preloží hodnoty buniek bez porušenia vzorcov a štýlov."""
    wb = load_workbook(filename=io.BytesIO(file_bytes))
    status_text = st.empty()
    status_text.text("⏳ Prekladám hárky v Exceli...")
    
    for sheet in wb.worksheets:
        for row in sheet.iter_rows():
            for cell in row:
                if cell.value and isinstance(cell.value, str) and not cell.value.startswith("="):
                    if cell.value.strip():
                        translated = translate_text(cell.value, target_lang)
                        cell.value = translated

    status_text.empty()
    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

def process_pptx(file_bytes: bytes, target_lang: str) -> bytes:
    """Spracuje PPTX súbor - preloží textové polia a tabuľky v snímkach."""
    prs = Presentation(io.BytesIO(file_bytes))
    status_text = st.empty()
    
    for i, slide in enumerate(prs.slides):
        status_text.text(f"⏳ Prekladám snímku {i+1} z {len(prs.slides)}...")
        for shape in slide.shapes:
            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    if paragraph.text.strip():
                        translated = translate_text(paragraph.text, target_lang)
                        paragraph.text = translated
            
            if shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        if cell.text.strip():
                            translated = translate_text(cell.text, target_lang)
                            cell.text = translated

    status_text.empty()
    output = io.BytesIO()
    prs.save(output)
    return output.getvalue()

def process_pdf(file_bytes: bytes, target_lang: str) -> bytes:
    """Prekonvertuje PDF na DOCX, preloží ho a vráti výsledný vygenerovaný DOCX súbor."""
    status_text = st.empty()
    status_text.text("⏳ Konvertujem PDF na upraviteľný dokument...")
    
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_pdf:
        temp_pdf.write(file_bytes)
        temp_pdf_path = temp_pdf.name

    temp_docx_path = temp_pdf_path.replace(".pdf", ".docx")

    try:
        cv = Converter(temp_pdf_path)
        cv.convert(temp_docx_path)
        cv.close()

        with open(temp_docx_path, "rb") as f:
            docx_bytes = f.read()

        status_text.empty()
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
    
    with st.spinner("Prekladám dokument. Prosím počkaj..."):
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
                
            st.success("✅ Preklad je dokončený!")
            
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
