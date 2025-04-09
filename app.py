import streamlit as st
from PyPDF2 import PdfMerger
from io import BytesIO
import pandas as pd
import subprocess
import fitz  # PyMuPDF
import yagmail
#librerias para OCR
import pytesseract
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
import platform
from pdf2image import convert_from_bytes
from PIL import Image
import io

def merge_pdfs(pdf_files):
    merger = PdfMerger()
    for pdf in pdf_files:
        merger.append(pdf)
    output = BytesIO()
    merger.write(output)
    merger.close()
    output.seek(0)  # Regresamos al inicio del archivo
    return output

def ocr_pdf_windows(pdf_bytes):
    poppler_path = r"C:\poppler-24.08.0\Library\bin"
    imagenes = convert_from_bytes(pdf_bytes, poppler_path=poppler_path)

    os.environ["TESSDATA_PREFIX"] = r"C:\Program Files\Tesseract-OCR\tessdata"

    texto = ""
    for i, img in enumerate(imagenes):
        st.write(f"Página {i + 1}")
        st.image(img, use_container_width=True)
        texto += f"\n--- Página {i + 1} ---\n" + pytesseract.image_to_string(img, lang="spa")
    return texto


def ocr_pdf_linux_mac(pdf_bytes):
    texto = ""
    with fitz.open(stream=pdf_bytes, filetype="pdf") as doc:
        for i, page in enumerate(doc):
            pix = page.get_pixmap(dpi=300)
            img = Image.open(io.BytesIO(pix.tobytes("png")))
            st.write(f"Página {i + 1}")
            st.image(img, use_container_width=True)
            texto += f"\n--- Página {i + 1} ---\n" + pytesseract.image_to_string(img, lang="spa")
    return texto

st.title("Mi Automatizador Personal en Python")

menu = st.sidebar.selectbox("Selecciona una tarea", ["Enviar correo",
                                                     "Procesar Excel",
                                                     "Lanzar Script PowerShell",
                                                     "Extraer texto de pdf",
                                                     "Extraer texto de pdf escaneado",
                                                     "Unir PDF's"])

if menu == "Enviar correo":
    st.subheader("Enviar un correo automático")
    to = st.text_input("Destinatario")
    subject = st.text_input("Asunto")
    body = st.text_area("Mensaje")
    if st.button("Enviar"):
        yag = yagmail.SMTP("charliegarcia0@gmail.com", "ejte skso uzcv qors")
        yag.send(to=to, subject=subject, contents=body)
        st.success("Correo enviado")

elif menu == "Procesar Excel":
    file = st.file_uploader("Sube un archivo Excel", type=["xlsx"])
    if file:
        df = pd.read_excel(file)
        st.write("Vista previa:", df.head())
        resumen = df.groupby("Categoría").sum()
        st.write("Resumen:", resumen)

elif menu == "Lanzar Script PowerShell":
    st.subheader("Lanzar script remoto")
    comando = st.text_input("Escribe el comando PowerShell")
    if st.button("Ejecutar"):
        result = subprocess.run(["powershell", "-Command", comando], capture_output=True, text=True)
        st.text("Salida:")
        st.code(result.stdout)

elif menu == "Extraer texto de pdf":
    file_pdf = st.file_uploader("📂 Carga un archivo PDF", type="pdf")
    if file_pdf:
        # Abrir el archivo usando PyMuPDF desde un buffer (archivo en memoria)
        with fitz.open(stream=file_pdf.read(), filetype="pdf") as doc:
            texto_total = ""
            for pagina in doc:
                texto_total += pagina.get_text()

        st.subheader("📄 Texto extraído del PDF:")
        st.text_area("Contenido", texto_total, height=400)

elif menu == "Extraer texto de pdf escaneado":

    archivo_pdf = st.file_uploader("📂 Carga un PDF escaneado", type="pdf")
    sistema = platform.system()

    if archivo_pdf:

        st.info("Procesando OCR...")

        # Leer el archivo una sola vez

        pdf_bytes = archivo_pdf.read()

        if sistema == "Windows":
            texto_total = ocr_pdf_windows(pdf_bytes)
        else:
            texto_total = ocr_pdf_linux_mac(pdf_bytes)

        if texto_total.strip():
            st.subheader("📄 Texto extraído (OCR):")
            st.text_area("Resultado OCR", texto_total, height=400)
            st.download_button("Descargar texto como .txt", data=texto_total, file_name="texto_extraido.txt",
                               mime="text/plain")
        else:
            st.warning("No se detectó texto en el documento.")

        st.download_button(
            label="Descargar texto como .txt",
            data=texto_total,
            file_name="texto_extraido.txt",
            mime="text/plain"
        )

elif menu == "Unir PDF's":
    st.write("Sube varios archivos PDF y combínalos en uno solo.")

    uploaded_files = st.file_uploader("Selecciona los archivos PDF", type="pdf", accept_multiple_files=True)

    if uploaded_files:
        if st.button("Unir PDFs"):
            with st.spinner("Uniendo archivos..."):
                merged_pdf = merge_pdfs(uploaded_files)

                st.success("✅ PDF combinado correctamente")
                st.download_button(
                    label="📥 Descargar PDF combinado",
                    data=merged_pdf,
                    file_name="pdf_unido.pdf",
                    mime="application/pdf"
                )