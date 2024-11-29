import streamlit as st
from PyPDF2 import PdfMerger, PdfReader, PdfWriter
from comtypes.client import CreateObject
from comtypes import CoInitialize, CoUninitialize
from os import path, remove


# Configuração da página
st.set_page_config(page_title="Manipulação de PDFs", layout="wide")

st.title("Ferramentas de Manipulação de PDF :umbrella_with_rain_drops:")
st.sidebar.header("Escolha uma funcionalidade :balloon:")
# Seleção da funcionalidade
funcionalidade = st.sidebar.selectbox(
    "O que você quer fazer?",
    ["Unir PDFs", "Dividir PDF", "Compactar PDF", "Proteger PDF", "Converter Word para PDF", "Converter Imagem em PDF"]
)
# Função: Converter Word para PDF
def converter_word_para_pdf(caminho_word, caminho_pdf):
    """
    Converte um arquivo Word (.docx) em PDF.
    """
    try:
        CoInitialize()  # Inicializa o sistema COM
        word = CreateObject("Word.Application")
        doc = word.Documents.Open(caminho_word)
        doc.SaveAs(caminho_pdf, FileFormat=17)  # 17 corresponde ao formato PDF
        doc.Close()
        word.Quit()
        return caminho_pdf
    except Exception as e:
        return f"Erro ao converter Word para PDF: {e}"
    finally:
        CoUninitialize()

with st.sidebar: "[By Instagram: @wemersonsilva_](https://www.instagram.com/wemersonsillva_/)"
with st.sidebar: "[By Linkedin: @wemersonsilva_](https://www.linkedin.com/in/wemerson-silva-5098071a9/)"
if funcionalidade == "Unir PDFs":
    st.subheader("Unir Múltiplos PDFs :sparkles:")
    pdf_files = st.file_uploader("Envie os arquivos PDF para unir", type="pdf", accept_multiple_files=True)
    if st.button("Unir PDFs"):
        if pdf_files:
            merger = PdfMerger()
            for pdf in pdf_files:
                merger.append(pdf)
            # Salvar o arquivo combinado
            output_pdf = "unidos.pdf"
            merger.write(output_pdf)
            merger.close()
            # Oferecer o download do arquivo combinado
            with open(output_pdf, "rb") as f:
                st.download_button("Baixar PDF Unido", f, file_name="unidos.pdf")
        else:
            st.warning("Por favor, envie pelo menos dois arquivos PDF.")

elif funcionalidade == "Dividir PDF":
    st.subheader("Dividir um PDF")
    pdf_file = st.file_uploader("Envie o arquivo PDF para dividir", type="pdf")
    if pdf_file:
        pages = st.text_input("Digite o intervalo de páginas (ex: 1-3, 5, 7-9):")
        if st.button("Dividir PDF"):
            try:
                reader = PdfReader(pdf_file)
                writer = PdfWriter()
                ranges = pages.split(",")
                for r in ranges:
                    if "-" in r:
                        start, end = map(int, r.split("-"))
                        for page_num in range(start - 1, end):
                            writer.add_page(reader.pages[page_num])
                    else:
                        writer.add_page(reader.pages[int(r) - 1])
                # Salvar o arquivo dividido
                output_pdf = "dividido.pdf"
                with open(output_pdf, "wb") as f:
                    writer.write(f)
                with open(output_pdf, "rb") as f:
                    st.download_button("Baixar PDF Dividido", f, file_name="dividido.pdf")
            except Exception as e:
                st.error(f"Erro ao dividir PDF: {e}")


elif funcionalidade == "Converter Word para PDF":
    st.subheader("Converter Word para PDF")
    word_file = st.file_uploader("Envie o arquivo Word para converter", type=["docx", "doc"])
    if st.button("Converter para PDF"):
        if word_file:
            # Salvar o arquivo Word enviado
            word_temp = word_file.name
            with open(word_temp, "wb") as f:
                f.write(word_file.read())

            # Nome do arquivo PDF gerado
            pdf_output = word_temp.replace(".docx", ".pdf").replace(".doc", ".pdf")
            resultado = converter_word_para_pdf(word_temp, pdf_output)

            if "Erro" in resultado:
                st.error(resultado)
            else:
                with open(pdf_output, "rb") as f:
                    st.download_button("Baixar PDF Convertido", f, file_name=pdf_output)

            # Limpeza do arquivo temporário
            remove(word_temp)
            if path.exists(pdf_output):
                remove(pdf_output)
        else:
            st.warning("Por favor, envie um arquivo Word.")