import streamlit as st
import requests
from PIL import Image
from io import BytesIO
from openpyxl import load_workbook
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import fitz  # PyMuPDF

# Configuração da página do Streamlit
st.set_page_config(page_title="SPOT - Automação FFM", layout="wide")

file_path = './planilha_modelo.xlsx'  # Caminho do arquivo Excel
file_name = 'planilha_modelo.xlsx'  # Nome do arquivo para download

# Função para criar o botão de download
def download_button(text, file_path, file_name, mime_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'):
    with open(file_path, 'rb') as f:
        file_bytes = f.read()
    st.download_button(label=text, data=file_bytes, file_name=file_name, mime=mime_type)

# Função para inserir o logo PNG
def inserir_logo():
    logo_path = "./logo_spot.png"  # Caminho do logo
    logo_img = Image.open(logo_path)
    st.image(logo_img, width=250)

# Chamando a função para inserir o logo
inserir_logo()
st.title("Automação FFM - Evidências Ação de Relacionamento")
download_button('Planilha Modelo', file_path, file_name)

# Função para extrair links por relatório
def extrair_links_por_relatorio(file):
    wb = load_workbook(file, data_only=True)
    ws = wb.active

    cabecalho = [cell.value for cell in ws[1]]
    colunas_por_categoria = {
        "FOTO DA AÇÃO E CONQUISTA": list(range(1, 5)),
        "FOTO DA LISTA DE PRESENÇA": list(range(5, 9)),
        "UMA FOTO DA NOTA FISCAL": list(range(9, 13)),
    }

    dados = []
    for row in ws.iter_rows(min_row=2):
        numero_relatorio = row[0].value
        if not numero_relatorio:
            continue

        imagens_por_tipo = {}
        for categoria, indices in colunas_por_categoria.items():
            imagens = []
            for idx in indices:
                cell = row[idx]
                url = cell.hyperlink.target if cell.hyperlink else None
                if url:
                    imagens.append(url)
            imagens_por_tipo[categoria] = imagens

        dados.append((numero_relatorio, imagens_por_tipo))

    return dados

# Função para converter PDF em imagens
def pdf_para_imagens(pdf_bytes):
    imagens = []
    with fitz.open(stream=pdf_bytes, filetype="pdf") as doc:
        for page in doc:
            pix = page.get_pixmap(dpi=150)
            img = Image.open(BytesIO(pix.tobytes("png")))
            imagens.append(img)
    return imagens

# Função para ajustar altura do parágrafo no documento Word
def ajustar_altura_doc_paragrafo(paragraph):
    p = paragraph._element
    pPr = p.get_or_add_pPr()
    pPr.append(OxmlElement('w:keepLines'))
    pPr.append(OxmlElement('w:keepNext'))

# Função para inserir imagem redimensionada no documento Word
def inserir_imagem_redimensionada(paragraph, img, largura_max=5.5, altura_max=7):
    img_io = BytesIO()
    img.save(img_io, format='PNG')
    img_io.seek(0)
    largura, altura = img.size
    escala = min((largura_max * 96) / largura, (altura_max * 96) / altura)
    escala *= 1.1  # leve aumento

    nova_largura = largura * escala / 96
    run = paragraph.add_run()
    run.add_picture(img_io, width=Inches(nova_largura))

# Função para aplicar fonte Arial no documento Word
def aplicar_fonte_arial(run):
    run.font.name = "Arial"
    run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Arial')
    run.font.size = Pt(12)

# Upload da planilha Excel
uploaded_file = st.file_uploader("📂 Envie a planilha (.xlsx)", type=["xlsx"])

if uploaded_file:
    try:
        info_links = extrair_links_por_relatorio(uploaded_file)

        if not info_links:
            st.error("❌ Nenhum link encontrado na planilha.")
        else:
            st.success(f"✅ {len(info_links)} relatórios encontrados.")

            if st.button("📝 Gerar Documento Word"):
                erros = []
                doc = Document()
                log_area = st.empty()

                for i, (num_relatorio, grupos) in enumerate(info_links, 1):
                    try:
                        log_area.markdown(f"🔄 Processando relatório {num_relatorio} ({i}/{len(info_links)})")

                        for categoria, links in grupos.items():
                            if not links:
                                continue

                            doc.add_page_break()

                            # Título do relatório + categoria
                            p = doc.add_paragraph()
                            ajustar_altura_doc_paragrafo(p)
                            p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                            run = p.add_run(f"Relatório: {num_relatorio} — {categoria}")
                            aplicar_fonte_arial(run)

                            for link in links:
                                try:
                                    if not link.startswith("http"):
                                        link = "https://" + link
                                    response = requests.get(link, timeout=20)
                                    response.raise_for_status()

                                    content_type = response.headers.get('Content-Type', '')
                                    if 'pdf' in content_type:
                                        imagens = pdf_para_imagens(response.content)
                                    else:
                                        img = Image.open(BytesIO(response.content)).convert("RGB")
                                        extrema = img.getextrema()
                                        if all(e[0] == e[1] for e in extrema):
                                            raise ValueError("Imagem aparentemente em branco.")
                                        imagens = [img]

                                    for img in imagens:
                                        p_img = doc.add_paragraph()
                                        p_img.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                                        inserir_imagem_redimensionada(p_img, img)

                                except Exception as e_img:
                                    p_err = doc.add_paragraph()
                                    p_err.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                                    run_err = p_err.add_run(f"⚠️ Erro ao carregar imagem: {e_img}")
                                    aplicar_fonte_arial(run_err)

                    except Exception as e:
                        erros.append((num_relatorio, e))

                buffer = BytesIO()
                doc.save(buffer)
                buffer.seek(0)

                log_area.empty()
                st.success("✅ Documento Word gerado com sucesso!")
                st.download_button(
                    label="📥 Baixar Word",
                    data=buffer,
                    file_name="evidencias_acao_relacionamento.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )

                if erros:
                    st.markdown("### ❌ Falhas detectadas")
                    for relatorio, erro in erros:
                        st.write(f"Relatório {relatorio} → {erro}")

    except Exception as e:
        st.error(f"Erro ao processar: {e}")
