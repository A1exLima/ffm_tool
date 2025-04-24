import os
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

# Definições de caminho
BASE_DIR = os.path.dirname(__file__)
TEMPLATE_RELACIONAMENTO = os.path.join(BASE_DIR, "planilha_modelo_relacionamento.xlsx")
TEMPLATE_ALIMENTACAO = os.path.join(BASE_DIR, "planilha_modelo_alimentacao.xlsx")

# Configuração da página
st.set_page_config(page_title="SPOT - Automação FFM", layout="wide")

# Função para inserir o logo PNG
def inserir_logo():
    logo_path = os.path.join(BASE_DIR, "logo_spot.png")
    logo_img = Image.open(logo_path)
    st.image(logo_img, width=250)

# Função para criar o botão de download
def download_button(text, file_path, file_name, mime_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'):
    with open(file_path, 'rb') as f:
        file_bytes = f.read()
    st.download_button(label=text, data=file_bytes, file_name=file_name, mime=mime_type)

# Função para ajustar altura do parágrafo no Word
def ajustar_altura_doc_paragrafo(paragraph):
    p = paragraph._element
    pPr = p.get_or_add_pPr()
    pPr.append(OxmlElement('w:keepLines'))
    pPr.append(OxmlElement('w:keepNext'))

# Função para inserir imagem redimensionada no Word
def inserir_imagem_redimensionada(paragraph, img, largura_max=5.5, altura_max=7):
    img_io = BytesIO()
    img.save(img_io, format='PNG')
    img_io.seek(0)
    largura, altura = img.size
    escala = min((largura_max * 96) / largura, (altura_max * 96) / altura) * 1.1
    nova_largura = largura * escala / 96
    run = paragraph.add_run()
    run.add_picture(img_io, width=Inches(nova_largura))

# Função para aplicar fonte Arial em um run
def aplicar_fonte_arial(run):
    run.font.name = "Arial"
    run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Arial')
    run.font.size = Pt(12)

# Extrair links para Ação de Relacionamento
def extrair_links_por_relatorio(file):
    wb = load_workbook(file, data_only=True)
    ws = wb.active
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

# Extrair links para Alimentação (URLs como texto)
def extrair_links_por_alimentacao(file):
    wb = load_workbook(file, data_only=True)
    ws = wb.active
    header = [cell.value for cell in ws[1]]
    cat1 = header[1] or "Comprovante Capturado"
    cat2 = header[2] or "Outras Evidências"
    dados = []
    for row in ws.iter_rows(min_row=2):
        id_reembolso = row[0].value
        if not id_reembolso:
            continue
        grupos = {}
        links1 = []
        if row[1].value:
            links1.append(str(row[1].value))
        grupos[cat1] = links1
        links2 = []
        if row[2].value:
            links2.append(str(row[2].value))
        grupos[cat2] = links2
        dados.append((id_reembolso, grupos))
    return dados

# Converter PDF em imagens
def pdf_para_imagens(pdf_bytes):
    imagens = []
    with fitz.open(stream=pdf_bytes, filetype="pdf") as doc:
        for page in doc:
            pix = page.get_pixmap(dpi=150)
            img = Image.open(BytesIO(pix.tobytes("png")))
            imagens.append(img)
    return imagens

# Interface
inserir_logo()
funcao = st.sidebar.radio("Opções", ["Ação de Relacionamento", "Alimentação"])

if funcao == "Ação de Relacionamento":
    st.title("Automação FFM - Evidências Ação de Relacionamento")
    download_button('Planilha Modelo', TEMPLATE_RELACIONAMENTO, 'planilha_modelo_relacionamento.xlsx')
    uploaded_file = st.file_uploader("📂 Envie a planilha de Relacionamento (.xlsx)", type=["xlsx"], key="rel_upload")
    if uploaded_file:
        info_links = extrair_links_por_relatorio(uploaded_file)
        if not info_links:
            st.error("❌ Nenhum link encontrado na planilha.")
        else:
            st.success(f"✅ {len(info_links)} relatórios encontrados.")
            if st.button("📝 Gerar Documento Word", key="btn_rel"): 
                doc = Document()
                log_area = st.empty()
                for i, (num_relatorio, grupos) in enumerate(info_links, 1):
                    log_area.markdown(f"🔄 Processando relatório {num_relatorio} ({i}/{len(info_links)})")
                    for categoria, links in grupos.items():
                        if not links:
                            continue
                        doc.add_page_break()
                        p = doc.add_paragraph()
                        ajustar_altura_doc_paragrafo(p)
                        p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                        run = p.add_run(f"Relatório: {num_relatorio} — {categoria}")
                        aplicar_fonte_arial(run)
                        for link in links:
                            try:
                                if not link.startswith("http"):
                                    link = "https://" + link
                                resp = requests.get(link, timeout=20)
                                resp.raise_for_status()
                                ct = resp.headers.get('Content-Type', '')
                                if 'pdf' in ct:
                                    imgs = pdf_para_imagens(resp.content)
                                else:
                                    img = Image.open(BytesIO(resp.content)).convert("RGB")
                                    extrema = img.getextrema()
                                    if all(e[0]==e[1] for e in extrema):
                                        raise ValueError("Imagem em branco.")
                                    imgs = [img]
                                for img in imgs:
                                    p_img = doc.add_paragraph()
                                    p_img.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                                    inserir_imagem_redimensionada(p_img, img)
                            except Exception as e_img:
                                p_err = doc.add_paragraph()
                                p_err.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                                r_err = p_err.add_run(f"⚠️ Erro ao carregar imagem: {e_img}")
                                aplicar_fonte_arial(r_err)
                buffer = BytesIO()
                doc.save(buffer)
                buffer.seek(0)
                log_area.empty()
                st.success("✅ Documento Word gerado!")
                st.download_button(
                    label="📥 Baixar Word - Relacionamento", 
                    data=buffer, 
                    file_name="evidencias_acao_relacionamento.docx", 
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
elif funcao == "Alimentação":
    st.title("Automação FFM - Evidências Alimentação")
    download_button('Planilha Modelo', TEMPLATE_ALIMENTACAO, 'planilha_modelo_alimentacao.xlsx')
    uploaded_file = st.file_uploader("📂 Envie a planilha de Alimentação (.xlsx)", type=["xlsx"], key="alim_upload")
    if uploaded_file:
        info_links = extrair_links_por_alimentacao(uploaded_file)
        if not info_links:
            st.error("❌ Nenhum link encontrado na planilha.")
        else:
            st.success(f"✅ {len(info_links)} reembolsos encontrados.")
            if st.button("📝 Gerar Documento Word - Alimentação", key="btn_alim"): 
                doc = Document()
                log_area = st.empty()
                for i, (id_reembolso, grupos) in enumerate(info_links, 1):
                    log_area.markdown(f"🔄 Processando reembolso {id_reembolso} ({i}/{len(info_links)})")
                    for categoria, links in grupos.items():
                        if not links:
                            continue
                        doc.add_page_break()
                        p = doc.add_paragraph()
                        ajustar_altura_doc_paragrafo(p)
                        p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                        run = p.add_run(f"Reembolso: {id_reembolso} — {categoria}")
                        aplicar_fonte_arial(run)
                        for link in links:
                            try:
                                if not link.startswith("http"):
                                    link = "https://" + link
                                resp = requests.get(link, timeout=20)
                                resp.raise_for_status()
                                ct = resp.headers.get('Content-Type', '')
                                if 'pdf' in ct:
                                    imgs = pdf_para_imagens(resp.content)
                                else:
                                    img = Image.open(BytesIO(resp.content)).convert("RGB")
                                    extrema = img.getextrema()
                                    if all(e[0]==e[1] for e in extrema):
                                        raise ValueError("Imagem em branco.")
                                    imgs = [img]
                                for img in imgs:
                                    p_img = doc.add_paragraph()
                                    p_img.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                                    inserir_imagem_redimensionada(p_img, img)
                            except Exception as e_img:
                                p_err = doc.add_paragraph()
                                p_err.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                                r_err = p_err.add_run(f"⚠️ Erro ao carregar imagem: {e_img}")
                                aplicar_fonte_arial(r_err)
                buffer = BytesIO()
                doc.save(buffer)
                buffer.seek(0)
                log_area.empty()
                st.success("✅ Documento Word gerado!")
                st.download_button(
                    label="📥 Baixar Word - Alimentação", 
                    data=buffer, 
                    file_name="evidencias_alimentacao.docx", 
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
