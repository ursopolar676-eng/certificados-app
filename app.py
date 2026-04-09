
import os
import re
import unicodedata
from io import BytesIO

import pandas as pd
import streamlit as st
from reportlab.lib.colors import white, black
from reportlab.lib.enums import TA_JUSTIFY, TA_CENTER
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas
from reportlab.platypus import Paragraph

PAGE_WIDTH = 843
PAGE_HEIGHT = 597
BACKGROUND_FILE = "background_certificado.png"

SIGNATURE_OPTIONS = {
    "Kathlenn W. Leal Martins": "kathlenn.png",
    "Vanusa Moutinho Batista Silva": "vanusa.png",
    "Elaine Salomão": "elaine.png",
    "Patricia Silva": "patricia.png",
    "Simone Ferreira Alves": "simone.png",
    "Marina Jesus do Nascimento": "marina.png",
}

def normalize_text(value: str) -> str:
    text = str(value or "").strip()
    text = unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode("ascii")
    return text.upper()

def only_digits(value: str) -> str:
    return re.sub(r"\D", "", str(value or ""))

def find_column(columns, candidates):
    normalized = {normalize_text(c): c for c in columns}
    for candidate in candidates:
        key = normalize_text(candidate)
        if key in normalized:
            return normalized[key]
    return None

@st.cache_data
def load_base():
    df = pd.read_excel("base.xlsx")
    cpf_col = find_column(df.columns, ["CPF", "CPF/CNPJ", "CPF CNPJ"])
    nome_col = find_column(df.columns, ["NOME", "NOME COMPLETO", "COLABORADOR"])
    funcao_col = find_column(df.columns, ["FUNCAO", "FUNÇÃO", "CARGO"])
    empresa_col = find_column(df.columns, ["EMPRESA", "EMPREGADOR", "RAZAO SOCIAL", "RAZÃO SOCIAL"])

    missing = []
    for label, col in [("CPF", cpf_col), ("NOME", nome_col), ("FUNÇÃO", funcao_col), ("EMPRESA", empresa_col)]:
        if col is None:
            missing.append(label)
    if missing:
        raise ValueError("Colunas não encontradas na base.xlsx: " + ", ".join(missing))

    df = df.copy()
    df["_cpf_norm"] = df[cpf_col].apply(only_digits)
    return df, cpf_col, nome_col, funcao_col, empresa_col

def build_pdf(nome, funcao, empresa, data_cert, obra, endereco, signature_filename):
    packet = BytesIO()
    c = canvas.Canvas(packet, pagesize=(PAGE_WIDTH, PAGE_HEIGHT))

    if os.path.exists(BACKGROUND_FILE):
        c.drawImage(BACKGROUND_FILE, 0, 0, width=PAGE_WIDTH, height=PAGE_HEIGHT)

    # Cover the dynamic body text area from the sample certificate
    c.setFillColor(white)
    c.rect(22, 430, 800, 92, fill=1, stroke=0)

    body_style = ParagraphStyle(
        "Body",
        fontName="Times-Roman",
        fontSize=11.2,
        leading=14.0,
        alignment=TA_JUSTIFY,
        textColor=black,
    )

    body_text = (
        f'Certifico que o Sr. <b>{nome}</b>, na função de <b>{funcao}</b>, '
        f'funcionário da empresa <b>{empresa}</b>, participou do Treinamento '
        f'de Segurança de Trabalho em conformidade com a NR 1 item 1.7 e NR18 item 18.14, '
        f'com carga horária de 06h, realizado na Obra <b>{obra}</b>, {endereco}, '
        f'onde foi orientado e treinado de acordo com o seguinte conteúdo programático:'
    )

    paragraph = Paragraph(body_text, body_style)
    paragraph.wrapOn(c, 790, 90)
    paragraph.drawOn(c, 26, 441)

    # Cover and redraw the date
    c.setFillColor(white)
    c.rect(300, 128, 260, 36, fill=1, stroke=0)
    c.setFillColor(black)
    c.setFont("Times-Bold", 12)
    c.drawCentredString(PAGE_WIDTH / 2, 144, f"São Paulo, {data_cert}.")

    # Cover original signature block and place chosen signature image
    c.setFillColor(white)
    c.rect(548, 20, 250, 118, fill=1, stroke=0)

    signature_path = os.path.join("assinaturas", signature_filename)
    if os.path.exists(signature_path):
        img = Image.open(signature_path)
        iw, ih = img.size
        target_w = 250
        target_h = target_w * (ih / iw)
        if target_h > 115:
            target_h = 115
            target_w = target_h * (iw / ih)
        x = 548 + (250 - target_w) / 2
        y = 22 + (118 - target_h) / 2
        c.drawImage(ImageReader(img), x, y, width=target_w, height=target_h, mask='auto')
    else:
        c.setFillColor(black)
        c.setFont("Helvetica-Bold", 10)
        c.drawString(560, 75, "Assinatura não encontrada")

    c.save()
    packet.seek(0)
    return packet

st.set_page_config(page_title="Certificados", page_icon="📄", layout="centered")
st.title("Gerador de Certificados")

try:
    df, cpf_col, nome_col, funcao_col, empresa_col = load_base()
except Exception as e:
    st.error(str(e))
    st.stop()

cpf = st.text_input("CPF")
data_cert = st.text_input("Data", placeholder="3 de março de 2026")
obra = st.text_input("Obra")
endereco = st.text_input("Endereço completo")
tecnica = st.selectbox("Técnica", list(SIGNATURE_OPTIONS.keys()))

if st.button("Gerar PDF"):
    cpf_norm = only_digits(cpf)
    if not cpf_norm:
        st.error("Digite um CPF.")
        st.stop()
    if not data_cert or not obra or not endereco:
        st.error("Preencha Data, Obra e Endereço.")
        st.stop()

    match = df[df["_cpf_norm"] == cpf_norm]
    if match.empty:
        st.error("CPF não encontrado na base.")
        st.stop()

    row = match.iloc[0]
    nome = str(row[nome_col]).strip()
    funcao = str(row[funcao_col]).strip()
    empresa = str(row[empresa_col]).strip()

    pdf_file = build_pdf(
        nome=nome,
        funcao=funcao,
        empresa=empresa,
        data_cert=data_cert,
        obra=obra,
        endereco=endereco,
        signature_filename=SIGNATURE_OPTIONS[tecnica],
    )

    safe_name = only_digits(cpf_norm) or "certificado"
    st.success("PDF gerado com sucesso.")
    st.download_button(
        label="Baixar PDF",
        data=pdf_file,
        file_name=f"certificado_{safe_name}.pdf",
        mime="application/pdf",
    )
