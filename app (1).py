import streamlit as st
import pandas as pd
from io import BytesIO

# Configuração da página
st.set_page_config(
    page_title="Sistema de Controle e Comparação de Preços",
    page_icon="logo-eqtl-app-teste2.png",
    layout="wide"
)

# Caminho das planilhas base e exceção (definidos manualmente no código)
CAMINHO_BASE = "planilha_base.xlsx"
CAMINHO_EXCECAO = "planilha_excecao.XLSX"

# Lista de colunas esperadas na planilha base
COLUNAS_ESPERADAS_BASE = ["EMPRESA", "Equipamento", "DESC_MATERIAL", "MAX_PU", "MIN_PU"]

# Lista de colunas esperadas na nova planilha de comparação
COLUNAS_ESPERADAS_COMPARACAO = [
    "Empresa", "Elemento PEP", "Objeto", "Denominação de objeto", "Classe de custo",
    "Descr.classe custo", "Denom.classe custo", "Documento de compras", "Nº documento",
    "Material", "Texto breve de material", "Qtd.total entrada", "Unid.medida lançada",
    "Valor/moeda objeto", "Denominação", "Nome do usuário", "Nº doc.de referência",
    "Data de lançamento", "Hora do registro", "Centro", "Data de entrada",
    "Tipo de documento", "Exercício", "Divisão", "Data do documento",
    "Linha de lançamento", "Classificação", "ODI Aneel", "Descrição SA",
    "Setor de atividade", "Documento de estorno", "Org.estorno", "estornado",
    "Nº ref.estorno", "Operação ref."
]

# Colunas que devem estar na planilha processada
COLUNAS_PROCESSADAS = [
    "Empresa", "Elemento PEP", "Material", "DESC_MATERIAL", "Qtd.total entrada",
    "
