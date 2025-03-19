import streamlit as st
import pandas as pd
from io import BytesIO

# Configuração da página
st.set_page_config(
    page_title="Sistema de Controle e Comparação de Preços",
    page_icon="/content/logo-eqtl-app-teste2.png",
    layout="wide"
)

# Caminho das planilhas base e exceção (definidos manualmente no código)
CAMINHO_BASE = "/planilha_base.xlsx"
CAMINHO_EXCECAO = "/planilha_excecao.xlsx"

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
    "Valor/moeda objeto", "MAX_PU", "MIN_PU", "PU", "Resultado"
]

@st.cache_data
def load_base_planilha():
    try:
        return pd.read_excel(CAMINHO_BASE)
    except Exception:
        return None

@st.cache_data
def load_excecao_planilha():
    try:
        return pd.read_excel(CAMINHO_EXCECAO)
    except Exception:
        return None

def verificar_preco(row, base_df):
    material = row['Material']
    valor_proposto = row['Valor/moeda objeto']
    base_info = base_df[base_df['Equipamento'] == material]
    if not base_info.empty:
        preco_min = base_info['MIN_PU'].iloc[0]
        preco_max = base_info['MAX_PU'].iloc[0]
        return "✅ OK" if preco_min <= valor_proposto <= preco_max else "❌ Indevido"
    else:
        return "⚠️ Equipamento não encontrado"

def gerar_arquivo_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Resultado')
        writer.close()
    return output.getvalue()

def verificar_colunas(df, colunas_esperadas):
    colunas_presentes = set(df.columns)
    colunas_esperadas = set(colunas_esperadas)
    if colunas_presentes == colunas_esperadas:
        return True
    else:
        colunas_faltantes = colunas_esperadas - colunas_presentes
        colunas_extras = colunas_presentes - colunas_esperadas
        return False, colunas_faltantes, colunas_extras

def filtrar_excecoes(comparacao_df, excecao_df):
    df_agrupado = comparacao_df.copy()
    df_ex = excecao_df.copy()
    df_agrupado = df_agrupado[~df_agrupado['Material'].isin(df_ex['Nº de serviço'])].copy()
    return df_agrupado

def main():
    # Exibir logo
    try:
        st.sidebar.image("/content/GRUPO-EQUATORIAL-ENERGIA-LOGO_PADRAO_COR.png", width=400)
    except Exception:
        st.sidebar.info("🔹 Adicione um logo no diretório do aplicativo para exibição.")

    st.sidebar.title("📊 Menu")
    st.sidebar.info("Gerencie e valide os preços de equipamentos com base na planilha de referência.")
    st.title("Sistema de Controle e Comparação de Preços")
    st.write("Este sistema verifica se os preços fornecidos estão dentro dos valores permitidos pela base.")

    # Opção de atualizar as planilhas base e exceção
    st.sidebar.subheader("📂 Atualizar Planilha Base e Exceção")
    
    # Atualizar a planilha base
    st.sidebar.subheader("📂 Atualizar Planilha Base")
    new_base_file = st.sidebar.file_uploader("Carregar Nova Planilha Base (Excel)", type=["xlsx"])
    if new_base_file:
        new_base_df = pd.read_excel(new_base_file)
        new_base_df.to_excel(CAMINHO_BASE, index=False)
        st.sidebar.success("✅ Planilha base atualizada com sucesso!")

    # Atualizar a planilha de exceção
    st.sidebar.subheader("📂 Atualizar Planilha de Exceção")
    new_excecao_file = st.sidebar.file_uploader("Carregar Nova Planilha de Exceção (Excel)", type=["xlsx"])
    if new_excecao_file:
        new_excecao_df = pd.read_excel(new_excecao_file)
        new_excecao_df.to_excel(CAMINHO_EXCECAO, index=False)
        st.sidebar.success("✅ Planilha de exceção atualizada com sucesso!")

    # Carregar planilhas a partir dos caminhos configurados manualmente
    base_df = load_base_planilha()
    if base_df is None:
        st.error("⚠️ Nenhuma planilha base encontrada no caminho fornecido! Verifique o caminho e tente novamente.")
      
