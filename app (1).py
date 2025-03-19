import streamlit as st
import pandas as pd
from io import BytesIO

# Configuração da página
st.set_page_config(
    page_title="Sistema de Controle e Comparação de Preços",
    page_icon="logo-eqtl-app-teste2.png",
    layout="wide"
)

# Caminhos das planilhas
CAMINHO_BASE = "planilha_base.xlsx"
CAMINHO_EXCECAO = "planilha_excecao.XLSX"

# Colunas esperadas
COLUNAS_ESPERADAS_BASE = ["EMPRESA", "Equipamento", "DESC_MATERIAL", "MAX_PU", "MIN_PU"]
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
COLUNAS_PROCESSADAS = [
    "Empresa", "Elemento PEP", "Material", "DESC_MATERIAL", "Qtd.total entrada",
    "Valor/moeda objeto", "MAX_PU", "MIN_PU", "PU", "Resultado"
]

@st.cache_data
def load_base_planilha():
    try:
        return pd.read_excel(CAMINHO_BASE)
    except Exception as e:
        st.error(f"Erro ao carregar a planilha base: {e}")
        return None

@st.cache_data
def load_excecao_planilha():
    try:
        return pd.read_excel(CAMINHO_EXCECAO)
    except Exception as e:
        st.error(f"Erro ao carregar a planilha de exceção: {e}")
        return None

def verificar_preco(row, base_df):
    material = row['Material']
    valor_proposto = row['Valor/moeda objeto']
    base_info = base_df[base_df['Equipamento'] == material]
    if not base_info.empty:
        preco_min = base_info['MIN_PU'].iloc[0]
        preco_max = base_info['MAX_PU'].iloc[0]
        return "✅ OK" if preco_min <= valor_proposto <= preco_max else "❌ Indevido"
    return "⚠️ Equipamento não encontrado"

def gerar_arquivo_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Resultado')
        writer.close()
    return output.getvalue()

def filtrar_excecoes(comparacao_df, excecao_df):
    df_agrupado = comparacao_df.copy()
    df_ex = excecao_df.copy()
    df_agrupado = df_agrupado[~df_agrupado['Material'].isin(df_ex['Nº de serviço'])].copy()
    return df_agrupado

def main():
    # Customização de CSS
    st.markdown("""
        <style>
        .stButton>button {
            background-color: #004080;
            color: white;
            border-radius: 5px;
            padding: 8px 16px;
        }
        .stDataFrame {
            border: 1px solid #004080;
            border-radius: 5px;
        }
        </style>
    """, unsafe_allow_html=True)

    # Exibir logo no sidebar
    try:
        st.sidebar.image("GRUPO-EQUATORIAL-ENERGIA-LOGO_PADRAO_COR.png", width=200)
    except Exception:
        st.sidebar.info("🔹 Adicione um logo no diretório para exibição.")

    # Menu lateral
    st.sidebar.title("📊 Menu")
    st.sidebar.markdown("---")
    st.sidebar.subheader("📂 Atualizar Planilhas")
    new_base_file = st.sidebar.file_uploader("Carregar Nova Planilha Base (Excel)", type=["xlsx"], key="base")
    if new_base_file:
        new_base_df = pd.read_excel(new_base_file)
        new_base_df.to_excel(CAMINHO_BASE, index=False)
        st.sidebar.success("✅ Planilha base atualizada com sucesso!")

    new_excecao_file = st.sidebar.file_uploader("Carregar Nova Planilha de Exceção (Excel)", type=["xlsx"], key="excecao")
    if new_excecao_file:
        new_excecao_df = pd.read_excel(new_excecao_file)
        new_excecao_df.to_excel(CAMINHO_EXCECAO, index=False)
        st.sidebar.success("✅ Planilha de exceção atualizada com sucesso!")

    st.sidebar.markdown("---")
    st.sidebar.subheader("📂 Carregar para Comparação")
    new_file = st.sidebar.file_uploader("Escolha um arquivo Excel para comparação", type=["xlsx"], key="comparacao")

    # Título principal
    st.title("Sistema de Controle e Comparação de Preços")
    st.write("Verifique se os preços estão dentro dos valores permitidos pela base.")

    # Carregar planilhas base e exceção
    base_df = load_base_planilha()
    if base_df is None:
        st.error("⚠️ Planilha base não encontrada! Verifique o caminho.")
        return

    excecao_df = load_excecao_planilha()
    if excecao_df is None:
        st.error("⚠️ Planilha de exceção não encontrada! Verifique o caminho.")
        return

    # Processar nova planilha
    if new_file:
        with st.spinner("Processando a planilha..."):
            try:
                new_df = pd.read_excel(new_file)
                new_df = filtrar_excecoes(new_df, excecao_df)
                new_df = new_df.dropna(subset=['Material'])
                new_df['Resultado'] = new_df.apply(lambda row: verificar_preco(row, base_df), axis=1)

                # Agrupar dados
                df_agrupado = new_df.groupby(['Empresa', 'Elemento PEP', 'Material'], as_index=False).agg({
                    'Qtd.total entrada': 'sum',
                    'Valor/moeda objeto': 'sum',
                    'Resultado': 'first'
                })
                df_agrupado['PU'] = (df_agrupado['Valor/moeda objeto'] / df_agrupado['Qtd.total entrada']).round(2)

                # Merge com a base
                df_agrupado = pd.merge(
                    df_agrupado,
                    base_df[['Equipamento', 'DESC_MATERIAL', 'MAX_PU', 'MIN_PU']],
                    left_on='Material',
                    right_on='Equipamento',
                    how='left'
                )
                df_agrupado.drop(columns=['Equipamento'], inplace=True)

                # Reordenar colunas
                final_columns = [
                    "Empresa", "Elemento PEP", "Material", "DESC_MATERIAL", "Qtd.total entrada",
                    "Valor/moeda objeto", "MAX_PU", "MIN_PU", "PU", "Resultado"
                ]
                processed_df = df_agrupado[final_columns]
                processed_file = gerar_arquivo_excel(processed_df)

            except Exception as e:
                st.error(f"Erro ao processar a planilha: {e}")
                return

        # Exibir resultados em colunas
        col1, col2 = st.columns([3, 1])
        with col1:
            st.subheader("📊 Resumo dos Resultados")
            st.dataframe(processed_df)
        with col2:
            st.subheader("📥 Download")
            st.download_button(
                label="Baixar Planilha Processada",
                data=processed_file,
                file_name="planilha_processada.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

if __name__ == '__main__':
    main()

