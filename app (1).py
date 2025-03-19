import streamlit as st
import pandas as pd
from io import BytesIO

# DEVE SER A PRIMEIRA INSTRUÇÃO DO SCRIPT:
st.set_page_config(
    page_title="Sistema de Controle e Comparação de Preços",
    page_icon="logo-eqtl-app-teste2.png",
    layout="wide"
)

# Injetar CSS para personalizar a altura do botão de dúvidas
st.markdown(
    """
    <style>
    /* Aplica apenas ao botão dentro do container com classe 'doubt-button' */
    div.doubt-button > button {
        height: 60px !important;
        font-size: 20px !important;
    }
    </style>
    """, unsafe_allow_html=True
)

# Inicializa a variável de sessão para controlar a exibição da mensagem
if "show_info" not in st.session_state:
    st.session_state.show_info = False

# Cria uma linha de header com duas colunas: uma para o botão e outra para o título
col1, col2 = st.columns([1, 8])
with col1:
    # Envolve o botão em um container com classe 'doubt-button'
    with st.container():
        st.markdown('<div class="doubt-button">', unsafe_allow_html=True)
        if st.button("❓", key="toggle_info_button"):
            st.session_state.show_info = not st.session_state.show_info
        st.markdown('</div>', unsafe_allow_html=True)
with col2:
    st.title("Sistema de Controle e Comparação de Preços")

# Alterna a exibição da mensagem de instrução
if st.session_state.show_info:
    try:
        with st.modal("Instruções de Uso"):
            st.write(
                "Para otimizar o uso das funcionalidades, por favor, carregue o arquivo CJI3 "
                "extraído do SAP com o layout BRP_RAW, utilizando o campo 'Drag and drop file here'. "
                "Certifique-se de que o formato das colunas permaneça inalterado e remova a linha amarela "
                "localizada na última linha do arquivo extraído da CJI3."
            )
    except Exception:
        st.info(
            "Para otimizar o uso das funcionalidades, por favor, carregue o arquivo CJI3 "
            "extraído do SAP com o layout BRP_RAW, utilizando o campo 'Drag and drop file here'. "
            "Certifique-se de que o formato das colunas permaneça inalterado e remova a linha amarela "
            "localizada na última linha do arquivo extraído da CJI3."
        )

# ---------------------------------------------------------
# O restante do código permanece inalterado:
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
    "Valor/moeda objeto", "MAX_PU", "MIN_PU", "PU", "Resultado"
]

@st.cache_data
def load_base_planilha():
    try:
        df = pd.read_excel(CAMINHO_BASE)
        return df
    except Exception as e:
        st.error(f"Erro ao tentar carregar a planilha base: {e}")
        return None

@st.cache_data
def load_excecao_planilha():
    try:
        return pd.read_excel(CAMINHO_EXCECAO)
    except Exception as e:
        st.error(f"Erro ao tentar carregar a planilha de exceção: {e}")
        return None

def safe_write(worksheet, row, col, value, cell_format):
    if pd.isna(value):
        worksheet.write(row, col, "", cell_format)
    elif isinstance(value, (int, float)):
        try:
            worksheet.write_number(row, col, value, cell_format)
        except TypeError:
            worksheet.write(row, col, str(value), cell_format)
    else:
        worksheet.write(row, col, str(value), cell_format)

def gerar_arquivo_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Resultado', header=False)
        workbook  = writer.book
        worksheet = writer.sheets['Resultado']
        worksheet.autofilter(0, 0, 0, len(df.columns)-1)
        for i, col in enumerate(df.columns):
            max_len = df[col].astype(str).map(len).max()
            worksheet.set_column(i, i, max_len + 2)
        header_format = workbook.add_format({
            'align': 'center',
            'valign': 'vcenter',
            'bg_color': '#003a63',
            'font_color': '#ffffff',
            'bold': True
        })
        cell_format = workbook.add_format({'align': 'center', 'valign': 'vcenter'})
        for col_num, header in enumerate(df.columns):
            worksheet.write(0, col_num, header, header_format)
        for row_num in range(1, len(df) + 1):
            for col_num, value in enumerate(df.iloc[row_num - 1]):
                safe_write(worksheet, row_num, col_num, value, cell_format)
        writer.close()
    return output.getvalue()

def filtrar_excecoes(comparacao_df, excecao_df):
    df_filtrado = comparacao_df.copy()
    df_ex = excecao_df.copy()
    df_filtrado = df_filtrado[~df_filtrado['Material'].isin(df_ex['Nº de serviço'])].copy()
    return df_filtrado

def main():
    try:
        st.sidebar.image("GRUPO-EQUATORIAL-ENERGIA-LOGO_PADRAO_COR.png", width=400)
    except Exception:
        st.sidebar.info("🔹 Adicione um logo no diretório do aplicativo para exibição.")

    st.sidebar.title("📊 Menu")
    st.sidebar.info("Gerencie e valide os preços de equipamentos com base na planilha de referência.")
    st.sidebar.subheader("📂 Atualizar Planilha Base")
    new_base_file = st.sidebar.file_uploader("Carregar Nova Planilha Base (Excel)", type=["xlsx"])
    if new_base_file:
        new_base_df = pd.read_excel(new_base_file)
        new_base_df.to_excel(CAMINHO_BASE, index=False)
        st.sidebar.success("✅ Planilha base atualizada com sucesso!")
    st.sidebar.subheader("📂 Atualizar Planilha de Exceção")
    new_excecao_file = st.sidebar.file_uploader("Carregar Nova Planilha de Exceção (Excel)", type=["xlsx"])
    if new_excecao_file:
        new_excecao_df = pd.read_excel(new_excecao_file)
        new_excecao_df.to_excel(CAMINHO_EXCECAO, index=False)
        st.sidebar.success("✅ Planilha de exceção atualizada com sucesso!")
    base_df = load_base_planilha()
    if base_df is None:
        st.error("⚠️ Nenhuma planilha base encontrada! Verifique o caminho e tente novamente.")
        return
    excecao_df = load_excecao_planilha()
    if excecao_df is None:
        st.error("⚠️ Nenhuma planilha de exceção encontrada! Verifique o caminho e tente novamente.")
        return
    st.subheader("📂 Carregar Planilha para Comparação")
    new_file = st.file_uploader("Escolha um arquivo Excel para comparação", type=["xlsx"])
    if new_file:
        try:
            new_df = pd.read_excel(new_file)
            new_df = filtrar_excecoes(new_df, excecao_df)
            new_df = new_df.dropna(subset=['Material'])
            df_agrupado = new_df.groupby(['Empresa', 'Elemento PEP', 'Material'], as_index=False).agg({
                'Qtd.total entrada': 'sum',
                'Valor/moeda objeto': 'sum'
            })
            df_agrupado['PU'] = (df_agrupado['Valor/moeda objeto'] / df_agrupado['Qtd.total entrada']).round(2)
        except Exception as e:
            st.error(f"Ocorreu um erro ao processar a planilha: {e}")
            return
        df_agrupado = pd.merge(
            df_agrupado,
            base_df[['Equipamento', 'DESC_MATERIAL', 'MAX_PU', 'MIN_PU']],
            left_on='Material',
            right_on='Equipamento',
            how='left'
        )
        df_agrupado.drop(columns=['Equipamento'], inplace=True)
        df_agrupado = df_agrupado[
            (df_agrupado['Qtd.total entrada'] != 0) & (df_agrupado['Valor/moeda objeto'] != 0)
        ]
        df_agrupado['Resultado'] = df_agrupado.apply(
            lambda row: "✅ OK" if pd.notnull(row['MIN_PU']) and pd.notnull(row['MAX_PU']) and row['MIN_PU'] <= row['PU'] <= row['MAX_PU'] 
            else ("❌ Indevido" if pd.notnull(row['MIN_PU']) and pd.notnull(row['MAX_PU']) else "⚠️ Equipamento não encontrado"), axis=1
        )
        final_columns = [
            "Empresa", "Elemento PEP", "Material", "DESC_MATERIAL", "Qtd.total entrada",
            "Valor/moeda objeto", "MAX_PU", "MIN_PU", "PU", "Resultado"
        ]
        df_agrupado = df_agrupado[final_columns]
        processed_df = df_agrupado.copy()
        processed_file = gerar_arquivo_excel(processed_df)
        st.subheader("📊 Resumo dos Resultados Agrupados")
        st.dataframe(processed_df)
        st.download_button(
            label="📥 Baixar Planilha Processada",
            data=processed_file,
            file_name="planilha_processada.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

if __name__ == '__main__':
    main()


