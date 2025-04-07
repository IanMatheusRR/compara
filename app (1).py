import streamlit as st
import pandas as pd
from io import BytesIO

# DEVE SER A PRIMEIRA INSTRUÇÃO DO SCRIPT:
st.set_page_config(
    page_title="Sistema de Controle e Comparação de Preços",
    page_icon="logo-eqtl-app-teste2.png",
    layout="wide"
)

# Função auxiliar para verificar se as colunas do DataFrame são as esperadas
def verificar_colunas(df, colunas_esperadas):
    df_cols = set(df.columns.tolist())
    esperado = set(colunas_esperadas)
    faltando = esperado - df_cols
    extras = df_cols - esperado
    return faltando, extras

# Defina uma lista de códigos autorizados
CODIGOS_AUTORIZADOS = ["E3719", "U8877", "T667788"]

# Inicializa a variável de sessão para controlar a exibição da mensagem
if "show_info" not in st.session_state:
    st.session_state.show_info = False

# Botão de dúvidas que alterna a exibição da mensagem
if st.button("❓", key="toggle_info_button"):
    st.session_state.show_info = not st.session_state.show_info

# Exibe a mensagem de instrução se show_info for True
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

# Caminho das planilhas base e exceção (definidos manualmente no código)
CAMINHO_BASE = "planilha_base.xlsx"
CAMINHO_EXCECAO = "planilha_excecao.XLSX"

# Listas de colunas esperadas
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

# A ordem final das colunas processadas será:
# ["Empresa", "Elemento PEP", "Material", "DESC_MATERIAL", "Qtd.total entrada",
#  "Valor/moeda objeto", "PU", "MAX_PU", "MIN_PU", "Resultado"]
COLUNAS_PROCESSADAS = [
    "Empresa", "Elemento PEP", "Material", "DESC_MATERIAL", "Qtd.total entrada",
    "Valor/moeda objeto", "PU", "MAX_PU", "MIN_PU", "Resultado"
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
        
        # Formato padrão para cabeçalho
        header_format = workbook.add_format({
            'align': 'center',
            'valign': 'vcenter',
            'bg_color': '#003a63',
            'font_color': '#ffffff',
            'bold': True
        })
        # Formatos customizados para colunas específicas
        header_format_max = workbook.add_format({
            'align': 'center',
            'valign': 'vcenter',
            'bg_color': '#af0000',
            'font_color': '#ffffff',
            'bold': True
        })
        header_format_min = workbook.add_format({
            'align': 'center',
            'valign': 'vcenter',
            'bg_color': '#0070c0',
            'font_color': '#ffffff',
            'bold': True
        })
        header_format_result = workbook.add_format({
            'align': 'center',
            'valign': 'vcenter',
            'bg_color': '#9bbb59',
            'font_color': '#ffffff',
            'bold': True
        })
        
        cell_format = workbook.add_format({'align': 'center', 'valign': 'vcenter'})
        
        # Reescrever o cabeçalho com formatação customizada para colunas específicas
        for col_num, header in enumerate(df.columns):
            if header == "MAX_PU":
                fmt = header_format_max
            elif header == "MIN_PU":
                fmt = header_format_min
            elif header == "Resultado":
                fmt = header_format_result
            else:
                fmt = header_format
            worksheet.write(0, col_num, header, fmt)
        
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
    
    st.title("Sistema de Controle e Comparação de Preços")
    st.write("Este sistema verifica se os preços fornecidos estão dentro dos valores permitidos pela base.")
    
    # Atualizar a planilha base com verificação de colunas e código
    st.sidebar.subheader("📂 Atualizar Planilha Base")
    codigo_base = st.sidebar.text_input("Insira sua matrícula para atualizar a planilha base", type="default")
    new_base_file = st.sidebar.file_uploader("Carregar Nova Planilha Base (Excel)", type=["xlsx"], key="base_file")
    if new_base_file:
        new_base_df = pd.read_excel(new_base_file)
        faltando, extras = verificar_colunas(new_base_df, COLUNAS_ESPERADAS_BASE)
        if faltando or extras:
            st.sidebar.error(
                f"O arquivo base possui colunas incorretas!\nFaltando: {list(faltando)}\nExtras: {list(extras)}"
            )
        elif codigo_base not in CODIGOS_AUTORIZADOS:
            st.sidebar.error("Você não tem permissão para alterar")
        else:
            new_base_df.to_excel(CAMINHO_BASE, index=False)
            load_base_planilha.clear()  # Limpa o cache para forçar recarregamento
            st.sidebar.success("✅ Planilha base atualizada com sucesso!")
    
    # Atualizar a planilha de exceção com verificação de código
    st.sidebar.subheader("📂 Atualizar Planilha de Exceção")
    codigo_excecao = st.sidebar.text_input("Insira sua matrícula para atualizar a planilha de exceção", type="default", key="excecao_code")
    new_excecao_file = st.sidebar.file_uploader("Carregar Nova Planilha de Exceção (Excel)", type=["xlsx"], key="excecao_file")
    if new_excecao_file:
        new_excecao_df = pd.read_excel(new_excecao_file)
        if codigo_excecao not in CODIGOS_AUTORIZADOS:
            st.sidebar.error("Você não tem permissão para alterar")
        else:
            new_excecao_df.to_excel(CAMINHO_EXCECAO, index=False)
            load_excecao_planilha.clear()
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
    new_file = st.file_uploader("Escolha um arquivo Excel para comparação", type=["xlsx"], key="comp_file")
    if new_file:
        try:
            new_df = pd.read_excel(new_file)
            faltando, extras = verificar_colunas(new_df, COLUNAS_ESPERADAS_COMPARACAO)
            if faltando or extras:
                st.error(
                    f"O arquivo de comparação possui colunas incorretas!\nFaltando: {list(faltando)}\nExtras: {list(extras)}"
                )
                return
            # Excluir linhas onde 'Elemento PEP' termina com ".D"
            new_df = new_df[~new_df["Elemento PEP"].astype(str).str.strip().str.endswith(".D")]
            new_df = filtrar_excecoes(new_df, excecao_df)
            new_df = new_df.dropna(subset=['Material'])
            
            # Agrupar dados por Empresa, Elemento PEP e Material
            df_agrupado = new_df.groupby(['Empresa', 'Elemento PEP', 'Material'], as_index=False).agg({
                'Qtd.total entrada': 'sum',
                'Valor/moeda objeto': 'sum'
            })
            df_agrupado['PU'] = (df_agrupado['Valor/moeda objeto'] / df_agrupado['Qtd.total entrada']).round(2)
        except Exception as e:
            st.error(f"Ocorreu um erro ao processar a planilha: {e}")
            return
        
        # Atualize a ordem das colunas: troque "PU" com "MAX_PU"
        # Primeiro, mescle com os dados da planilha base
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
        
        # Nova lógica para a coluna Resultado:
        # Se PU > MAX_PU: "Acima do máximo"
        # Se PU < MIN_PU: "Abaixo do mínimo"
        # Caso contrário: "OK"
        df_agrupado['Resultado'] = df_agrupado.apply(
            lambda row: ("Acima do máximo" if row['PU'] > row['MAX_PU'] 
                         else "Abaixo do mínimo" if row['PU'] < row['MIN_PU'] 
                         else "OK") if pd.notnull(row['MIN_PU']) and pd.notnull(row['MAX_PU'])
                         else "⚠️ Equipamento não encontrado", axis=1
        )
        
        # Reordenar as colunas: trocar "PU" com "MAX_PU"
        final_columns = [
            "Empresa", "Elemento PEP", "Material", "DESC_MATERIAL", "Qtd.total entrada",
            "Valor/moeda objeto", "PU", "MAX_PU", "MIN_PU", "Resultado"
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
