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
    """
    Escreve o valor na célula utilizando o método adequado.
    Se for numérico, utiliza write_number; caso contrário, write_string.
    Se o valor for NaN, escreve uma string vazia.
    """
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
    # Escreve o DataFrame sem cabeçalho (pois iremos reescrevê-lo com formatação)
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Resultado', header=False)
        workbook  = writer.book
        worksheet = writer.sheets['Resultado']

        # Adicionar filtro para as colunas existentes
        worksheet.autofilter(0, 0, 0, len(df.columns)-1)
        
        # Ajustar a largura de cada coluna baseada no conteúdo
        for i, col in enumerate(df.columns):
            max_len = df[col].astype(str).map(len).max()
            worksheet.set_column(i, i, max_len + 2)
        
        # Definir os formatos
        header_format = workbook.add_format({
            'align': 'center',
            'valign': 'vcenter',
            'bg_color': '#003a63',
            'font_color': '#ffffff',
            'bold': True
        })
        cell_format = workbook.add_format({'align': 'center', 'valign': 'vcenter'})

        # Reescrever o cabeçalho com o formato do cabeçalho
        for col_num, header in enumerate(df.columns):
            worksheet.write(0, col_num, header, header_format)
        
        # Reescrever as células de dados (a partir da linha 1)
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
    # Exibir logo
    try:
        st.sidebar.image("GRUPO-EQUATORIAL-ENERGIA-LOGO_PADRAO_COR.png", width=400)
    except Exception:
        st.sidebar.info("🔹 Adicione um logo no diretório do aplicativo para exibição.")

    st.sidebar.title("📊 Menu")
    st.sidebar.info("Gerencie e valide os preços de equipamentos com base na planilha de referência.")
    
    # Campo interativo com ícone de interrogação e mensagem de dica
    tooltip_html = """
    <div style="display: inline-block; margin-top: 10px; padding: 4px; border: 1px solid #ccc; border-radius: 4px; width: 24px; height: 24px; text-align: center; vertical-align: middle; cursor: pointer;" title="Para otimizar o uso das funcionalidades, por favor, carregue o arquivo CJI3 extraído do SAP com o layout BRP_RAW, utilizando o campo 'Drag and drop file here'. Assegure-se de que o formato das colunas permaneça inalterado, por favor exclua a linha amarela que fica na última linha do arquivo extraída da CJI3.">
      ?
    </div>
    """
    st.sidebar.markdown(tooltip_html, unsafe_allow_html=True)
    
    st.title("Sistema de Controle e Comparação de Preços")
    st.write("Este sistema verifica se os preços fornecidos estão dentro dos valores permitidos pela base.")
    
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

    # Carregar planilhas
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
            
            # Agrupar dados por Empresa, Elemento PEP e Material
            df_agrupado = new_df.groupby(['Empresa', 'Elemento PEP', 'Material'], as_index=False).agg({
                'Qtd.total entrada': 'sum',
                'Valor/moeda objeto': 'sum'
            })
            
            # Calcular o PU e arredondar para 2 casas decimais
            df_agrupado['PU'] = (df_agrupado['Valor/moeda objeto'] / df_agrupado['Qtd.total entrada']).round(2)
            
        except Exception as e:
            st.error(f"Ocorreu um erro ao processar a planilha: {e}")
            return
        
        # Merge com a planilha base para obter DESC_MATERIAL, MAX_PU e MIN_PU
        df_agrupado = pd.merge(
            df_agrupado,
            base_df[['Equipamento', 'DESC_MATERIAL', 'MAX_PU', 'MIN_PU']],
            left_on='Material',
            right_on='Equipamento',
            how='left'
        )
        df_agrupado.drop(columns=['Equipamento'], inplace=True)
        
        # Excluir linhas onde "Qtd.total entrada" ou "Valor/moeda objeto" sejam zero
        df_agrupado = df_agrupado[
            (df_agrupado['Qtd.total entrada'] != 0) & (df_agrupado['Valor/moeda objeto'] != 0)
        ]
        
        # Criar a coluna Resultado comparando PU com MIN_PU e MAX_PU
        df_agrupado['Resultado'] = df_agrupado.apply(
            lambda row: "✅ OK" if pd.notnull(row['MIN_PU']) and pd.notnull(row['MAX_PU']) and row['MIN_PU'] <= row['PU'] <= row['MAX_PU'] 
            else ("❌ Indevido" if pd.notnull(row['MIN_PU']) and pd.notnull(row['MAX_PU']) else "⚠️ Equipamento não encontrado"), axis=1
        )
        
        # Reordenar as colunas conforme solicitado
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
