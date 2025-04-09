import streamlit as st
import pandas as pd
import os
import io
from datetime import datetime

def salvar_arquivo_temporario(uploaded_file, tipo):
    extensao = os.path.splitext(uploaded_file.name)[-1]
    caminho = f"temp_{tipo}{extensao}"
    with open(caminho, "wb") as f:
        f.write(uploaded_file.getbuffer())
    return caminho

def carregar_planilha(caminho):
    return pd.read_excel(caminho)

def gerar_arquivo_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Resultado')
    output.seek(0)
    return output

@st.cache_data

def load_base_planilha(path):
    return carregar_planilha(path)

@st.cache_data

def load_excecao_planilha(path):
    return carregar_planilha(path)

def main():
    st.title("Comparador de Preços por Material")

    uploaded_base = st.file_uploader("Selecione a planilha BASE", type=[".xls", ".xlsx"], key="base")
    uploaded_excecao = st.file_uploader("Selecione a planilha de EXCEÇÃO", type=[".xls", ".xlsx"], key="excecao")

    if uploaded_base:
        caminho_base = salvar_arquivo_temporario(uploaded_base, "base")
        load_base_planilha.clear()
        df_base = load_base_planilha(caminho_base)

    else:
        df_base = None

    if uploaded_excecao:
        caminho_excecao = salvar_arquivo_temporario(uploaded_excecao, "excecao")
        load_excecao_planilha.clear()
        df_excecao = load_excecao_planilha(caminho_excecao)
    else:
        df_excecao = None

    if df_base is not None:
        df = df_base.copy()

        df["PU"] = df["Valor/moeda objeto"] / df["Qtd.total entrada"]

        colunas_agrupamento = ["Empresa", "Elemento PEP", "Material", "DESC_MATERIAL"]

        df_agrupado = df.groupby(colunas_agrupamento).agg({
            "Qtd.total entrada": "sum",
            "Valor/moeda objeto": "sum",
            "PU": ["max", "min"]
        }).reset_index()

        df_agrupado.columns = colunas_agrupamento + ["Qtd.total entrada", "Valor/moeda objeto", "MAX_PU", "MIN_PU"]

        df = df.merge(df_agrupado, on=colunas_agrupamento, how="left")

        df["Resultado"] = df.apply(
            lambda row: ("ACIMA" if row["PU"] > row["MAX_PU"] else "ABAIXO" if row["PU"] < row["MIN_PU"] else "OK"), axis=1
        )

        if df_excecao is not None:
            df = df.merge(df_excecao, on="Material", how="left", suffixes=("", "_ex"))
            df["Resultado"] = df.apply(
                lambda row: "EXCEÇÃO" if row.get("Excecao", "") == "SIM" else row["Resultado"], axis=1
            )

        df_agrupado = df[colunas_agrupamento + ["Qtd.total entrada", "Valor/moeda objeto", "PU", "MAX_PU", "MIN_PU", "Resultado"]].copy()

        df_agrupado = df_agrupado.drop_duplicates()

        df_agrupado["DIF"] = df_agrupado.apply(
            lambda row: (
                round(row["PU"] - row["MAX_PU"], 2) if row["PU"] > row["MAX_PU"]
                else round(row["MIN_PU"] - row["PU"], 2) if row["PU"] < row["MIN_PU"]
                else None
            ) if pd.notnull(row["PU"]) and pd.notnull(row["MAX_PU"]) and pd.notnull(row["MIN_PU"])
            else None,
            axis=1
        )

        df_agrupado["% DIF"] = df_agrupado.apply(
            lambda row: (
                f"{(row['DIF'] / row['PU'] * 100):.2f}%" if pd.notnull(row["DIF"]) and row["PU"] > row["MAX_PU"] and row["PU"] != 0
                else f"{(row['DIF'] / row['MIN_PU'] * 100):.2f}%" if pd.notnull(row["DIF"]) and row["PU"] < row["MIN_PU"] and row["MIN_PU"] != 0
                else None
            ), axis=1
        )

        for col in ["Valor/moeda objeto", "PU", "MAX_PU", "MIN_PU", "DIF"]:
            df_agrupado[col] = df_agrupado[col].apply(lambda x: f"{x:,.2f}" if pd.notnull(x) else "")

        final_columns = [
            "Empresa", "Elemento PEP", "Material", "DESC_MATERIAL", "Qtd.total entrada",
            "Valor/moeda objeto", "PU", "MAX_PU", "MIN_PU", "DIF", "% DIF", "Resultado"
        ]
        df_agrupado = df_agrupado[final_columns]

        processed_df = df_agrupado.copy()
        processed_file = gerar_arquivo_excel(processed_df)

        st.success("Processamento concluído.")
        st.dataframe(df_agrupado, use_container_width=True)

        st.download_button(
            label="📥 Baixar planilha processada",
            data=processed_file,
            file_name=f"resultado_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

if __name__ == "__main__":
    main()



