import streamlit as st
import pandas as pd
import warnings

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

caminho_arquivo = r"C:\Users\arthu\OneDrive\Documentos\Repositórios\projeto_usi_terc\Controle de Terceiros 2024 Atualizada copia - Copia.xlsm"
nome_planilha = "TABELA UNIFICADA 2024  "
df = pd.read_excel(caminho_arquivo, sheet_name=nome_planilha, engine="openpyxl", skiprows=2)

novo_cabecalho = ["FORNECEDOR", "OP", "ITEM", "ITEM NF", "DIÂMETRO FP", "QTD", "VALOR", "TOTAL", 
"TIPO DE BUCHA", "OPERAÇÃO", "QTD FF", "COLETA", "ENTREGA PREVISTA", "ENTREGA REALIZADA", "OBS", "NF/OC"]
df.columns = novo_cabecalho

st.title("Usinagem de Terceiro")

colunas = (df.columns.tolist())
tab1, tab2, tab3, tab4 = st.tabs(["Entregas em Aberto", "Adicionar Entregas", "Requisitar O.C", "Entregas Realizadas"])



with tab1:
    st.subheader("Entregas em Aberto")
    fornecedor_col = colunas[0]
    if fornecedor_col in df.columns:
        fornecedores_ordenados = sorted(df[fornecedor_col].dropna().astype(str).unique())
        fornecedores_ordenados.insert(0, "TODOS")
        escolha = st.selectbox("Fornecedor:", fornecedores_ordenados)
        if escolha == "TODOS":
            colunas_filtro = colunas[0]
            df_fornecedor = df.dropna(subset=colunas_filtro, how="all")
        else:
            df_fornecedor = df[df[fornecedor_col].astype(str) == escolha]
        entrega_realizada_col = colunas[13]
        if entrega_realizada_col in df_fornecedor.columns:
            entregas_em_aberto = df_fornecedor[df_fornecedor[entrega_realizada_col].isna()]
            if "nan" not in fornecedor_col:
                st.write(entregas_em_aberto)
        else:
            st.error(f"A planilha não possui a coluna '{entrega_realizada_col}'.")
    else:
        st.error(f"A planilha não possui a coluna '{fornecedor_col}'.")

with tab4:
    st.subheader("Entregas Realizadas")
    entregas_realizadas_col = colunas[13]
    if entrega_realizada_col in df.columns:
        df_entrega_realziada = df[df[entregas_realizadas_col].notnull()]
        st.write(df_entrega_realziada)