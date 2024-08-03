import streamlit as st
import pandas as pd
import openpyxl as op

caminho_arquivo = r"C:\Users\arthu\OneDrive\Documentos\Repositórios\projeto_usi_terc\Controle de Terceiros 2024 Atualizada copia - Copia.xlsm"
df = pd.read_excel(caminho_arquivo, engine="openpyxl")

st.title("Usinagem de Terceiro")

tab1, tab2, tab3, tab4 = st.tabs(["Entregas em Aberto", "Adicionar Entregas", "Requisitar O.C", "Entregas Realizadas"])



with tab1:
    st.subheader("Entregas em Aberto")
    escolha = st.selectbox("Fornecedor", df['FORNECEDOR'].unique())
    df_fornecedor = df[df['FORNECEDOR'] == escolha]
    if "ENTREGA REALIZADA" in df.columns == "":
        entregas_em_aberto = df[df_fornecedor['Entrega Prevista'].isna()]
        st.write(entregas_em_aberto)