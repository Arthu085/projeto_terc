import streamlit as st
import pandas as pd
import warnings
import openpyxl as op

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

caminho_arquivo = r"C:\Users\arthu\OneDrive\Documentos\Repositórios\projeto_usi_terc\Controle de Terceiros 2024 Atualizada copia - Copia.xlsm"
nome_planilha = "TABELA UNIFICADA 2024"
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

with tab2:
    st.subheader("Adicionar Entrega")
    fornecedores = ["ALBATROZ", "FELTRIN", "J.VIEIRA", "JULFER", "METALTECNICA", "METOLL",
                    "METROMAQ", "MIRANDA", "NAVY TOOLS", "NOBRE", "JR"]
    
    col1, col2 = st.columns(2)

    with col1:
        escolha_fornecedor = st.selectbox("Fornecedor:", fornecedores)
        op_numero = st.text_input("Digite o número da OP:")
        item_codigo = st.text_input("Digite o código do item:")
        item_nf_codigo = st.text_input("Digite o código NF do item:")
        try:
            diametro_numero = st.number_input("Digite o diâmetro do FP:", min_value=0.00)
        except:
            st.error("Por favor não digite valores negativos")
            diametro_numero = 0.00
        qtd_numero = st.number_input("Digite a quantidade:", min_value=0, max_value=10000)
        obs = st.text_input("Digite a observação:")

    with col2:
        try:
            valor_numero = st.number_input("Digite o preço:", min_value=0.00)
        except:
            st.error("Por favor não digite valores negativos")
            valor_numero = 0.00
        tipos_buchas = ["JA", "SH", "SDS", "SD", "SK", "SF", "E", "F", "J", "M", "-"]
        escolha_bucha = st.selectbox("Selecione o tipo de bucha:", tipos_buchas, index=tipos_buchas.index("-"))
        operacao_col = colunas[9]
        if operacao_col in df.columns:
            operacao_ordenados = sorted(df[operacao_col].dropna().astype(str).unique())
            operacao_ordenados.insert(0, "ADICIONAR OPERAÇÃO")
            escolha = st.selectbox("Operação:", operacao_ordenados)
            if escolha == "ADICIONAR OPERAÇÃO":
                operacao = st.text_input("Digite a operação:")
            else:
                operacao = escolha
        qtd_fix = int(st.number_input("Selecione a quantidade de FF:", min_value=0, max_value=100))
        data_coleta = st.date_input("Selecione a data da coleta:")
        data_prevista = st.date_input("Selecione a data de entrega prevista:")

        adicionar = st.button("Adicionar")
        if adicionar and all([escolha_fornecedor and op_numero == "" or op_numero and item_codigo and item_nf_codigo and diametro_numero == 0 or diametro_numero and qtd_numero
                               and valor_numero == 0 or valor_numero and escolha_bucha and operacao and qtd_fix == 0 or qtd_fix and data_coleta and data_prevista and obs == "" or obs]):
            new_data = {"FORNECEDOR": [escolha_fornecedor],
                        "OP": [op_numero],
                        "ITEN": [item_codigo],
                        "ITEM NF": [item_nf_codigo],
                        "DIÂMETRO FP": [diametro_numero],
                        "QTD": [qtd_numero],
                        "VALOR": [valor_numero],
                        "TIPO DE BUCHA": [escolha_bucha],
                        "OPERAÇÃO": [operacao],
                        "QTD FF": [qtd_fix],
                        "COLETA": [data_coleta],
                        "ENTREGA PREVISTA": [data_prevista],
                        "OBS": [obs]}
            
            new_df = pd.DataFrame(new_data)
            workbook = op.load_workbook(caminho_arquivo)
            sheet = workbook[nome_planilha]
            for row in sheet.iter_rows(min_row=sheet.max_row, max_row=sheet.max_row):
                if all(cell.value is None for cell in row):
                    row_index = row[0].row
                    break
            else:
                row_index = sheet.max_row + 1
            for index, row in new_df.iterrows():
                sheet.insert_rows(idx=row_index)  
                for j, value in enumerate(row.tolist()):
                    sheet.cell(row=row_index, column=j+1).value = value
            workbook.save(caminho_arquivo)

            st.success("Dados adicionados com sucesso!")
        elif adicionar:
            st.error("Preencha todos os campos obrigatórios!")