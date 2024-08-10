import streamlit as st
import pandas as pd
import openpyxl as op

pagina = st.set_page_config(page_title='Terceiro',
                    layout="wide")

caminho_arquivo = r"C:\Users\arthu\OneDrive\Documentos\Repositórios\projeto_usi_terc\Controle de Terceiros 2024 Atualizada copia - Copia.xlsm"
nome_planilha = "TABELA UNIFICADA 2024"
df = pd.read_excel(caminho_arquivo, sheet_name=nome_planilha, skiprows=1)

st.title("Usinagem de Terceiro")

try:
    df['OP'] = df['OP'].astype(int) 
except:
    df['OP'] = df['OP'].astype(str) 

try:
    df['DIÂMETRO FP'] = df['DIÂMETRO FP'].astype(int) 
except:
    df['DIÂMETRO FP'] = df['DIÂMETRO FP'].astype(str) 

df['COLETA'] = df['COLETA'].dt.strftime('%d/%m/%Y')
df['ENTREGA REALIZADA'] = df['ENTREGA REALIZADA'].dt.strftime('%d/%m/%Y')
df['ENTREGA PREVISTA'] = df['ENTREGA PREVISTA'].dt.strftime('%d/%m/%Y')

tab1, tab2, tab3, tab4 = st.tabs(["Entregas em Aberto", "Adicionar Entregas", "Requisitar O.C", "Entregas Realizadas"])

with tab1:
    st.subheader("Entregas em Aberto")
    fornecedor_col = 'FORNECEDOR'

    if fornecedor_col in df.columns:
        fornecedores_ordenados = sorted(df[fornecedor_col].dropna().astype(str).unique())
        fornecedores_ordenados.insert(0, "TODOS")
        escolha = st.selectbox("Fornecedor:", fornecedores_ordenados)
        if escolha == "TODOS":
            df_fornecedor = df.dropna(subset=[fornecedor_col], how="all")
        else:
            df_fornecedor = df[df[fornecedor_col].astype(str) == escolha]
        entrega_realizada_col = 'ENTREGA REALIZADA'
        if entrega_realizada_col in df_fornecedor.columns:
            entregas_em_aberto = df_fornecedor[df_fornecedor[entrega_realizada_col].isna()]
            st.write(entregas_em_aberto)
        else:
            st.error(f"A planilha não possui a coluna '{entrega_realizada_col}'.")
    else:
        st.error(f"A planilha não possui a coluna '{fornecedor_col}'.")

with tab4:
    st.subheader("Entregas Realizadas")
    entregas_realizadas_col = 'ENTREGA REALIZADA'
    if entrega_realizada_col in df.columns:
        df_entrega_realziada = df[df[entregas_realizadas_col].notnull()]
        st.write(df_entrega_realziada)

with tab2:
    st.subheader("Adicionar Entrega")
    
    col1, col2,col3 = st.columns([1,1,1])

    with col1:
        fornecedores_ordenados = sorted(df['FORNECEDOR'].dropna().astype(str).unique())
        escolha_fornecedor = st.selectbox("Fornecedor:", fornecedores_ordenados)
        op_numero = st.text_input("Digite o número da OP:")
        item_codigo = st.text_input("Digite o código do item:")
        item_nf_codigo = st.text_input("Digite o código NF do item:")

    with col2:
        try:
            diametro_fp = st.number_input("Digite o diâmetro do FP:", min_value=0.00)
        except:
            st.error("Por favor não digite valores negativos")
            diametro_fp = 0.00
        qtd_pecas = st.number_input("Digite a quantidade:", min_value=0, max_value=10000)
        obs = st.text_input("Digite a observação:")
        try:
            preco = st.number_input("Digite o preço:", min_value=0.00)
        except:
            st.error("Por favor não digite valores negativos")
            preco = 0.00
        buchas_ordenadas = sorted(df['TIPO DE BUCHA'].dropna().astype(str).unique())
        escolha_bucha = st.selectbox("Selecione o tipo de bucha:", buchas_ordenadas)

    with col3:
        operacao_col = 'OPERAÇÃO'
        if operacao_col in df.columns:
            operacao_ordenados = sorted(df[operacao_col].dropna().astype(str).unique())
            operacao_ordenados.insert(0, "ADICIONAR OPERAÇÃO")
            escolha = st.selectbox("Operação:", operacao_ordenados)
            if escolha == "ADICIONAR OPERAÇÃO":
                operacao = st.text_input("Digite a operação:")
            else:
                operacao = escolha
        qtd_fix = int(st.number_input("Selecione a quantidade de FF:", min_value=0, max_value=100))
        data_coleta = st.date_input("Selecione a data da coleta:", format="DD/MM/YYYY")
        data_prevista = st.date_input("Selecione a data de entrega prevista:", format="DD/MM/YYYY")
        
        if st.button('Adicionar Entrega'):
            try:
                nova_linha = {'FORNECEDOR': escolha_fornecedor, 
                            'OP': op_numero,
                            'ITEM': item_codigo,
                            'ITEM NF': item_nf_codigo,
                            'DIÂMETRO FP': diametro_fp,
                            'QTD': qtd_pecas,
                            'VALOR': preco,
                            '': 0,
                            'TIPO DE BUCHA': escolha_bucha,
                            'OPERAÇÃO': operacao,
                            'QTD FF': qtd_fix,
                            'COLETA': data_coleta,
                            'ENTREGA PREVISTA': data_prevista}
                df = pd.concat([df, pd.DataFrame([nova_linha])], ignore_index=True)
                with pd.ExcelWriter(caminho_arquivo, mode='a', if_sheet_exists='overlay') as writer:
                    df.to_excel(writer, sheet_name=nome_planilha, index=False)
                st.success('Entrega adcionada')
            except Exception as e:
                st.error(f'Erro ao adicionar a entrega: {e}')