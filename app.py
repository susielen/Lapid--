import streamlit as st
import pandas as pd
import io

# Título do seu site
st.set_page_config(page_title="Conciliador Domínio", page_icon="🤖")
st.title("🤖 Conciliador de Fornecedores")
st.write("Suba o arquivo CSV do seu Razão e eu organizo tudo!")

# Botão para colocar o arquivo
arquivo = st.file_uploader("Arraste o arquivo Razão aqui", type=["csv", "xlsx"])

if arquivo is not None:
    if arquivo.name.endswith('.csv'):
        df = pd.read_csv(arquivo, skiprows=6)
    else:
        df = pd.read_excel(arquivo, skiprows=6)

if arquivo is not None:
    # O robô lê o arquivo
    df = pd.read_csv(arquivo, skiprows=6)
    df.columns = df.columns.str.strip()

    dados = []
    fornecedor_atual = "Não Identificado"

    for i, linha in df.iterrows():
        hist = str(linha['Histórico'])
        data_col = str(linha['Data'])
        
        # Se achar a linha da Conta, ele guarda o nome do Fornecedor
        if 'Conta:' in data_col or 'Conta:' in hist:
            fornecedor_atual = hist.strip()
            continue
        
        # Pega os valores de Débito e Crédito
        d = pd.to_numeric(linha['Débito'], errors='coerce')
        c = pd.to_numeric(linha['Crédito'], errors='coerce')
        
        if pd.notna(d) or pd.notna(c):
            dados.append({
                'Fornecedor': fornecedor_atual,
                'Débito': d if pd.notna(d) else 0,
                'Crédito': c if pd.notna(c) else 0
            })

    if dados:
        df_limpo = pd.DataFrame(dados)
        # Soma tudo por fornecedor
        resumo = df_limpo.groupby('Fornecedor').agg({'Débito': 'sum', 'Crédito': 'sum'}).reset_index()
        resumo['Saldo Final'] = resumo['Crédito'] - resumo['Débito']

        st.success("Tudo pronto! Veja o resumo abaixo:")
        st.dataframe(resumo)

        # Prepara o botão de baixar para Excel
        saida = io.BytesIO()
        with pd.ExcelWriter(saida, engine='openpyxl') as writer:
            resumo.to_excel(writer, index=False)
        
        st.download_button("📥 Baixar em Excel", data=saida.getvalue(), file_name="resultado.xlsx")
