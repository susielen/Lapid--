import streamlit as st
import pandas as pd
import re

st.set_page_config(page_title="Conciliador de Fornecedores", layout="wide")

st.title("📑 Conciliação por Fornecedor (Domínio)")

arquivo = st.file_uploader("Suba o arquivo Razão do Domínio (CSV)", type=["csv"])

if arquivo:
    # Lendo o arquivo
    df = pd.read_csv(arquivo, skip_blank_lines=True)
    
    # Dicionário para guardar os dados de cada fornecedor
    banco_fornecedores = {}
    fornecedor_atual = None
    dados_acumulados = []

    # --- PARTE 1: O Robô Detetive separa os dados ---
    for _, linha in df.iterrows():
        # Identifica a linha de "Conta:" que tem o nome do fornecedor
        if str(linha[0]).startswith("Conta:"):
            # Se já vínhamos guardando dados de outro fornecedor, salva antes de mudar
            if fornecedor_atual and dados_acumulados:
                banco_fornecedores[fornecedor_atual] = pd.DataFrame(dados_acumulados)
            
            # Pega o nome do novo fornecedor
            fornecedor_atual = str(linha[5]) if pd.notna(linha[5]) else str(linha[2])
            dados_acumulados = []
            continue
        
        # Se a linha tem data (formato AAAA-MM-DD), é um movimento
        if pd.notna(linha[0]) and re.match(r'\d{4}-\d{2}-\d{2}', str(linha[0])):
            debito = float(str(linha[8]).replace(',', '.')) if pd.notna(linha[8]) else 0
            credito = float(str(linha[9]).replace(',', '.')) if pd.notna(linha[9]) else 0
            
            # Tenta extrair o número da nota do histórico
            historico = str(linha[2])
            nfe = re.findall(r'NFe\s(\d+)', historico)
            num_nota = nfe[0] if nfe else "S/N"
            
            dados_acumulados.append({
                "Data": linha[0],
                "Histórico": historico,
                "NF": num_nota,
                "Débito (Pago)": debito,
                "Crédito (Comprou)": credito
            })
    
    # Salva o último fornecedor da lista
    if fornecedor_atual:
        banco_fornecedores[fornecedor_atual] = pd.DataFrame(dados_acumulados)

    # --- PARTE 2: Criando as abas e colocando lado a lado ---
    if banco_fornecedores:
        nomes_fornecedores = list(banco_fornecedores.keys())
        abas = st.tabs(nomes_fornecedores) # Cria uma aba para cada nome

        for i, nome in enumerate(nomes_fornecedores):
            with abas[i]:
                st.subheader(f"Fornecedor: {nome}")
                
                df_razao = banco_fornecedores[nome]
                
                # Criando a Conciliação (Resumo por Nota)
                df_conciliado = df_razao.groupby("NF").agg({
                    "Débito (Pago)": "sum",
                    "Crédito (Comprou)": "sum"
                }).reset_index()
                df_conciliado["Diferença"] = df_conciliado["Débito (Pago)"] - df_conciliado["Crédito (Comprou)"]
                df_conciliado["Status"] = df_conciliado["Diferença"].apply(lambda x: "✅ OK" if x == 0 else "🚩 Erro")

                # Juntando Razão + 3 Colunas Vazias + Conciliação
                espaco_vazio = pd.DataFrame({"": [""] * len(df_razao)})
                
                # Criando a visualização lado a lado usando colunas do Streamlit
                col_esq, col_dir = st.columns([1.5, 1]) # O Razão é maior que a Conciliação
                
                with col_esq:
                    st.write("**📄 Razão Detalhado**")
                    st.dataframe(df_razao, use_container_width=True)
                
                with col_dir:
                    st.write("**⚖️ Conciliação (Resumo)**")
                    st.dataframe(df_conciliado, use_container_width=True)
