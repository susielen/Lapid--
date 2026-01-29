import streamlit as st
import pandas as pd
import re

st.set_page_config(page_title="Conciliador Domínio", layout="wide")

st.title("🤖 Robô de Conciliação (Excel)")

# 1. O robô agora aceita arquivos .xlsx
arquivo = st.file_uploader("Suba o Razão do Domínio (Excel)", type=["xlsx"])

if arquivo:
    # Lendo o Excel (usando o motor openpyxl)
    df = pd.read_excel(arquivo, engine='openpyxl')
    
    banco_fornecedores = {}
    fornecedor_atual = None
    dados_acumulados = []

    # 2. O Robô Detetive limpa e organiza
    for _, linha in df.iterrows():
        # Identifica a linha que tem o nome do fornecedor (Coluna 'Data' diz "Conta:")
        if str(linha.iloc[0]).strip().startswith("Conta:"):
            if fornecedor_atual and dados_acumulados:
                banco_fornecedores[fornecedor_atual] = pd.DataFrame(dados_acumulados)
            
            # Pega o nome do fornecedor que geralmente está na coluna 5 ou 6
            fornecedor_atual = str(linha.iloc[5]) if pd.notna(linha.iloc[5]) else "Desconhecido"
            dados_acumulados = []
            continue
        
        # Verifica se a linha tem uma data válida para ser um movimento
        if pd.notna(linha.iloc[0]) and any(char.isdigit() for char in str(linha.iloc[0])):
            # Pega os valores de Débito e Crédito
            deb = float(linha.iloc[8]) if pd.notna(linha.iloc[8]) else 0
            cre = float(linha.iloc[9]) if pd.notna(linha.iloc[9]) else 0
            
            # Tenta achar o número da nota no histórico
            historico = str(linha.iloc[2])
            nfe = re.findall(r'NFe\s?(\d+)', historico)
            num_nota = nfe[0] if nfe else "S/N"
            
            dados_acumulados.append({
                "Data": linha.iloc[0],
                "Histórico": historico,
                "NF": num_nota,
                "Débito (Pago)": deb,
                "Crédito (Comprou)": cre
            })

    # Salva o último fornecedor
    if fornecedor_atual and dados_acumulados:
        banco_fornecedores[fornecedor_atual] = pd.DataFrame(dados_acumulados)

    # 3. Criando as Abas e Colunas lado a lado
    if banco_fornecedores:
        nomes = list(banco_fornecedores.keys())
        tabs = st.tabs(nomes)

        for i, nome in enumerate(nomes):
            with tabs[i]:
                st.subheader(f"🏢 {nome}")
                
                # Prepara o Razão e a Conciliação
                df_razao = banco_fornecedores[nome]
                df_conc = df_razao.groupby("NF").agg({
                    "Débito (Pago)": "sum",
                    "Crédito (Comprou)": "sum"
                }).reset_index()
                
                df_conc["Diferença"] = df_conc["Débito (Pago)"] - df_conc["Crédito (Comprou)"]
                df_conc["Status"] = df_conc["Diferença"].apply(lambda x: "✅ OK" if abs(x) < 0.01 else "🚩 Divergente")

                # Divide a tela em duas colunas (Razão | Conciliação)
                col_razao, col_espaco, col_conc = st.columns([1.5, 0.2, 1])
                
                with col_razao:
                    st.markdown("### 📄 Razão")
                    st.dataframe(df_razao, use_container_width=True, hide_index=True)
                
                # A col_espaco fica vazia para "pular" as colunas que você pediu
                
                with col_conc:
                    st.markdown("### ⚖️ Conciliação")
                    st.dataframe(df_conc, use_container_width=True, hide_index=True)
                    
                    # Resumo rápido embaixo da conciliação
                    st.info(f"Saldo Geral deste Fornecedor: R$ {df_conc['Diferença'].sum():,.2f}")
