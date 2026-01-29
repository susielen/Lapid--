import streamlit as st
import pandas as pd
import re

st.set_page_config(page_title="Conciliador Pro", layout="wide")

st.title("🤖 Robô Conciliador Multi-Arquivos")
st.write("Suba o Razão do Domínio em **Excel** ou **CSV**")

# 1. O robô agora aceita os dois tipos!
arquivo = st.file_uploader("Escolha o arquivo", type=["xlsx", "csv"])

if arquivo:
    # Verifica qual o tipo do arquivo para saber como ler
    if arquivo.name.endswith('.csv'):
        df = pd.read_csv(arquivo, skip_blank_lines=True)
    else:
        df = pd.read_excel(arquivo, engine='openpyxl')
    
    banco_fornecedores = {}
    fornecedor_atual = None
    dados_acumulados = []

    # 2. Lógica para separar os fornecedores (O "Detetive")
    for _, linha in df.iterrows():
        primeira_celula = str(linha.iloc[0]).strip()
        
        if primeira_celula.startswith("Conta:"):
            if fornecedor_atual and dados_acumulados:
                banco_fornecedores[fornecedor_atual] = pd.DataFrame(dados_acumulados)
            
            # Pega o nome do fornecedor (ajustado para o padrão do Domínio)
            fornecedor_atual = str(linha.iloc[5]) if len(linha) > 5 and pd.notna(linha.iloc[5]) else str(linha.iloc[2])
            dados_acumulados = []
            continue
        
        # Se tem data, é movimento
        if pd.notna(linha.iloc[0]) and any(char.isdigit() for char in str(linha.iloc[0])):
            deb = float(str(linha.iloc[8]).replace(',', '.')) if pd.notna(linha.iloc[8]) else 0
            cre = float(str(linha.iloc[9]).replace(',', '.')) if pd.notna(linha.iloc[9]) else 0
            
            # Limpeza do histórico para pegar a NF
            hist = str(linha.iloc[2])
            nfe = re.findall(r'NFe\s?(\d+)', hist)
            num_nota = nfe[0] if nfe else "S/N"
            
            dados_acumulados.append({
                "Data": linha.iloc[0],
                "NF": num_nota,
                "Histórico": hist,
                "Débito (Pago)": deb,
                "Crédito (Comprou)": cre
            })

    # Salva o último do arquivo
    if fornecedor_atual and dados_acumulados:
        banco_fornecedores[fornecedor_atual] = pd.DataFrame(dados_acumulados)

    # 3. Criando as Abas e Colunas Lado a Lado
    if banco_fornecedores:
        nomes = list(banco_fornecedores.keys())
        abas = st.tabs(nomes)

        for i, nome in enumerate(nomes):
            with abas[i]:
                st.subheader(f"🏢 Fornecedor: {nome}")
                
                df_razao = banco_fornecedores[nome]
                
                # Criando a Conciliação (Resumo)
                df_conc = df_razao.groupby("NF").agg({
                    "Débito (Pago)": "sum",
                    "Crédito (Comprou)": "sum"
                }).reset_index()
                df_conc["Diferença"] = df_conc["Débito (Pago)"] - df_conc["Crédito (Comprou)"]
                df_conc["Status"] = df_conc["Diferença"].apply(lambda x: "✅ OK" if abs(x) < 0.01 else "🚩 Divergente")

                # Layout: Razão | Espaço | Conciliação
                col_esq, col_pulo, col_dir = st.columns([1.5, 0.2, 1])
                
                with col_esq:
                    st.markdown("**📄 Razão Detalhado**")
                    st.dataframe(df_razao, use_container_width=True, hide_index=True)
                
                # col_pulo fica vazia (são as 3 colunas de espaço que você pediu)
                
                with col_dir:
                    st.markdown("**⚖️ Conciliação Automática**")
                    st.dataframe(df_conc, use_container_width=True, hide_index=True)
                    
                    # Cartão de resumo
                    total_dif = df_conc["Diferença"].sum()
                    if abs(total_dif) < 0.01:
                        st.success(f"Saldo Total: R$ {total_dif:,.2f} - TUDO CERTO!")
                    else:
                        st.warning(f"Saldo Total: R$ {total_dif:,.2f} - VERIFICAR!")
