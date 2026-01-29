import streamlit as st
import pandas as pd
import re
from io import BytesIO

# 1. Configuração inicial do Robô
st.set_page_config(page_title="Conciliador Mestre", layout="wide")

st.title("🤖 Robô Conciliador (Versão Ultra Blindada)")
st.write("Suba o arquivo do Domínio e eu farei a mágica!")

arquivo = st.file_uploader("Suba o arquivo (XLSX ou CSV)", type=["xlsx", "csv"])

if arquivo:
    try:
        # 2. Lendo o papel (arquivo)
        if arquivo.name.endswith('.csv'):
            df_bruto = pd.read_csv(arquivo, header=None)
        else:
            df_bruto = pd.read_excel(arquivo, engine='openpyxl', header=None)
        
        # 3. Procurando o nome da Empresa
        nome_empresa = "EMPRESA NÃO IDENTIFICADA"
        for i in range(min(20, len(df_bruto))):
            txt = str(df_bruto.iloc[i, 0])
            if "Empresa:" in txt or "EMPRESA:" in txt.upper():
                nome_empresa = str(df_bruto.iloc[i, 2]) if pd.notna(df_bruto.iloc[i, 2]) else nome_empresa
                break

        banco_fornecedores = {}
        fornecedor_atual = None
        dados_acumulados = []

        # 4. Organizando a bagunça (Processamento)
        for i in range(len(df_bruto)):
            linha = df_bruto.iloc[i]
            col0 = str(linha[0]).strip() if pd.notna(linha[0]) else ""

            # Se achar a palavra "Conta:", é um novo fornecedor
            if "Conta:" in col0:
                if fornecedor_atual and dados_acumulados:
                    banco_fornecedores[fornecedor_atual] = pd.DataFrame(dados_acumulados)
                
                codigo = str(linha.iloc[1]).strip() if pd.notna(linha.iloc[1]) else "000"
                nome_forn = "Sem Nome"
                for c in [5, 6, 2]:
                    if len(linha) > c and pd.notna(linha.iloc[c]):
                        nome_forn = str(linha.iloc[c]).strip()
                        break
                fornecedor_atual = f"{codigo} - {nome_forn}"
                dados_acumulados = []
                continue
            
            # Tentando ler os números de Débito e Crédito
            try:
                if len(linha) > 9:
                    def para_numero(valor):
                        if pd.isna(valor) or str(valor).strip() == "": return 0.0
                        # Limpa pontos de milhar e troca vírgula por ponto
                        v = str(valor).replace('.', '').replace(',', '.')
                        try:
                            return float(v)
                        except:
                            return 0.0

                    d = para_numero(linha[8])
                    c = para_numero(linha[9])

                    if d == 0 and c == 0: continue

                    # Pega a Data e a Nota Fiscal
                    data_txt = str(linha[0])
                    hist = str(linha[2])
                    nfe = re.findall(r'NFe\s?(\d+)', hist)
                    num_nota = nfe[0] if nfe else str(linha[1])

                    dados_acumulados.append({
                        "Data": data_txt, "NF": num_nota, "Histórico": hist, 
                        "Débito": -d, "Crédito": c
                    })
            except:
                continue

        # Guarda o último da lista
        if fornecedor_atual and dados_acumulados:
            banco_fornecedores[fornecedor_atual] = pd.DataFrame(dados_acumulados)

        # 5. Criando o desenho (Excel)
        if banco_fornecedores:
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter
