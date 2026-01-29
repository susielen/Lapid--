import streamlit as st
import pandas as pd
import re
from io import BytesIO

# Configuração da página
st.set_page_config(page_title="Conciliador Mestre", layout="wide")

st.title("🤖 Robô Conciliador (Versão Blindada)")

arquivo = st.file_uploader("Suba o Razão (Excel ou CSV)", type=["xlsx", "csv"])

if arquivo:
    try:
        # 1. Leitura do arquivo
        if arquivo.name.endswith('.csv'):
            df_bruto = pd.read_csv(arquivo, header=None)
        else:
            df_bruto = pd.read_excel(arquivo, engine='openpyxl', header=None)
        
        # Busca nome da empresa nas primeiras linhas
        nome_empresa = "EMPRESA NÃO IDENTIFICADA"
        for i in range(min(15, len(df_bruto))):
            txt = str(df_bruto.iloc[i, 0])
            if "Empresa:" in txt or "EMPRESA:" in txt.upper():
                nome_empresa = str(df_bruto.iloc[i, 2]) if pd.notna(df_bruto.iloc[i, 2]) else nome_empresa
                break

        banco_fornecedores = {}
        fornecedor_atual = None
        dados_acumulados = []

        # 2. Processamento dos Dados
        for i in range(len(df_bruto)):
            linha = df_bruto.iloc[i]
            col0 = str(linha[0]).strip() if pd.notna(linha[0]) else ""

            # Identifica novo fornecedor
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
            
            # Tratamento de valores (Blindagem contra ValueError)
            try:
                if len(linha) > 9:
                    def limpar_numero(valor):
                        if pd.isna(valor) or str(valor).strip() == "": return 0.0
                        v = str(valor).replace('.', '').replace(',', '.')
                        try: return float(v)
                        except: return 0.0

                    d = limpar_numero(linha[8])
                    c = limpar_numero(linha[9])

                    if d == 0 and c == 0: continue

                    hist = str(linha[2])
                    nfe = re.findall(r'NFe\s?(\d+)', hist)
                    num_nota = nfe[0] if nfe else str(linha[1])

                    dados_acumulados.append({
                        "Data": str(linha[0]), "NF": num_nota, "Histórico": hist, 
                        "Débito": -d, "Crédito": c
                    })
            except:
                continue

        # Salva o último fornecedor da lista
        if fornecedor_atual and dados_acumulados:
            banco_fornecedores[fornecedor_atual] = pd.DataFrame(dados_acumulados)

        # 3. Geração do Excel Formatado
        if banco_fornecedores:
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                workbook = writer.book
                
                # Formatos (Bordas e Moeda)
                fmt_emp = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#D3D3D3', 'border': 1})
