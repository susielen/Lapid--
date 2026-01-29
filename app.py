import streamlit as st
import pandas as pd
import re
from io import BytesIO

st.set_page_config(page_title="Conciliador Mestre", layout="wide")

st.title("🤖 Robô Conciliador (Versão Final)")

arquivo = st.file_uploader("Suba o Razão (Excel ou CSV)", type=["xlsx", "csv"])

if arquivo:
    try:
        # 1. Leitura do arquivo
        if arquivo.name.endswith('.csv'):
            df_bruto = pd.read_csv(arquivo, header=None)
        else:
            df_bruto = pd.read_excel(arquivo, engine='openpyxl', header=None)
        
        # Busca nome da empresa
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
            
            try:
                # Verifica débito e crédito (colunas 8 e 9)
                if len(linha) > 9:
                    val_deb = str(linha[8]).replace('.', '').replace(',', '.') if pd.notna(linha[8]) else "0"
                    val_cre = str(linha[9]).replace('.', '').replace(',', '.') if pd.notna(linha[9]) else "0"
                    
                    try:
                        d = float(val_deb)
                        c = float(val_cre)
                    except ValueError:
                        continue # Pula se não for número

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

        if fornecedor_atual and dados_acumulados:
            banco_fornecedores[fornecedor_atual] = pd.DataFrame(dados_acumulados)

        # 3. Geração do Excel
        if banco_fornecedores:
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                workbook = writer.book
                
                # FORMATOS
                fmt_emp = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#D3D3D3', 'border': 1})
                fmt_forn = workbook.add_format({'bold': True, 'align': 'left'})
                fmt_moeda = workbook.add_format({'num_format': '_-R$ * #,##0.00_-;-R$ * #,##0.00_-;_-R$ * "-"??_-;_-@_-', 'border': 1})
                fmt_vde = workbook.add_format({'num_format': '_-R$ * #,##0.00_-', 'font_color': 'green', 'bold': True, 'border': 1})
                fmt_vrm = workbook.add_format({'num_format': '-R$ * #,##0.00_-', 'font_color': 'red', 'bold': True, 'border': 1})
                fmt_neg = workbook.add_format({'bold': True, 'border': 1})

                for f_nome, df_f in banco_fornecedores.items():
                    aba = "".join(c for c in f_nome[:31] if c.isalnum() or c in " -")
                    ws = workbook.add_worksheet(aba)
                    ws.hide_gridlines(2)
                    writer.sheets[aba] = ws
                    
                    # Layout
                    ws.merge_range('B2:M3', f"EMPRESA: {nome_empresa}", fmt_emp)
                    ws.write('B8', f"FORNECEDOR: {f_nome}", fmt_forn)
                    
                    # Tabelas na Linha 10
                    df_f.to_excel(writer, sheet_name=aba, startrow=9, startcol=1, index=False)
                    df_res = df_f.groupby("NF").agg({"Débito":"sum", "Crédito":"sum"}).reset_index()
                    df_res["D
