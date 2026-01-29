import streamlit as st
import pandas as pd
import io
import re

st.set_page_config(page_title="Conciliador Contábil", layout="wide")
st.title("🤖 Conciliador: Padrão Contábil Final")

arquivo = st.file_uploader("Suba o Razão do Domínio aqui", type=["csv", "xlsx"])

def extrair_nfe(texto):
    match = re.search(r'(?:NFE|NF|NOTA|Nº)\s*(\d+)', str(texto).upper())
    return match.group(1) if match else "(vazio)"

def limpar_nome_simples(linha_txt):
    match_cod = re.search(r'CONTA:\s*(\d+)', linha_txt)
    codigo = match_cod.group(1) if match_cod else ""
    nome = linha_txt.split("CONTA:")[-1]
    nome = re.sub(r'(\d+\.)+\d+', '', nome) 
    nome = nome.replace(codigo, '').replace('NOME:', '').replace('CONTA:', '').strip()
    if nome.startswith('-'): nome = nome[1:].strip()
    return f"{codigo} - {nome}" if codigo else nome

if arquivo is not None:
    try:
        if arquivo.name.endswith('.xlsx'):
            df_raw = pd.read_excel(arquivo)
        else:
            df_raw = pd.read_csv(arquivo, encoding='latin-1', sep=None, engine='python')

        dict_fornecedores = {}
        fornecedor_atual = None

        for i, linha in df_raw.iterrows():
            linha_txt = " ".join([str(v) for v in linha.values]).upper()
            if "CONTA:" in linha_txt:
                fornecedor_atual = limpar_nome_simples(linha_txt)
                if fornecedor_atual not in dict_fornecedores:
                    dict_fornecedores[fornecedor_atual] = []
                continue
            
            data_orig = str(linha.iloc[0])
            if "/" in data_orig or (len(data_orig) >= 8 and "-" in data_orig):
                def limpar_num(v):
                    if pd.isna(v): return 0
                    v = str(v).replace('.', '').replace(',', '.')
                    return pd.to_numeric(v, errors='coerce') or 0

                deb = limpar_num(linha.iloc[8]) if len(linha) > 8 else 0
                cre = limpar_num(linha.iloc[9]) if len(linha) > 9 else 0

                if (deb > 0 or cre > 0) and fornecedor_atual:
                    hist = str(linha.iloc[2])
                    dict_fornecedores[fornecedor_atual].append({
                        'Data': data_orig,
                        'Nº NF': extrair_nfe(hist),
                        'Histórico': hist,
                        'Débito': deb,
                        'Crédito': cre
                    })

        if dict_fornecedores:
            lista_nomes = sorted([n for n in dict_fornecedores.keys() if dict_fornecedores[n]])
            st.success(f"Encontrados {len(lista_nomes)} fornecedores!")

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                for forn in lista_nomes:
                    df_f = pd.DataFrame(dict_fornecedores[forn])
                    df_c = df_f.groupby('Nº NF').agg({'Débito': 'sum', 'Crédito': 'sum'}).reset_index()
                    df_c['DIFERENÇA'] = df_c['Crédito'] - df_c['Débito']
                    df_c['STATUS'] = df_c['DIFERENÇA'].apply(lambda x: "OK" if abs(x) < 0.05 else "DIVERGENTE")
                    
                    nome_aba = re.sub(r'[\\/*?:\[\]]', '', forn)[:31]
                    df_f.to_excel(writer, sheet_name=nome_aba, index=False, startrow=2)
                    
                    sheet = writer.sheets[nome_aba]
                    fmt_contabil = '_-R$ * #,##0.00_-;-R$ * #,##0.00_-;_-R$ * "-"??_-;_-@_-'
                    
                    # Nome do Fornecedor no Topo (Célula A1)
                    sheet['A1'] = f"FORNECEDOR: {forn}"
                    sheet['A1'].font = pd.io.formats.excel.ExcelFormatter.styler_to_excel_style({'font': {'bold': True}})['font']

                    # Formatação das colunas de valor do Razão
                    for row in range(4, len(df_f) + 4):
                        sheet.cell(row=row, column=5).number_format = fmt_contabil
                        sheet.cell(row=row, column=6).number_format = fmt_contabil
                    
                    # Totais do Razão (Pula uma linha)
                    row_tot = len(df_f) + 5
                    sheet.cell(row=row_tot, column=4, value="TOTAIS:").font = pd.io.formats.excel.ExcelFormatter.styler_to_excel_style({'font': {'bold': True}})['font']
                    
                    c_deb = sheet.cell(row=row_tot, column=5, value=df_f['Débito'].sum())
                    c_cre = sheet.cell(row=row_tot, column=6, value=df_f['Crédito'].sum())
                    c_deb.number_format = fmt_contabil
                    c_cre.number_format = fmt_contabil
                    c_deb.font = c_cre.font = pd.io.formats.excel.ExcelFormatter.styler_to_excel_style({'font': {'bold': True}})['font']
                    
                    sheet.cell(row=row_tot + 1, column=4, value="SALDO FINAL:").font = pd.io.formats.excel.ExcelFormatter.styler_to_excel_style({'font': {'bold': True}})['font']
                    c_saldo = sheet.cell(row=row_tot + 1, column=6, value=df_f['Crédito'].sum() - df_f['Débito'].sum())
                    c_saldo.number_format = fmt_contabil
                    c_saldo.font = pd.io.formats.excel.ExcelFormatter.styler_to_excel_style({'font': {'bold': True}})['font']

                    # Conciliação com Totalizador na Linha 1
                    sheet.cell(row=1, column=11, value="TOTAL ABERTO:").font = pd.io.formats.excel.ExcelFormatter.styler_to_excel_style({'font': {'bold': True}})['font']
                    c_total_geral = sheet.cell(row=1, column=12, value=df_f['Crédito'].sum() - df_f['Débito'].sum())
                    c_total_geral.number_format = fmt_contabil
                    c_total_geral.font = pd.io.formats.excel.ExcelFormatter.styler_to_excel_style({'font': {'bold': True}})['font']
                    
                    df_c.to_excel(writer, sheet_name=nome_aba, index=False, startrow=2, startcol=8)
                    for row in range(4, len(df_c) + 4):
                        sheet.cell(row=row, column=10).number_format = fmt_contabil
                        sheet.cell(row=row, column=11).number_format = fmt_contabil
                        sheet.cell(row=row, column=12).number_format = fmt_contabil

            st.download_button("📥 Baixar Relatório Contábil Final", data=output.getvalue(), file_name="conciliacao_completa.xlsx")
            
    except Exception as e:
        st.error(f"Erro: {e}")
