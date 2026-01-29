import streamlit as st
import pandas as pd
import io
import re
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Conciliador Contábil Pro", layout="wide")
st.title("🤖 Conciliador: Versão Estável e Final")

arquivo = st.file_uploader("Suba o Razão do Domínio aqui", type=["csv", "xlsx"])

def extrair_nfe(texto):
    match = re.search(r'(?:NFE|NF|NOTA|Nº)\s*(\d+)', str(texto).upper())
    if match:
        try: return int(match.group(1))
        except: return match.group(1)
    return ""

def limpar_nome_simples(linha_txt):
    linha_txt = str(linha_txt).replace('nan', '').replace('NAN', '').replace('NaN', '')
    match_cod = re.search(r'CONTA:\s*(\d+)', linha_txt)
    codigo = match_cod.group(1) if match_cod else ""
    nome = linha_txt.split("CONTA:")[-1]
    nome = re.sub(r'(\d+\.)+\d+', '', nome) 
    nome = nome.replace(codigo, '').replace('NOME:', '').strip()
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
            valores_limpos = [str(v).replace('nan', '').strip() for v in linha.values]
            linha_txt = " ".join(valores_limpos).upper()
            
            if "CONTA:" in linha_txt:
                fornecedor_atual = limpar_nome_simples(linha_txt)
                dict_fornecedores[fornecedor_atual] = []
                continue
            
            data_orig = str(linha.iloc[0])
            if "/" in data_orig or (len(data_orig) >= 8 and "-" in data_orig):
                try: data_dt = pd.to_datetime(data_orig)
                except: data_dt = data_orig

                def limpar_num(v):
                    if pd.isna(v) or str(v).lower() == 'nan' or str(v).strip() == '': return 0.0
                    v = str(v).replace('.', '').replace(',', '.')
                    try: return float(v)
                    except: return 0.0
                
                deb = limpar_num(linha.iloc[8])
                cre = limpar_num(linha.iloc[9])

                if (deb > 0 or cre > 0) and fornecedor_atual:
                    hist = str(linha.iloc[2]).replace('nan', '')
                    dict_fornecedores[fornecedor_atual].append({
                        'Data': data_dt, 'Nº NF': extrair_nfe(hist),
                        'Histórico': hist, 'Débito': deb, 'Crédito': cre
                    })

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            for forn, lancamentos in dict_fornecedores.items():
                if not lancamentos: continue
                
                df_f = pd.DataFrame(lancamentos)
                df_c = df_f.groupby('Nº NF').agg({'Débito': 'sum', 'Crédito': 'sum'}).reset_index()
                df_c['DIFERENÇA'] = df_c['Crédito'] - df_c['Débito']
                df_c['STATUS'] = df_c['DIFERENÇA'].apply(lambda x: "OK" if abs(x) < 0.05 else "DIVERGENTE")
                
                nome_aba = re.sub(r'[\\/*?:\[\]]', '', forn)[:31]
                df_f.to_excel(writer, sheet_name=nome_aba, index=False, startrow=6)
                df_c.to_excel(writer, sheet_name=nome_aba, index=False, startrow=6, startcol=8)
                
                sheet = writer.sheets[nome_aba]
                
                fmt_contabil = '_-R$ * #,##0.00_-;-R$ * #,##0.00_-;_-R$ * "-"??_-;_-@_-'
                borda_fina = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
                alinhar_centro = Alignment(horizontal='center')
                alinhar_direita = Alignment(horizontal='right')

                # 1. TÍTULO
                sheet.merge_cells('A1:M1')
                sheet['A1'] = forn
                sheet['A1'].font = Font(bold=True, size=14)
                sheet['A1'].alignment = alinhar_centro

                # 2. SALDO (LINHA 3)
                sheet.cell(row=3, column=4, value="SALDO").font = Font(bold=True)
                sheet.cell(row=3, column=4).alignment = alinhar_direita
                saldo_val = df_f['Crédito'].sum() - df_f['Débito'].sum()
                c_saldo = sheet.cell(row=3, column=5, value=saldo_val)
                
