import streamlit as st
import pandas as pd
import io
import re
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Conciliador Pro", layout="wide")
st.title("🤖 Conciliador: Versão Blindada")

arquivo = st.file_uploader("Suba o Razão aqui", type=["csv", "xlsx"])

def extrair_nfe(texto):
    match = re.search(r'(?:NFE|NF|NOTA|Nº)\s*(\d+)', str(texto).upper())
    if match:
        try: return int(match.group(1))
        except: return match.group(1)
    return ""

def limpar_nome(linha_txt):
    linha_txt = str(linha_txt).replace('nan', '').upper()
    match_cod = re.search(r'CONTA:\s*(\d+)', linha_txt)
    codigo = match_cod.group(1) if match_cod else ""
    nome = linha_txt.split("CONTA:")[-1].replace('NOME:', '').strip()
    nome = re.sub(r'(\d+\.)+\d+', '', nome).strip()
    return f"{codigo} - {nome}" if codigo else nome

if arquivo is not None:
    try:
        df_raw = pd.read_excel(arquivo) if arquivo.name.endswith('.xlsx') else pd.read_csv(arquivo, encoding='latin-1', sep=None, engine='python')
        
        dict_forn = {}
        atual = None

        for i, linha in df_raw.iterrows():
            txt = " ".join([str(v) for v in linha.values]).upper()
            if "CONTA:" in txt:
                atual = limpar_nome(txt)
                dict_forn[atual] = []
                continue
            
            data_orig = str(linha.iloc[0])
            if "/" in data_orig or "-" in data_orig:
                try:
                    def num(v):
                        v = str(v).replace('.', '').replace(',', '.')
                        try: return float(v)
                        except: return 0.0
                    
                    d, c = num(linha.iloc[8]), num(linha.iloc[9])
                    if (d > 0 or c > 0) and atual:
                        h = str(linha.iloc[2]).replace('nan', '')
                        dict_forn[atual].append({'Data': data_orig, 'Nº NF': extrair_nfe(h), 'Histórico': h, 'Débito': d, 'Crédito': c})
                except: continue

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            for nome, dados in dict_forn.items():
                if not dados: continue
                df_f = pd.DataFrame(dados)
                df_c = df_f.groupby('Nº NF').agg({'Débito': 'sum', 'Crédito': 'sum'}).reset_index()
                df_c['DIFERENÇA'] = df_c['Crédito'] - df_c['Débito']
                df_c['STATUS'] = df_c['DIFERENÇA'].apply(lambda x: "OK" if abs(x) < 0.05 else "ERRO")
                
                aba = re.sub(r'[\\/*?:\[\]]', '', nome)[:31]
                df_f.to_excel(writer, sheet_name=aba, index=False, startrow=6)
                df_c.to_excel(writer, sheet_name=aba, index=False, startrow=6, startcol=8)
                ws = writer.sheets[aba]
                
                # Estilos
                b = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
                cnt, dir = Alignment(horizontal='center'), Alignment(horizontal='right')
                fmt = '#,##0.00'

                # Título
                ws.merge_cells('A1:M1')
                ws['A1'] = nome
                ws['A1'].font, ws['A1'].alignment = Font(bold=True, size=12), cnt

                # Linha 3: SALDO
                ws.cell(row=3, column=4, value="SALDO").font = Font(bold=True)
                ws.cell(row=3, column=4).alignment = dir
                val_s = df_f['Crédito'].sum() - df_f['Débito'].sum()
                cel_s = ws.cell(row=3, column=5, value=val_s)
                cel_s.font, cel_s.border, cel_s.number_format = Font(bold=True), b, fmt

                # Linha 5: TOTAIS e CONCILIAÇÃO
                ws.cell(row=5, column=3, value="TOTAIS").alignment = dir
                ws.cell(row=4, column=4, value="DÉBITO").alignment, ws.cell(row=4, column=5, value="CRÉDITO").alignment = cnt, cnt
                
                ws.cell(row=5, column=4, value=df_f['Débito'].sum()).border = b
                ws.cell(row=5, column=5, value=df_f['Crédito'].sum()).border
