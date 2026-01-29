import streamlit as st
import pandas as pd
import re
from io import BytesIO

st.set_page_config(page_title="Conciliador com Bordas", layout="wide")
st.title("🤖 Robô Conciliador (Bordas no Fornecedor e Tabelas)")

def to_num(val):
    try:
        if pd.isna(val): return 0.0
        return float(str(val).replace('.', '').replace(',', '.'))
    except: return 0.0

arquivo = st.file_uploader("Suba o arquivo XLSX ou CSV", type=["xlsx", "csv"])

if arquivo:
    try:
        df_bruto = pd.read_excel(arquivo, header=None) if arquivo.name.endswith('xlsx') else pd.read_csv(arquivo, header=None)
        
        nome_emp = "EMPRESA"
        for i in range(min(15, len(df_bruto))):
            if "Empresa:" in str(df_bruto.iloc[i, 0]):
                nome_emp = str(df_bruto.iloc[i, 2])
                break

        banco = {}
        f_atual, dados = None, []

        for i in range(len(df_bruto)):
            lin = df_bruto.iloc[i]
            if "Conta:" in str(lin[0]):
                if f_atual and dados: banco[f_atual] = pd.DataFrame(dados)
                cod = str(lin[1]).strip()
                nom = str(lin[5]) if pd.notna(lin[5]) else str(lin[2])
                f_atual = f"{cod} - {nom}"
                dados = []
            elif len(lin) > 9:
                d, c = to_num(lin[8]), to_num(lin[9])
                if d != 0 or c != 0:
                    data_orig = str(lin[0])
                    try:
                        data_curta = pd.to_datetime(data_orig).strftime('%d/%m/%y')
                    except:
                        data_curta = data_orig[:8]

                    nf = re.findall(r'NFe\s?(\d+)', str(lin[2]))
                    dados.append({
                        "Data": data_curta, 
                        "NF": nf[0] if nf else str(lin[1]), 
                        "Hist": str(lin[2]), 
                        "Deb": -d, 
                        "Cred": c
                    })

        if f_atual and dados: banco[f_atual] = pd.DataFrame(dados)

        if banco:
            out = BytesIO()
            with pd.ExcelWriter(out, engine='xlsxwriter') as writer:
                wb = writer.book
                
                # --- FORMATOS COM BORDAS ---
                f_tit = wb.add_format({'bold':True, 'align':'center', 'bg_color':'#D3D3D3', 'border':1})
                f_std = wb.add_format({'align':'left', 'border':1})
                f_cur = wb.add_format({'num_format':'R$ #,##0.00', 'border':1})
                f_vde = wb.add_format({'num_format':'R$ #,##0.00', 'font_color':'green', 'bold':1, 'border':1})
                f_vrm = wb.add_format({'num_format':'R$ #,##0.00', 'font_color':'red', 'bold':1, 'border':1})
                f_neg = wb.add_format({'bold':True, 'border':1})
                f_cab = wb.add_format({'bold':True, 'bg_color':'#F2F2F2', 'align':'center', 'border':1})

                for f, df in banco.items():
                    aba = re.sub(r'[\\/*?:\[\]]', '', f)[:31]
                    ws = wb.add_worksheet(aba)
                    ws.hide_gridlines(2)
                    writer.sheets[aba] = ws
                    
                    # 1. Nome da Empresa (B2:M3)
                    ws.merge_range('B2:M3', f"EMPRESA: {nome_emp}", f_tit)
                    
                    # 2. Nome do Fornecedor (B8:F8) - Agora com borda e mesclado
                    ws.merge_range('B8:F8', f"FORNECEDOR: {f}", f_neg)
                    
                    # --- TABELA RAZÃO (Linha 10) ---
                    for col_num, value in enumerate(df.columns.values):
                        ws.write(9, col_num + 1, value, f_cab)
                    
                    # Escreve os dados com bordas (usando set_column para garantir os formatos)
                    for r_idx, row
