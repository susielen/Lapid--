import streamlit as st
import pandas as pd
import re
from io import BytesIO

st.set_page_config(page_title="Conciliador Mestre", layout="wide")
st.title("🤖 Robô Conciliador (Versão Restaurada)")

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
                f_atual = f"{str(lin[1])} - {str(lin[5]) if pd.notna(lin[5]) else str(lin[2])}"
                dados = []
            elif len(lin) > 9:
                d, c = to_num(lin[8]), to_num(lin[9])
                if d != 0 or c != 0:
                    try: dt = pd.to_datetime(lin[0]).strftime('%d/%m/%Y')
                    except: dt = str(lin[0])[:10] if pd.notna(lin[0]) else ""
                    
                    try:
                        nf_find = re.findall(r'NFe\s?(\d+)', str(lin[2]))
                        nf_final = nf_find[0] if nf_find else str(lin[1])
                    except:
                        nf_final = str(lin[1])
                        
                    dados.append({"Data": dt, "NF": nf_final, "Hist": str(lin[2]), "Deb": -d, "Cred": c})

        if f_atual and dados: banco[f_atual] = pd.DataFrame(dados)

        if banco:
            out = BytesIO()
            with pd.ExcelWriter(out, engine='xlsxwriter') as writer:
                wb = writer.book
                f_tit = wb.add_format({'bold':1,'align':'center','valign':'vcenter','bg_color':'#D3D3D3','border':1, 'font_size': 14})
                f_std = wb.add_format({'border':1})
                f_cen = wb.add_format({'border':1, 'align':'center'})
                f_contabil = wb.add_format({'num_format': '_-R$ * #,##0.00_-;-R$ * #,##0.00_-;_-R$ * "-"??_-;_-@_-','border': 1})
                f_vde = wb.add_format({'num_format': '_-R$ * #,##0.00_-', 'font_color':'green','bold':1,'border':1})
                f_vrm = wb.add_format({'num_format': '_-R$ * #,##0.00_-', 'font_color':'red', 'bold':1,'border':1})
                f_cab = wb.add_format({'bold':1,'bg_color':'#F2F2F2','border':1, 'align':'center'})

                for f, df in banco.items():
                    aba = re.sub(r'[\\/*?:\[\]]', '', f)[:31]
                    ws = wb.add_worksheet(aba)
                    ws.hide_gridlines(2)
                    
                    # Ajustes de bordas e erros visuais
                    ws.set_column('A:A', 2)
                    ws.set_row(0, 5)
                    ws.ignore_errors({'number_stored_as_text': 'B1:C5000 I1:I5000'})
                    
                    ws.merge_range('B2:M3', f"EMPRESA: {nome_emp}", f_tit)
                    ws.merge_range('B5:F5', f, f_cab)
                    
                    # Tabelas na Linha 7
                    for ci, v in enumerate(["Data","NF","Histórico","Débito","Crédito"]):
                        ws.write(6, ci+1, v, f_cab)
                    
                    for ri, row in enumerate(df.values):
                        ws.write(7+ri, 1, row[0], f_cen)
                        ws.write(7+ri, 2, row[1], f
