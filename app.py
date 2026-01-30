import streamlit as st
import pandas as pd
import re
from io import BytesIO
import time

st.set_page_config(page_title="LAPIDÔ", page_icon="💎", layout="wide")

st.markdown("""
    <style>.titulo {color: #1E90FF; font-size: 48px; font-weight: bold; text-align: center;}</style>
    <p class="titulo">💎 LAPIDÔ: Fornecedores</p>
    """, unsafe_allow_html=True)

def to_num(val):
    try:
        if pd.isna(val) or str(val).strip().lower() in ['nan', '']: return 0.0
        return float(str(val).replace('.', '').replace(',', '.'))
    except: return 0.0

with st.sidebar:
    st.header("⚙️ Painel")
    arquivo = st.file_uploader("Arquivo", type=["xlsx", "csv"])

if not arquivo:
    st.warning("👈 O LAPIDÔ aguarda o arquivo na barra lateral.")
else:
    with st.spinner('💎 Polindo o diamante...'):
        try:
            time.sleep(1)
            df_bruto = pd.read_excel(arquivo, header=None) if arquivo.name.endswith('xlsx') else pd.read_csv(arquivo, header=None)
            nome_emp = "EMPRESA"
            for i in range(min(15, len(df_bruto))):
                if "Empresa:" in str(df_bruto.iloc[i, 0]):
                    nome_emp = str(df_bruto.iloc[i, 2])
                    break
            banco, f_info = {}, {}
            f_cod_atual, dados = None, []
            for i in range(len(df_bruto)):
                lin = df_bruto.iloc[i]
                if "Conta:" in str(lin[0]):
                    if f_cod_atual and dados: banco[f_cod_atual] = pd.DataFrame(dados)
                    f_cod_atual = str(lin[1]).strip()
                    f_info[f_cod_atual] = f"{f_cod_atual} - {str(lin[5]) if pd.notna(lin[5]) else str(lin[2])}"
                    dados = []
                elif len(lin) > 9:
                    deb, cre = to_num(lin[8]), to_num(lin[9])
                    hist = str(lin[2]).strip()
                    if (deb != 0 or cre != 0) and pd.notna(lin[0]):
                        if hist.upper() in ['TOTAL', 'SUBTOTAL']: continue
                        dt = pd.to_datetime(lin[0]).strftime('%d/%m/%Y')
                        nf_find = re.findall(r'NFe\s?(\d+)', hist)
                        nf = nf_find[0] if nf_find else str(lin[1])
                        dados.append({"Data": dt, "NF": nf, "Hist": hist, "Deb": -deb, "Cred": cre})
            if f_cod_atual and dados: banco[f_cod_atual] = pd.DataFrame(dados)
            if banco:
                out = BytesIO()
                with pd.ExcelWriter(out, engine='xlsxwriter') as writer:
                    wb = writer.book
                    f_tit = wb.add_format({'bold':1,'align':'center','bg_color':'#D3D3D3','border':1})
                    f_con = wb.add_format({'num_format': 'R$ #,##0.00', 'border': 1})
                    f_vde = wb.add_format({'num_format': 'R$ #,##0.00', 'font_color': 'green', 'bold': 1, 'border': 1})
                    f_vrm = wb.add_format({'num_format': 'R$ #,##0.00', 'font_color': 'red', 'bold': 1, 'border': 1})
                    f_cab = wb.add_format({'bold':1,'bg_color':'#F2F2F2','border':1, 'align':'center'})
                    for cod, df in banco.items():
                        aba = re.sub(r'[\\/*?:\[\]]', '', cod)[:31]
                        ws = wb.add_worksheet(aba)
                        ws.merge_range('B2:F3', f"EMPRESA: {nome_emp}", f_tit)
                        ws.set_column('G:H', 1) # COLUNAS FININHAS
                        ws.set_column('D:D', 35)
                        for ci, v in enumerate(["Data","NF","Histórico","Débito","Crédito"]):
                            ws.write(6, ci+1, v, f_cab)
                        for ri, row in enumerate(df.values):
                            ws.write(7+ri, 1, row[0]); ws.write(7+ri, 2, row[1])
                            ws.write(7+ri, 3, row[2]); ws.write(7+ri, 4, row[3], f_con); ws.write(7+ri, 5, row[4], f_con)
                        res = df.groupby("NF").agg({"Deb":"sum","Cred":"sum"}).reset_index()
                        res["Dif"] = res["Deb"] + res["Cred"]
                        for ci, v in enumerate(["NF","Deb","Cred","Dif"]): ws.write(6, ci+8, v, f_cab)
                        for ri, row in enumerate(res.values):
                            ws.write(7+ri, 8, row[0]); ws.write(7+ri, 9, row[1], f_con)
                            ws.write(7+ri, 10, row[2], f_con); ws.write(7+ri, 11, row[3], f_con)
                        s = res["Dif"].sum(); rf = 8 + len(res)
                        ws.write(rf, 10, "Saldo Final:", f_cab)
                        ws.write(rf, 11, s, f_vde if s >= 0 else f_vrm)
                st.success("✅ Diamante lapidado!")
                st.download_button("📥 Baixar Planilha", out.getvalue(), "fornecedores.xlsx")
        except Exception as e: st.error(f"Erro: {e}")
