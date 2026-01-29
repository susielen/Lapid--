import streamlit as st
import pandas as pd
import re
from io import BytesIO

st.set_page_config(page_title="Conciliador Final", layout="wide")
st.title("🤖 Robô Conciliador (Versão Corrigida)")

def to_num(val):
    try:
        if pd.isna(val): return 0.0
        return float(str(val).replace('.', '').replace(',', '.'))
    except: return 0.0

arquivo = st.file_uploader("Suba o arquivo XLSX ou CSV", type=["xlsx", "csv"])

if arquivo:
    try:
        df_bruto = pd.read_excel(arquivo, header=None) if arquivo.name.endswith('xlsx') else pd.read_csv(arquivo, header=None)
        
        # Busca nome da empresa
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
                    try: dt = pd.to_datetime(lin[0]).strftime('%d/%m/%y')
                    except: dt = str(lin[0])[:8]
                    nf = re.findall(r'NFe\s?(\d+)', str(lin[2]))
                    dados.append({"Data": dt, "NF": nf[0] if nf else str(lin[1]), "Hist": str(lin[2]), "Deb": -d, "Cred": c})

        if f_atual and dados: banco[f_atual] = pd.DataFrame(dados)

        if banco:
            out = BytesIO()
            with pd.ExcelWriter(out, engine='xlsxwriter') as writer:
                wb = writer.book
                # Estilos com Bordas
                f_tit = wb.add_format({'bold':1,'align':'center','bg_color':'#D3D3D3','border':1})
                f_std = wb.add_format({'border':1})
                f_cur = wb.add_format({'num_format':'R$ #,##0.00','border':1})
                f_vde = wb.add_format({'num_format':'R$ #,##0.00','font_color':'green','bold':1,'border':1})
                f_vrm = wb.add_format({'num_format':'R$ #,##0.00','font_color':'red', 'bold':1,'border':1})
                f_cab = wb.add_format({'bold':1,'bg_color':'#F2F2F2','border':1})

                for f, df in banco.items():
                    aba = re.sub(r'[\\/*?:\[\]]', '', f)[:31]
                    ws = wb.add_worksheet(aba)
                    ws.hide_gridlines(2)
                    
                    # Título e Fornecedor
                    ws.merge_range('B2:M3', f"EMPRESA: {nome_emp}", f_tit)
                    ws.merge_range('B8:F8', f"FORNECEDOR: {f}", f_cab)
                    
                    # Razão
                    cols = ["Data","NF","Histórico","Débito","Crédito"]
                    for c_idx, val in enumerate(cols): ws.write(9, c_idx+1, val, f_cab)
                    for r_idx, row in enumerate(df.values):
                        for c_idx, val in enumerate(row):
                            fmt = f_cur if c_idx >= 3 else f_std
                            ws.write(10+r_idx, c_idx+1, val, fmt)
                    
                    # Totais Razão
                    r_fim = 10 + len(df)
                    ws.write(r_fim, 3, "TOTAIS:", f_cab)
                    ws.write(r_fim, 4, df['Deb'].sum(), f_cur)
                    ws.write(r_fim, 5, df['Cred'].sum(), f_cur)
                    
                    # Conciliação
                    res = df.groupby("NF").agg({"Deb":"sum","Cred":"sum"}).reset_index()
                    res["Dif"] = res["Deb"] + res["Cred"]
                    c_off = 8
                    cols_res = ["NF","Deb","Cred","Dif"]
                    for c_idx, val in enumerate(cols_res): ws.write(9, c_idx+c_off, val, f_cab)
                    for r_idx, row in enumerate(res.values):
                        for c_idx, val in enumerate(row):
                            fmt = f_cur if c_idx >= 1 else f_std
                            ws.write(10+r_idx, c_idx+c_off, val, fmt)
                    
                    # Saldo Final
                    r_res_fim = 10 + len(res)
                    s = res["Dif"].sum()
                    ws.write(r_res_fim, c_off+2, "Saldo Final:", f_cab)
                    ws.write(r_res_fim, c_off+3, s, f_vde if s >= 0 else f_vrm)
                    ws.set_column('B:M', 15)

            st.success("✅ Finalmente pronto e sem erros!")
            st.download_button("📥 Baixar Planilha", out.getvalue(), "conciliacao.xlsx")
    except Exception as e:
        st.error(f"Erro: {e}")
