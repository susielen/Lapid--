import streamlit as st
import pandas as pd
import re
from io import BytesIO

st.set_page_config(page_title="Conciliador Mestre", layout="wide")
st.title("🤖 Robô Conciliador (Ajuste de Colunas e Filtro de Erros)")

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
                    # IGNORA ERRO DA DATA
                    try: 
                        dt = pd.to_datetime(lin[0]).strftime('%d/%m/%y')
                    except: 
                        dt = str(lin[0])[:10] if pd.notna(lin[0]) else ""
                    
                    # IGNORA ERRO DA NF (Se não achar NFe, usa o que estiver na coluna 1)
                    nf_find = re.findall(r'NFe\s?(\d+)', str(lin[2]))
                    nf_final = nf_find[0] if nf_find else str(lin[1])
                    
                    dados.append({"Data": dt, "NF": nf_final, "Hist": str(lin[2]), "Deb": -d, "Cred": c})

        if f_atual and dados: banco[f_atual] = pd.DataFrame(dados)

        if banco:
            out = BytesIO()
            with pd.ExcelWriter(out, engine='xlsxwriter') as writer:
                wb = writer.book
                # Formatos
                f_tit = wb.add_format({'bold':1,'align':'center','valign':'vcenter','bg_color':'#D3D3D3','border':1, 'font_size': 14})
                f_std = wb.add_format({'border':1})
                f_cen = wb.add_format({'border':1, 'align':'center'})
                f_cur = wb.add_format({'num_format':'R$ #,##0.00','border':1})
                f_vde = wb.add_format({'num_format':'R$ #,##0.00','font_color':'green','bold':1,'border':1})
                f_vrm = wb.add_format({'num_format':'R$ #,##0.00','font_color':'red', 'bold':1,'border':1})
                f_cab = wb.add_format({'bold':1,'bg_color':'#F2F2F2','border':1, 'align':'center'})

                for f, df in banco.items():
                    aba = re.sub(r'[\\/*?:\[\]]', '', f)[:31]
                    ws = wb.add_worksheet(aba)
                    ws.hide_gridlines(2)
                    
                    # Nome Empresa (Linhas 2 e 3)
                    ws.merge_range('B2:M3', f"EMPRESA: {nome_emp}", f_tit)
                    
                    # Nome Fornecedor na linha 8 (Linhas 5 e 6 excluídas no visual)
                    ws.merge_range('B8:F8', f, f_cab)
                    
                    # Cabeçalhos do Razão (Linha 10)
                    for ci, v in enumerate(["Data","NF","Histórico","Débito","Crédito"]):
                        ws.write(9, ci+1, v, f_cab)
                    
                    for ri, row in enumerate(df.values):
                        ws.write(10+ri, 1, row[0], f_cen)
                        ws.write(10+ri, 2, row[1], f_cen)
                        ws.write(10+ri, 3, row[2], f_std)
                        ws.write(10+ri, 4, row[3], f_cur)
                        ws.write(10+ri, 5, row[4], f_cur)
                    
                    # Totais Razão (Pula 1 linha)
                    r_fim = 11 + len(df)
                    ws.write(r_fim, 3, "TOTAIS:", f_cab)
                    ws.write(r_fim, 4, df['Deb'].sum(), f_cur)
                    ws.write(r_fim, 5, df['Cred'].sum(), f_cur)
                    
                    # Conciliação (Coluna I em diante)
                    res = df.groupby("NF").agg({"Deb":"sum","Cred":"sum"}).reset_index()
                    res["Dif"] = res["Deb"] + res["Cred"]
                    for ci, v in enumerate(["NF","Deb","Cred","Dif"]):
                        ws.write(9, ci+8, v, f_cab)
                    for ri, row in enumerate(res.values):
                        ws.write(10+ri, 8, row[0], f_cen)
                        ws.write(10+ri, 9, row[1], f_cur)
                        ws.write(10+ri, 10, row[2], f_cur)
                        ws.write(10+ri, 11, row[3], f_cur)
                    
                    # Saldo Final (Pula 1 linha)
                    rf_res = 11 + len(res)
                    s = res["Dif"].sum()
                    ws.write(rf_res, 10, "Saldo Final:", f_cab)
                    ws.write(rf_res, 11, s, f_vde if s >= 0 else f_vrm)
                    
                    # --- AJUSTE DE LARGURA DAS COLUNAS ---
                    ws.set_column('B:F', 18)  # Razão
                    ws.set_column('G:H', 2)   # COLUNAS G e H BEM MAGRINHAS
                    ws.set_column('I:L', 18)  # Conciliação

            st.success("✅ Feito! Colunas G e H diminuídas e erros de data/NF ignorados.")
            st.download_button("📥 Baixar Planilha Ajustada", out.getvalue(), "conciliacao.xlsx")
    except Exception as e:
        st.error(f"Erro: {e}")
