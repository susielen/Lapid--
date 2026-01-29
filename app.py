import streamlit as st
import pandas as pd
import re
from io import BytesIO

st.set_page_config(page_title="Conciliador Clean", layout="wide")
st.title("🤖 Robô Conciliador (Visual 100% Clean)")

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
                    nf = re.findall(r'NFe\s?(\d+)', str(lin[2]))
                    dados.append({"Data": str(lin[0]), "NF": nf[0] if nf else str(lin[1]), "Hist": str(lin[2]), "Deb": -d, "Cred": c})

        if f_atual and dados: banco[f_atual] = pd.DataFrame(dados)

        if banco:
            out = BytesIO()
            with pd.ExcelWriter(out, engine='xlsxwriter') as writer:
                wb = writer.book
                
                # --- FORMATOS SEM BORDAS ---
                f_tit = wb.add_format({'bold':True, 'align':'center', 'bg_color':'#D3D3D3'}) # Titulo cinza sem borda
                f_std = wb.add_format({'align':'left'})
                f_cur = wb.add_format({'num_format':'R$ #,##0.00'})
                f_vde = wb.add_format({'num_format':'R$ #,##0.00', 'font_color':'green', 'bold':1})
                f_vrm = wb.add_format({'num_format':'R$ #,##0.00', 'font_color':'red', 'bold':1})
                f_neg = wb.add_format({'bold':True})

                for f, df in banco.items():
                    aba = re.sub(r'[\\/*?:\[\]]', '', f)[:31]
                    ws = wb.add_worksheet(aba)
                    ws.hide_gridlines(2) # Apaga as linhas cinzas de fundo
                    writer.sheets[aba] = ws
                    
                    # Nome da Empresa
                    ws.merge_range('B2:M3', f"EMPRESA: {nome_emp}", f_tit)
                    
                    # Nome do Fornecedor
                    ws.write('B8', f"FORNECEDOR: {f}", f_neg)
                    
                    # Tabela Razão (Linha 10)
                    df.to_excel(writer, sheet_name=aba, startrow=9, startcol=1, index=False)
                    
                    # Totais do Razão no Final
                    row_razao_fim = 10 + len(df)
                    ws.write(row_razao_fim, 3, 'TOTAIS:', f_neg)
                    ws.write(row_razao_fim, 4, df['Deb'].sum(), f_cur)
                    ws.write(row_razao_fim, 5, df['Cred'].sum(), f_cur)
                    
                    # Tabela Conciliação
                    res = df.groupby("NF").agg({"Deb":"sum", "Cred":"sum"}).reset_index()
                    res["Dif"] = res["Deb"] + res["Cred"]
                    res.to_excel(writer, sheet_name=aba, startrow=9, startcol=7, index=False)
                    
                    # Saldo Final da Conciliação
                    row_res_fim = 10 + len(res)
                    saldo = res["Dif"].sum()
                    ws.write(row_res_fim, 8, "Saldo Final:", f_neg)
                    ws.write(row_res_fim, 9, saldo, f_vde if saldo >= 0 else f_vrm)
                    
                    # Alarga as colunas mas SEM aplicar bordas
                    ws.set_column('B:Z', 18, f_cur)

            st.success("✅ Relatório limpo e pronto!")
            st.download_button("📥 Baixar Excel Sem Bordas", out.getvalue(), "conciliacao_clean.xlsx")
            
    except Exception as e:
        st.error(f"Erro ao limpar: {e}")
