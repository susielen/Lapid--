import streamlit as st
import pandas as pd
import re
from io import BytesIO

st.set_page_config(page_title="Conciliador Clean", layout="wide")
st.title("🤖 Robô Conciliador (Datas Curtas e Estilo Tabela)")

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
                    # Encurta a Data (pega apenas os primeiros 10 caracteres e tenta formatar)
                    data_orig = str(lin[0])
                    try:
                        data_curta = pd.to_datetime(data_orig).strftime('%d/%m/%y')
                    except:
                        data_curta = data_orig[:8] # Se falhar, apenas corta o texto

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
                
                # FORMATOS
                f_tit = wb.add_format({'bold':True, 'align':'center', 'bg_color':'#D3D3D3'})
                f_std = wb.add_format({'align':'left'})
                f_cur = wb.add_format({'num_format':'R$ #,##0.00'})
                f_vde = wb.add_format({'num_format':'R$ #,##0.00', 'font_color':'green', 'bold':1})
                f_vrm = wb.add_format({'num_format':'R$ #,##0.00', 'font_color':'red', 'bold':1})
                f_neg = wb.add_format({'bold':True})
                f_cab = wb.add_format({'bold':True, 'bg_color':'#F2F2F2', 'align':'center'}) # Cabeçalho da tabela

                for f, df in banco.items():
                    aba = re.sub(r'[\\/*?:\[\]]', '', f)[:31]
                    ws = wb.add_worksheet(aba)
                    ws.hide_gridlines(2)
                    writer.sheets[aba] = ws
                    
                    # Nome da Empresa
                    ws.merge_range('B2:M3', f"EMPRESA: {nome_emp}", f_tit)
                    ws.write('B8', f"FORNECEDOR: {f}", f_neg)
                    
                    # --- TABELA RAZÃO (Linha 10) ---
                    # Escrevemos os cabeçalhos manualmente para dar estilo de tabela
                    for col_num, value in enumerate(df.columns.values):
                        ws.write(9, col_num + 1, value, f_cab)
                    df.to_excel(writer, sheet_name=aba, startrow=10, startcol=1, index=False, header=False)
                    
                    # Totais do Razão no Final
                    row_razao_fim = 10 + len(df)
                    ws.write(row_razao_fim, 3, 'TOTAIS:', f_neg)
                    ws.write(row_razao_fim, 4, df['Deb'].sum(), f_cur)
                    ws.write(row_razao_fim, 5, df['Cred'].sum(), f_cur)
                    
                    # --- TABELA CONCILIAÇÃO ---
                    res = df.groupby("NF").agg({"Deb":"sum", "Cred":"sum"}).reset_index()
                    res["Dif"] = res["Deb"] + res["Cred"]
                    
                    # Cabeçalhos da Conciliação
                    for col_num, value in enumerate(res.columns.values):
                        ws.write(9, col_num + 8, value, f_cab)
                    res.to_excel(writer, sheet_name=aba, startrow=10, startcol=8, index=False, header=False)
                    
                    # Saldo Final
                    row_res_fim = 10 + len(res)
                    saldo = res["Dif"].sum()
                    ws.write(row_res_fim, 9, "Saldo Final:", f_neg)
                    ws.write(row_res_fim, 10, saldo, f_vde if saldo >= 0 else f_vrm)
                    
                    # Ajuste das colunas
                    ws.set_column('B:C', 10, f_std) # Data e NF curtas
                    ws.set_column('D:D', 30, f_std) # Histórico largo
                    ws.set_column('E:F', 15, f_cur) # Valores
                    ws.set_column('I:K', 15, f_cur) # Conciliação

            st.success("✅ Tabelas organizadas com datas curtas!")
            st.download_button("📥 Baixar Excel Estilo Tabela", out.getvalue(), "conciliacao_tabelada.xlsx")
            
    except Exception as e:
        st.error(f"Erro ao formatar tabelas: {e}")
