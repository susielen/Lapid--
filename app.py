import streamlit as st
import pandas as pd
import re
from io import BytesIO
import time

# 1. Configuração
st.set_page_config(page_title="LAPIDÔ", layout="wide")

# 2. Título (Ajuste para #1E90FF se for Fornecedores)
st.markdown("""
    <style>
    .titulo {color: #FF8C00; font-size: 40px; font-weight: bold; text-align: center; padding: 10px;}
    </style>
    <p class="titulo">💎 LAPIDÔ: O Mestre das Contas</p>
    """, unsafe_allow_html=True)

def to_num(val):
    try:
        if pd.isna(val) or str(val).strip() == '': return 0.0
        return float(str(val).replace('.', '').replace(',', '.'))
    except: return 0.0

with st.sidebar:
    st.header("⚙️ Painel")
    arquivo = st.file_uploader("Suba o arquivo aqui", type=["xlsx", "csv"])

if arquivo:
    with st.spinner('💎 Colocando molduras em tudo...'):
        try:
            time.sleep(1)
            df_bruto = pd.read_excel(arquivo, header=None) if arquivo.name.endswith('xlsx') else pd.read_csv(arquivo, header=None)
            
            nome_emp = "EMPRESA"
            for i in range(min(15, len(df_bruto))):
                if "Empresa:" in str(df_bruto.iloc[i, 0]):
                    nome_emp = str(df_bruto.iloc[i, 2])
                    break

            banco, f_info = {}, {}
            f_cod, dados = None, []

            for i in range(len(df_bruto)):
                lin = df_bruto.iloc[i]
                if "Conta:" in str(lin[0]):
                    if f_cod and dados: banco[f_cod] = pd.DataFrame(dados)
                    f_cod = str(lin[1]).strip()
                    f_info[f_cod] = f"{f_cod} - {str(lin[5]) if pd.notna(lin[5]) else str(lin[2])}"
                    dados = []
                elif len(lin) > 9:
                    deb, cre = to_num(lin[8]), to_num(lin[9])
                    hist = str(lin[2]).strip()
                    if (deb != 0 or cre != 0) and pd.notna(lin[0]):
                        if 'TOTAL' in hist.upper(): continue
                        try: dt = pd.to_datetime(lin[0]).strftime('%d/%m/%Y')
                        except: dt = str(lin[0])
                        nf_f = re.findall(r'NFe\s?(\d+)', hist)
                        nf = nf_f[0] if nf_f else (str(lin[1]).strip() if pd.notna(lin[1]) else "S/N")
                        dados.append({"Data": dt, "NF": nf, "Hist": hist, "Deb": deb, "Cred": -cre})

            if f_cod and dados: banco[f_cod] = pd.DataFrame(dados)

            if banco:
                out = BytesIO()
                with pd.ExcelWriter(out, engine='xlsxwriter') as writer:
                    wb = writer.book
                    
                    # FORMATOS COM BORDAS
                    f_cent = wb.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1})
                    f_moeda = wb.add_format({'num_format': 'R$ #,##0.00', 'border': 1})
                    f_std = wb.add_format({'border': 1})
                    f_cab = wb.add_format({'bold': 1, 'bg_color': '#F2F2F2', 'align': 'center', 'valign': 'vcenter', 'border': 1})
                    f_empresa = wb.add_format({'bold': 1, 'font_size': 14, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True, 'bg_color': '#D3D3D3', 'border': 1})
                    f_vde = wb.add_format({'num_format': 'R$ #,##0.00', 'font_color': 'green', 'bold': 1, 'border': 1, 'align': 'center'})
                    f_vrm = wb.add_format({'num_format': 'R$ #,##0.00', 'font_color': 'red', 'bold': 1, 'border': 1, 'align': 'center'})
                    
                    # Formato para o Nome do Fornecedor com Borda
                    f_info_borda = wb.add_format({'bold': 1, 'font_size': 11, 'border': 1, 'valign': 'vcenter'})

                    for cod, df in banco.items():
                        ws = wb.add_worksheet(str(cod)[:31])
                        ws.hide_gridlines(2)
                        ws.set_column('A:A', 1)
                        ws.set_column('B:C', 15, f_cent)
                        ws.set_column('D:D', 45)
                        ws.set_column('E:F', 18)
                        ws.set_column('G:H', 1)
                        ws.set_column('I:I', 15, f_cent)
                        ws.set_column('J:L', 18)
                        
                        ws.set_row(1, 25); ws.set_row(2, 25)
                        ws.merge_range('B2:L3', f"EMPRESA: {nome_emp}", f_empresa)
                        
                        # BORDAS NAS INFORMAÇÕES DA LINHA 5
                        ws.write('B5', "FORNECEDOR/CLIENTE:", f_cab)
                        ws.merge_range('C5:F5', f_info[cod], f_info_borda) # Mesclado e com borda
                        ws.merge_range('I5:L5', "CONCILIAÇÃO POR NOTA", f_cab)
                        
                        for ci, v in enumerate(["Data","NF","Histórico","Débito","Crédito"]):
                            ws.write(6, ci+1, v, f_cab)
                        
                        for ri, row in enumerate(df.values):
                            ws.write(7+ri, 1, row[0], f_cent); ws.write(7+ri, 2, row[1], f_cent)
                            ws.write(7+ri, 3, row[2], f_std); ws.write(7+ri, 4, row[3], f_moeda); ws.write(7+ri, 5, row[4], f_moeda)
                        
                        lt = 7 + len(df) + 1 
                        ws.write(lt, 3, "TOTALIZADOR:", f_cab)
                        ws.write(lt, 4, df['Deb'].sum(), f_moeda)
                        ws.write(lt, 5, df['Cred'].sum(), f_moeda)

                        res = df.groupby("NF").agg({"Deb":"sum","Cred":"sum"}).reset_index()
                        res["Dif"] = res["Deb"] + res["Cred"]
                        for ci, v in enumerate(["NF","Deb","Cred","Dif"]): ws.write(6, ci+8, v, f_cab)
                        for ri, row in enumerate(res.values):
                            ws.write(7+ri, 8, str(row[0]), f_cent)
                            ws.write(7+ri, 9, row[1], f_moeda); ws.write(7+ri, 10, row[2], f_moeda); ws.write(7+ri, 11, row[3], f_moeda)
                        
                        s = res["Dif"].sum(); rf = 8 + len(res)
                        ws.write(rf, 10, "Saldo Final:", f_cab)
                        ws.write(rf, 11, s, f_vde if s >= 0 else f_vrm)
                
                st.success("✅ Tudo com bordas e molduras! O relatório ficou lindo.")
                st.download_button("📥 Baixar Planilha", out.getvalue(), "relatorio_final_luxo.xlsx")
        except Exception as e:
            st.error(f"Erro: {e}")
