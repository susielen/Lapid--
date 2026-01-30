import streamlit as st
import pandas as pd
import re
from io import BytesIO
import time
import random

# 1. Configuração
st.set_page_config(page_title="LAPIDÔ", layout="wide")

# 2. Título do Site
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
    tipo_robo = st.radio("Este robô é de:", ["Cliente", "Fornecedor"])
    arquivo = st.file_uploader("Suba o arquivo aqui", type=["xlsx", "csv"])

if arquivo:
    with st.spinner('💎 Colorindo as abas e ajustando tudo...'):
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
                        
                        if tipo_robo == "Fornecedor":
                            val_deb, val_cre = -deb, cre
                        else:
                            val_deb, val_cre = deb, -cre
                            
                        dados.append({"Data": dt, "NF": nf, "Hist": hist, "Deb": val_deb, "Cred": val_cre})

            if f_cod and dados: banco[f_cod] = pd.DataFrame(dados)

            if banco:
                out = BytesIO()
                # Lista de cores bonitas para as abas
                cores_abas = ['#FF0000', '#00FF00', '#0000FF', '#FFFF00', '#FF00FF', '#00FFFF', '#FFA500', '#800080', '#008000']
                
                with pd.ExcelWriter(out, engine='xlsxwriter') as writer:
                    wb = writer.book
                    wb.set_custom_property('ignore_errors', True) 
                    
                    f_cent = wb.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1})
                    f_moeda = wb.add_format({'num_format': 'R$ #,##0.00', 'border': 1})
                    f_std = wb.add_format({'border': 1})
                    f_cab = wb.add_format({'bold': 1, 'bg_color': '#F2F2F2', 'align': 'center', 'valign': 'vcenter', 'border': 1})
                    f_empresa = wb.add_format({'bold': 1, 'font_size': 14, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#D3D3D3', 'border': 1})
                    f_vde = wb.add_format({'num_format': 'R$ #,##0.00', 'font_color': 'green', 'bold': 1, 'border': 1, 'align': 'center'})
                    f_vrm = wb.add_format({'num_format': 'R$ #,##0.00', 'font_color': 'red', 'bold': 1, 'border': 1, 'align': 'center'})

                    for idx, (cod, df) in enumerate(banco.items()):
                        ws = wb.add_worksheet(str(cod)[:31])
                        
                        # 🪄 MÁGICA: COLORIR A ABA
                        cor_escolhida = cores_abas[idx % len(cores_abas)]
                        ws.set_tab_color(cor_escolhida)
                        
                        ws.hide_gridlines(2)
                        ws.ignore_errors({'number_stored_as_text': 'A1:X1000'})
                        
                        ws.set_column('A:A', 1)
                        ws.set_column('B:C', 15)
                        ws.set_column('D:D', 45)
                        ws.set_column('E:F', 18)
                        ws.set_column('G:H', 1)
                        ws.set_column('I:L', 18)
                        
                        # Alturas solicitadas
                        ws.set_row(0, 9)   # Linha 1: 9
                        ws.set_row(1, 22)  # Linha 2: Empresa
                        ws.set_row(2, 12)  # Linha 3: 12
                        ws.set_row(3, 20)  # Linha 4: Info Forn/Conc
                        ws.set_row(4, 15)  # Linha 5: 15
                        
                        ws.merge_range('B2:L2', f"EMPRESA: {nome_emp}", f_empresa)
                        ws.merge_range('B4:F4', f_info[cod], f_cab)
                        ws.merge_range('I4:L4', "CONCILIAÇÃO POR NOTA", f_cab)
                        
                        for ci, v in enumerate(["Data","NF","Histórico","Débito","Crédito"]):
                            ws.write(5, ci+1, v, f_cab)
                        
                        for ri, row in enumerate(df.values):
                            ws.write(6+ri, 1, row[0], f_cent); ws.write(6+ri, 2, row[1], f_cent)
                            ws.write(6+ri, 3, row[2], f_std); ws.write(6+ri, 4, row[3], f_moeda); ws.write(6+ri, 5, row[4], f_moeda)
                        
                        lt = 6 + len(df) + 1 
                        ws.write(lt, 3, "TOTALIZADOR:", f_cab)
                        ws.write(lt, 4, df['Deb'].sum(), f_moeda)
                        ws.write(lt, 5, df['Cred'].sum(), f_moeda)

                        res = df.groupby("NF").agg({"Deb":"sum","Cred":"sum"}).reset_index()
                        res["Dif"] = res["Deb"] + res["Cred"]
                        for ci, v in enumerate(["NF","Deb","Cred","Dif"]):
                            ws.write(5, ci+8, v, f_cab)
                        for ri, row in enumerate(res.values):
                            ws.write(6+ri, 8, str(row[0]), f_cent)
                            ws.write(6+ri, 9, row[1], f_moeda); ws.write(6+ri, 10, row[2], f_moeda); ws.write(6+ri, 11, row[3], f_moeda)
                        
                        rf = 7 + len(res)
                        ws.write(rf, 10, "Saldo Final:", f_cab)
                        ws.write(rf, 11, s := res["Dif"].sum(), f_vde if s >= 0 else f_vrm)
                
                st.success("✅ Arco-íris ativado! As abas estão coloridas e as alturas ajustadas.")
                st.download_button("📥 Baixar Relatório Colorido", out.getvalue(), "relatorio_lapidado.xlsx")
        except Exception as e:
            st.error(f"Erro: {e}")
