import streamlit as st
import pandas as pd
import re
from io import BytesIO
import time

# 1. Configuração da Página
st.set_page_config(page_title="LAPIDÔ", layout="wide")

# 2. Título (Ajuste a cor se for o de Fornecedores para #1E90FF)
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
    with st.spinner('💎 Lapidando e limpando as grades...'):
        try:
            time.sleep(1)
            df_bruto = pd.read_excel(arquivo, header=None) if arquivo.name.endswith('xlsx') else pd.read_csv(arquivo, header=None)
            
            nome_emp = "EMPRESA"
            for i in range(min(15, len(df_bruto))):
                if "Empresa:" in str(df_bruto.iloc[i, 0]):
                    nome_emp = str(df_bruto.iloc[i, 2])
                    break

            banco, f_cod, dados = {}, None, []
            for i in range(len(df_bruto)):
                lin = df_bruto.iloc[i]
                if "Conta:" in str(lin[0]):
                    if f_cod and dados: banco[f_cod] = pd.DataFrame(dados)
                    f_cod = str(lin[1]).strip()
                    dados = []
                elif len(lin) > 9:
                    deb, cre = to_num(lin[8]), to_num(lin[9])
                    hist = str(lin[2]).strip()
                    if (deb != 0 or cre != 0) and pd.notna(lin[0]):
                        if 'TOTAL' in hist.upper(): continue
                        dt = str(lin[0]) # Mantém a data como está
                        nf_f = re.findall(r'NFe\s?(\d+)', hist)
                        # IGNORA ERRO DE NF: Se não achar, usa o que estiver na coluna 1 sem reclamar
                        nf = nf_f[0] if nf_f else str(lin[1])
                        
                        # REGRA: Ajuste aqui se for Débito ou Crédito positivo conforme o robô
                        dados.append({"Data": dt, "NF": nf, "Hist": hist, "Deb": deb, "Cred": -cre})

            if f_cod and dados: banco[f_cod] = pd.DataFrame(dados)

            if banco:
                out = BytesIO()
                with pd.ExcelWriter(out, engine='xlsxwriter') as writer:
                    wb = writer.book
                    
                    # FORMATOS
                    f_moeda = wb.add_format({'num_format': 'R$ #,##0.00', 'border': 0})
                    f_cab = wb.add_format({'bold': 1, 'bg_color': '#F2F2F2', 'align': 'center', 'border': 0})
                    f_std = wb.add_format({'border': 0})
                    f_vde = wb.add_format({'num_format': 'R$ #,##0.00', 'font_color': 'green', 'bold': 1})
                    f_vrm = wb.add_format({'num_format': 'R$ #,##0.00', 'font_color': 'red', 'bold': 1})

                    for cod, df in banco.items():
                        ws = wb.add_worksheet(str(cod)[:31])
                        
                        # TIRAR AS GRADES (Linhas de fundo do Excel)
                        ws.hide_gridlines(2)
                        
                        # AJUSTE DE COLUNAS
                        ws.set_column('A:A', 1)      # COLUNA A FININHA
                        ws.set_column('B:C', 15)     # Data e NF
                        ws.set_column('D:D', 45)     # Histórico largo
                        ws.set_column('E:F', 18)     # Valores
                        ws.set_column('G:H', 1)      # G e H FININHAS
                        ws.set_column('I:L', 18)     # Resumo
                        
                        ws.merge_range('B2:F3', f"EMPRESA: {nome_emp}", f_cab)
                        
                        # Cabeçalho do Razão
                        for ci, v in enumerate(["Data","NF","Histórico","Débito","Crédito"]):
                            ws.write(6, ci+1, v, f_cab)
                        
                        for ri, row in enumerate(df.values):
                            ws.write(7+ri, 1, str(row[0]), f_std)
                            ws.write(7+ri, 2, str(row[1]), f_std)
                            ws.write(7+ri, 3, row[2], f_std)
                            ws.write(7+ri, 4, row[3], f_moeda)
                            ws.write(7+ri, 5, row[4], f_moeda)
                        
                        # TOTALIZADOR NO RAZÃO (DE VOLTA!)
                        lin_total = 7 + len(df)
                        ws.write(lin_total, 3, "TOTALIZADOR:", f_cab)
                        ws.write(lin_total, 4, df['Deb'].sum(), f_moeda)
                        ws.write(lin_total, 5, df['Cred'].sum(), f_moeda)
                        
                        # Resumo Lateral
                        res = df.groupby("NF").agg({"Deb":"sum","Cred":"sum"}).reset_index()
                        res["Dif"] = res["Deb"] + res["Cred"]
                        for ci, v in enumerate(["NF","Deb","Cred","Dif"]):
                            ws.write(6, ci+8, v, f_cab)
                        for ri, row in enumerate(res.values):
                            ws.write(7+ri, 8, str(row[0]), f_std)
                            ws.write(7+ri, 9, row[1], f_moeda)
                            ws.write(7+ri, 10, row[2], f_moeda)
                            ws.write(7+ri, 11, row[3], f_moeda)
                        
                        s = res["Dif"].sum(); rf = 8 + len(res)
                        ws.write(rf, 10, "Saldo Final:", f_cab)
                        ws.write(rf, 11, s, f_vde if s >= 0 else f_vrm)
                
                st.success("✅ Mestre, o relatório está limpo, sem grades e com totais!")
                st.download_button("📥 Baixar Planilha", out.getvalue(), "relatorio_lapidado.xlsx")
        except Exception as e:
            st.error(f"Erro: {e}")
