import streamlit as st
import pandas as pd
import re
from io import BytesIO
import time

# 1. Configuração da Página
st.set_page_config(page_title="LAPIDÔ", page_icon="💎", layout="wide")

# 2. Estilo do Título
st.markdown("""
    <style>
    .titulo {
        color: #1E90FF;
        font-size: 48px;
        font-weight: bold;
        text-align: center;
        padding: 20px;
    }
    </style>
    <p class="titulo">💎 LAPIDÔ: O Mestre das Contas</p>
    """, unsafe_allow_html=True)

def to_num(val):
    try:
        if pd.isna(val) or str(val).strip().lower() in ['nan', 'null', '']: return 0.0
        v = str(val).replace('.', '').replace(',', '.')
        return float(v)
    except: return 0.0

# 3. Painel Lateral
with st.sidebar:
    st.header("⚙️ Painel de Controle")
    arquivo = st.file_uploader("Suba seu arquivo aqui", type=["xlsx", "csv"])
    st.divider()
    st.info("As colunas G e H agora estão bem fininhas!")

# 4. Processamento
if not arquivo:
    st.warning("👈 Coloque o arquivo na barra lateral para começar.")
else:
    with st.spinner('💎 Polindo o diamante... Ajustando as colunas...'):
        try:
            time.sleep(1.2) 
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
                        if hist.upper() in ['N', 'NAN', '', 'TOTAL', 'SUBTOTAL']: continue
                        try: dt = pd.to_datetime(lin[0]).strftime('%d/%m/%Y')
                        except: continue
                        nf_find = re.findall(r'NFe\s?(\d+)', hist)
                        nf_final = nf_find[0] if nf_find else str(lin[1])
                        dados.append({"Data": dt, "NF": nf_final, "Hist": hist, "Deb": -deb, "Cred": cre})

            if f_cod_atual and dados: banco[f_cod_atual] = pd.DataFrame(dados)

            if banco:
                out = BytesIO()
                with pd.ExcelWriter(out, engine='xlsxwriter') as writer:
                    wb = writer.book
                    f_tit = wb.add_format({'bold':1,'align':'center','valign':'vcenter','bg_color':'#D3D3D3','border':1})
                    f_std = wb.add_format({'border':1})
                    f_cen = wb.add_format({'border':1, 'align':'center'})
                    f_con = wb.add_format({'num_format': '_-R$ * #,##0.00_-;-R$ * #,##0.00_-;_-R
