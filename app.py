import streamlit as st
import pandas as pd
import re
from io import BytesIO

# 1. Configuração da Aba e do Ícone
st.set_page_config(page_title="LAPIDÔ", page_icon="💎", layout="wide")

# 2. Deixando o Título Lindo e Centralizado
st.markdown("""
    <style>
    .titulo {
        color: #1E90FF;
        font-size: 50px;
        font-weight: bold;
        text-align: center;
        text-shadow: 2px 2px #F0F0F0;
    }
    </style>
    <p class="titulo">💎 LAPIDÔ: O Mestre das Contas</p>
    """, unsafe_allow_html=True)

# 3. Função para transformar texto em número (o cérebro do robô)
def to_num(val):
    try:
        if pd.isna(val) or str(val).strip().lower() in ['nan', 'null', '']: return 0.0
        v = str(val).replace('.', '').replace(',', '.')
        return float(v)
    except: return 0.0

# 4. Gavetinha Lateral (Sidebar)
with st.sidebar:
    st.image("https://cdn-icons-png.flaticon.com/512/3067/3067451.png", width=100)
    st.header("📍 Painel de Controle")
    arquivo = st.file_uploader("Suba seu arquivo aqui", type=["xlsx", "csv"])
    st.divider()
    st.info("O LAPIDÔ transforma dados brutos em diamantes organizados.")

# 5. O Trabalho do Robô
if not arquivo:
    st.warning("👈 Ei! O LAPIDÔ está esperando o arquivo na gavetinha lateral.")
else:
    try:
        with st.spinner('Lapidando seu diamante...'):
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
                        # REGRA DE FORNECEDORES: Débito Negativo
                        dados.append({"Data": dt, "NF": nf_final, "Hist": hist, "Deb": -deb, "Cred": cre})

            if f_cod_atual and dados: banco[f_cod_atual] = pd.DataFrame(dados)

            if banco:
                out = BytesIO()
                with pd.ExcelWriter(out, engine='xlsxwriter') as writer:
                    wb = writer.book
                    f_tit = wb.add_format({'bold':1,'align':'center','valign':'vcenter','bg_color':'#D3D3D3','border':1, 'font_size': 14})
                    f_std = wb.add_format({'border':1})
                    f_cen = wb.add_format({'border':1, 'align':'center'})
                    f_con = wb.add_format({'num_format': '_-R$ * #,##0.00_-;-R$ * #,##0.00_-;_-R$ * "-"??_-;_-@_-','border': 1})
                    f_vde = wb.add_format({'num_format': '_-R$ * #,##0.00_-', 'font_color':'green','bold':1,'border':1})
                    f_vrm = wb.add_format({'num_format': '_-R$
