import streamlit as st
import pandas as pd
import re
from io import BytesIO

# 1. Configuração da Aba e Ícone
st.set_page_config(page_title="LAPIDÔ", page_icon="💎", layout="wide")

# 2. Estilo do Título (Bonitão e sem festa)
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

def to_num(val):
    try:
        if pd.isna(val) or str(val).strip().lower() in ['nan', 'null', '']: return 0.0
        v = str(val).replace('.', '').replace(',', '.')
        return float(v)
    except: return 0.0

# 3. Painel Lateral (A gavetinha)
with st.sidebar:
    st.header("⚙️ Painel de Controle")
    arquivo = st.file_uploader("Suba seu arquivo aqui", type=["xlsx", "csv"])
    st.divider()
    st.info("O robô está pronto para polir seus dados.")

# 4. A Mágica da Lapidação
if not arquivo:
    st.warning("👈 Coloque a pedra bruta (arquivo) na gavetinha lateral.")
else:
    # EFEITO DE FORMANDO O DIAMANTE
    with st.spinner('💎 Polindo a pedra bruta... Transformando em diamante...'):
        try:
            # Simulando um tempinho para a animação aparecer (opcional)
            import time
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
                    deb, cre = to
