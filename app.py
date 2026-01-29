import streamlit as st
import pandas as pd
import io, re

st.set_page_config(page_title="Conciliador Original", layout="wide")
st.title("🔙 Versão Original (Com Data Limpa)")

arquivo = st.file_uploader("Suba o arquivo Razão aqui", type=["csv", "xlsx"])

if arquivo:
    try:
        # 1. LER O ARQUIVO EXATAMENTE COMO ANTES
        if arquivo.name.endswith('.xlsx'):
            df = pd.read_excel(arquivo)
        else:
            df = pd.read_csv(arquivo, encoding='latin-1', sep=None, engine='python')

        # 2. LIMPAR A DATA (SÓ TIRA O HORÁRIO, MAIS NADA)
        # Se a data for "2025-01-01 00:00:00", vira "2025-01-01"
        df.iloc[:, 0] = df.iloc[:, 0].astype(str).str.split(' ').str[0]

        # 3. EXTRAIR A NOTA (A PEDIDO SEU)
        def pegar_nota(texto):
            busca = re.search(r'(?:NFE|NF|NOTA|Nº)\s*(\d+)', str(texto).upper())
            return busca.group(1) if busca else ""

        # Aqui o robozinho só organiza as gavetas
        dados_finais = []
        for i, linha in df.iterrows():
            # Se for uma linha com valor, a gente guarda
            if pd.notna(linha.iloc[0]) and ("/" in str(linha.iloc[0]) or "-" in str(linha.iloc[0])):
                hist = str(linha.iloc[2])
                dados_finais.append({
                    'Data': linha.iloc[0],
                    'Nota': pegar_nota(hist),
                    'Histórico': hist,
                    'Débito': linha.iloc[8] if len(linha) > 8 else 0,
                    'Crédito': linha.iloc[9] if len(linha) > 9 else 0
                })

        # 4. GERAR O EXCEL IGUAL AO QUE VOCÊ GOSTAVA
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            pd.DataFrame(dados_finais).to_excel(writer, index=False, sheet_name='Conciliacao')
        
        st.success("✅ Voltamos para a versão que você gosta!")
        st.download_button("📥 Baixar Arquivo", output.getvalue(), "conciliacao_certa.xlsx")

    except Exception as e:
        st.error(f"Erro ao carregar: {e}")
