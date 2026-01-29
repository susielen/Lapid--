import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Conciliador Domínio", page_icon="🤖")
st.title("🤖 Conciliador de Fornecedores")

arquivo = st.file_uploader("Arraste o arquivo Razão aqui", type=["csv", "xlsx"])

if arquivo is not None:
    try:
        # Esta parte é a que resolve o erro das fotos!
        if arquivo.name.endswith('.xlsx'):
            df = pd.read_excel(arquivo, skiprows=6)
        else:
            df = pd.read_csv(arquivo, skiprows=6, encoding='latin-1')

        df.columns = df.columns.str.strip()
        dados = []
        fornecedor_atual = "Não Identificado"

        for i, linha in df.iterrows():
            hist = str(linha.get('Histórico', ''))
            data_col = str(linha.get('Data', ''))
            
            if 'Conta:' in data_col or 'Conta:' in hist:
                fornecedor_atual = hist.strip()
                continue
            
            d = pd.to_numeric(linha.get('Débito', 0), errors='coerce')
            c = pd.to_numeric(linha.get('Crédito', 0), errors='coerce')
            
            if pd.notna(d) or pd.notna(c):
                dados.append({
                    'Fornecedor': fornecedor_atual,
                    'Débito': d if pd.notna(d) else 0,
                    'Crédito': c if pd.notna(c) else 0
                })

        if dados:
            df_final = pd.DataFrame(dados)
            resumo = df_final.groupby('Fornecedor').agg({'Débito': 'sum', 'Crédito': 'sum'}).reset_index()
            resumo['Saldo Final'] = resumo['Crédito'] - resumo['Débito']
            st.success("✅ Mágica feita!")
            st.dataframe(resumo)
            
            saida = io.BytesIO()
            with pd.ExcelWriter(saida, engine='openpyxl') as writer:
                resumo.to_excel(writer, index=False)
            st.download_button("📥 Baixar Excel", data=saida.getvalue(), file_name="conciliacao.xlsx")
    except Exception as e:
        st.error(f"Erro: {e}")
