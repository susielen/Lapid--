import streamlit as st
import pandas as pd
import io
import re

st.set_page_config(page_title="Conciliador por Fornecedor", layout="wide")
st.title("🤖 Conciliador Inteligente: Razão + Conciliação Lado a Lado")

arquivo = st.file_uploader("Suba o Razão do Domínio aqui", type=["csv", "xlsx"])

def extrair_nfe(texto):
    match = re.search(r'(?:NFE|NF|NOTA|Nº)\s*(\d+)', str(texto).upper())
    return match.group(1) if match else "(vazio)"

if arquivo is not None:
    try:
        if arquivo.name.endswith('.xlsx'):
            df_raw = pd.read_excel(arquivo)
        else:
            df_raw = pd.read_csv(arquivo, encoding='latin-1', sep=None, engine='python')

        # Dicionário para guardar dados de cada fornecedor
        dict_fornecedores = {}
        fornecedor_atual = None

        for i, linha in df_raw.iterrows():
            linha_txt = " ".join([str(v) for v in linha.values]).upper()
            
            if "CONTA:" in linha_txt:
                fornecedor_atual = linha_txt.split("CONTA:")[-1].strip()
                if fornecedor_atual not in dict_fornecedores:
                    dict_fornecedores[fornecedor_atual] = []
                continue
            
            # Verifica se é linha de valores (tem data)
            data_val = str(linha.iloc[0])
            if "/" in data_val or (len(data_val) >= 8 and "-" in data_val):
                def limpar(v):
                    v = str(v).replace('.', '').replace(',', '.')
                    return pd.to_numeric(v, errors='coerce') or 0

                deb = limpar(linha.iloc[8]) if len(linha) > 8 else 0
                cre = limpar(linha.iloc[9]) if len(linha) > 9 else 0

                if (deb > 0 or cre > 0) and fornecedor_atual:
                    hist = str(linha.iloc[2])
                    dict_fornecedores[fornecedor_atual].append({
                        'Data': data_val,
                        'Nº NF': extrair_nfe(hist),
                        'Histórico': hist,
                        'Débito': deb,
                        'Crédito': cre
                    })

        if dict_fornecedores:
            # Seleção de Fornecedor no Site
            lista_nomes = sorted(list(dict_fornecedores.keys()))
            fornecedor_selecionado = st.selectbox("Selecione o Fornecedor para visualizar:", lista_nomes)

            # Prepara os dados do selecionado
            df_razao = pd.DataFrame(dict_fornecedores[fornecedor_selecionado])
            df_concilia = df_razao.groupby('Nº NF').agg({'Débito': 'sum', 'Crédito': 'sum'}).reset_index()
            df_concilia['DIFERENÇA'] = df_concilia['Crédito'] - df_concilia['Débito']
            df_concilia['STATUS'] = df_concilia['DIFERENÇA'].apply(lambda x: "OK" if abs(x) < 0.05 else "DIVERGENTE")

            # Layout Lado a Lado no Site
            col1, col2 = st.columns([2, 1]) # Razão maior que a Conciliação
            
            with col1:
                st.subheader(f"📋 Razão: {fornecedor_selecionado}")
                st.dataframe(df_razao, use_container_width=True)

            with col2:
                st.subheader("⚖️ Conciliação")
                st.dataframe(df_concilia, use_container_width=True)

            # GERAÇÃO DO EXCEL COM ABAS (RODAPÉS)
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                for forn in lista_nomes:
                    df_f = pd.DataFrame(dict_fornecedores[forn])
                    if not df_f.empty:
                        # Criar a tabela de conciliação para este fornecedor
                        df_c = df_f.groupby('Nº NF').agg({'Débito': 'sum', 'Crédito': 'sum'}).reset_index()
                        df_c['DIFERENÇA'] = df_c['Crédito'] - df_c['Débito']
                        df_c['STATUS'] = df_c['DIFERENÇA'].apply(lambda x: "OK" if abs(x) < 0.05 else "DIVERGENTE")
                        
                        # Escreve o Razão e a Conciliação ao lado no Excel
                        # Limita o nome da aba para 31 caracteres (regra do Excel)
                        nome_aba = (forn[:25] + '...') if len(forn) > 30 else forn
                        df_f.to_excel(writer, sheet_name=nome_aba, index=False, startrow=1)
                        df_c.to_excel(writer, sheet_name=nome_aba, index=False, startrow=1, startcol=7)
            
            st.download_button("📥 Baixar Excel com Abas por Fornecedor", data=output.getvalue(), file_name="conciliacao_completa.xlsx")
            
        else:
            st.error("Nenhum dado encontrado.")

    except Exception as e:
        st.error(f"Erro: {e}")
