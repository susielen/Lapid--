import streamlit as st
import pandas as pd
import io
import re

st.set_page_config(page_title="Conciliador Pro", layout="wide")
st.title("🤖 Conciliador de Fornecedores (Modelo Razão/Conciliação)")

arquivo = st.file_uploader("Suba o arquivo Razão do Domínio aqui", type=["csv", "xlsx"])

def extrair_nota(texto):
    # Procura padrões de nota fiscal no histórico (ex: NFe 1234, NF 567)
    match = re.search(r'(?:NF|NFE|NF-E|NOTA|Nº)\s*(\d+)', str(texto).upper())
    return match.group(1) if match else "(vazio)"

if arquivo is not None:
    try:
        # 1. LEITURA DOS DADOS
        if arquivo.name.endswith('.xlsx'):
            df_raw = pd.read_excel(arquivo)
        else:
            df_raw = pd.read_csv(arquivo, encoding='latin-1', sep=None, engine='python')

        lista_razao = []
        fornecedor_atual = "Não Identificado"

        # 2. PROCESSAMENTO (ESTILO RAZÃO FORNECEDOR)
        for _, linha in df_raw.iterrows():
            texto_linha = " ".join([str(v) for v in linha.values]).upper()
            
            # Identifica novo fornecedor
            if "CONTA:" in texto_linha:
                fornecedor_atual = texto_linha.split("CONTA:")[-1].strip()
                continue
            
            # Verifica se é linha de lançamento (tem data)
            tem_data = any("/20" in str(v) for v in linha.values[:3])
            if tem_data:
                hist = str(linha.iloc[2]) # Coluna Histórico
                nf = extrair_nota(hist)
                
                # Limpeza de valores
                def limpar(v):
                    v = str(v).replace('.', '').replace(',', '.')
                    return pd.to_numeric(v, errors='coerce') or 0

                deb = limpar(linha.iloc[8]) # Débito costuma ser coluna 8 no Razão Domínio
                cre = limpar(linha.iloc[9]) # Crédito costuma ser coluna 9

                if deb > 0 or cre > 0:
                    lista_razao.append({
                        'Fornecedor': fornecedor_atual,
                        'Data': linha.iloc[0],
                        'Histórico': hist,
                        'Nº NF': nf,
                        'Débito': deb,
                        'Crédito': cre
                    })

        if lista_razao:
            df_razao = pd.DataFrame(lista_razao)

            # 3. CRIAÇÃO DA ABA CONCILIAÇÃO
            # Agrupamos por Fornecedor e Nota Fiscal
            df_conciliacao = df_razao.groupby(['Fornecedor', 'Nº NF']).agg({
                'Débito': 'sum',
                'Crédito': 'sum'
            }).reset_index()

            df_conciliacao['DIFERENÇA'] = df_conciliacao['Crédito'] - df_conciliacao['Débito']
            df_conciliacao['STATUS'] = df_conciliacao['DIFERENÇA'].apply(lambda x: "OK" if abs(x) < 0.01 else "DIVERGENTE")

            # EXIBIÇÃO NO SITE
            tab1, tab2 = st.tabs(["📋 Razão Detalhado", "⚖️ Conciliação (Status)"])
            
            with tab1:
                st.subheader("Visualização estilo 'Razão Fornecedor'")
                st.dataframe(df_razao)

            with tab2:
                st.subheader("Visualização estilo 'Conciliação'")
                st.dataframe(df_conciliacao)

            # DOWNLOAD
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_razao.to_excel(writer, sheet_name='RAZÃO FORNECEDOR', index=False)
                df_conciliacao.to_excel(writer, sheet_name='CONCILIAÇÃO', index=False)
            
            st.download_button("📥 Baixar Planilha Pronta", data=output.getvalue(), file_name="conciliacao_feita.xlsx")
        else:
            st.warning("Não foi possível identificar lançamentos no arquivo.")

    except Exception as e:
        st.error(f"Erro: {e}")
