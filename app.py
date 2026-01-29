import streamlit as st
import pandas as pd
import io
import re

st.set_page_config(page_title="Conciliador Domínio", layout="wide")
st.title("🤖 Conciliador Estilo 'CONCILIAÇÃO NOVO'")

arquivo = st.file_uploader("Suba o Razão (Excel ou CSV) aqui", type=["csv", "xlsx"])

def extrair_nfe(texto):
    texto = str(texto).upper()
    # Busca números após NFe, NF, Nota ou Nº
    match = re.search(r'(?:NFE|NF|NOTA|Nº|NFE\s)\s*(\d+)', texto)
    if match:
        return match.group(1)
    return "(vazio)"

if arquivo is not None:
    try:
        # Lendo o arquivo sem pular linhas fixas (vamos procurar o cabeçalho)
        if arquivo.name.endswith('.xlsx'):
            df_raw = pd.read_excel(arquivo)
        else:
            df_raw = pd.read_csv(arquivo, encoding='latin-1', sep=None, engine='python')

        lista_razao = []
        fornecedor_atual = "Não Identificado"

        for i, linha in df_raw.iterrows():
            # Transforma a linha em texto para busca
            linha_txt = " ".join([str(v) for v in linha.values]).upper()
            
            # 1. Identifica o Fornecedor (procura pela palavra 'Conta:')
            if "CONTA:" in linha_txt:
                fornecedor_atual = linha_txt.split("CONTA:")[-1].strip()
                continue
            
            # 2. Identifica se é linha de valores (procura por data ex: 2025-01-03)
            # No Domínio, a data costuma estar na primeira coluna
            data_val = str(linha.iloc[0])
            if "/" in data_val or (len(data_val) >= 8 and "-" in data_val):
                
                hist = str(linha.iloc[2]) # Histórico geralmente é a 3ª coluna
                num_nf = extrair_nfe(hist)
                
                def converter_valor(val):
                    if pd.isna(val): return 0
                    v = str(val).replace('.', '').replace(',', '.')
                    return pd.to_numeric(v, errors='coerce') or 0

                # No seu arquivo TESTE: Débito é col 8, Crédito é col 9
                deb = converter_valor(linha.iloc[8]) if len(linha) > 8 else 0
                cre = converter_valor(linha.iloc[9]) if len(linha) > 9 else 0

                if deb > 0 or cre > 0:
                    lista_razao.append({
                        'Fornecedor': fornecedor_atual,
                        'Data': data_val,
                        'Nº NF': num_nf,
                        'Histórico': hist,
                        'Débito': deb,
                        'Crédito': cre
                    })

        if lista_razao:
            df_final_razao = pd.DataFrame(lista_razao)

            # Criando a Aba CONCILIAÇÃO (Agrupado por Fornecedor e NF)
            df_concilia = df_final_razao.groupby(['Fornecedor', 'Nº NF']).agg({
                'Débito': 'sum',
                'Crédito': 'sum'
            }).reset_index()
            
            df_concilia['DIFERENÇA'] = df_concilia['Crédito'] - df_concilia['Débito']
            df_concilia['STATUS'] = df_concilia['DIFERENÇA'].apply(lambda x: "OK" if abs(x) < 0.05 else "DIVERGENTE")

            # Exibição
            tab1, tab2 = st.tabs(["📄 Razão Processado", "⚖️ Aba Conciliação"])
            
            with tab1:
                st.dataframe(df_final_razao, use_container_width=True)
            
            with tab2:
                st.dataframe(df_concilia, use_container_width=True)

            # Botão de Download com as duas abas
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_final_razao.to_excel(writer, sheet_name='RAZÃO FORNECEDOR', index=False)
                df_concilia.to_excel(writer, sheet_name='CONCILIAÇÃO', index=False)
            
            st.download_button("📥 Baixar Planilha Conciliada", data=output.getvalue(), file_name="resultado_conciliacao.xlsx")
            
        else:
            st.error("❌ O robô leu o arquivo, mas não encontrou o padrão de 'Data' e 'Valores'. Verifique se é o Razão do Domínio.")

    except Exception as e:
        st.error(f"Erro ao processar: {e}")
