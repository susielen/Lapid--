import streamlit as st
import pandas as pd
import io
import re
from datetime import datetime

st.set_page_config(page_title="Conciliador Profissional", layout="wide")
st.title("🤖 Conciliador: Razão + Conciliação (Com Totais)")

arquivo = st.file_uploader("Suba o Razão do Domínio aqui", type=["csv", "xlsx"])

def extrair_nfe(texto):
    match = re.search(r'(?:NFE|NF|NOTA|Nº)\s*(\d+)', str(texto).upper())
    return match.group(1) if match else "(vazio)"

def limpar_nome_fornecedor(nome):
    # Remove códigos tipo "2.1.1.1.01.0002" ou "Conta: 2113"
    nome_limpo = re.sub(r'(\d+\.)+\d+', '', nome)
    nome_limpo = nome_limpo.replace('CONTA:', '').replace('NOME:', '').strip()
    # Remove traços no início se houver
    if nome_limpo.startswith('-'): nome_limpo = nome_limpo[1:].strip()
    return nome_limpo

if arquivo is not None:
    try:
        if arquivo.name.endswith('.xlsx'):
            df_raw = pd.read_excel(arquivo)
        else:
            df_raw = pd.read_csv(arquivo, encoding='latin-1', sep=None, engine='python')

        dict_fornecedores = {}
        fornecedor_atual = None

        for i, linha in df_raw.iterrows():
            linha_txt = " ".join([str(v) for v in linha.values]).upper()
            
            if "CONTA:" in linha_txt:
                fornecedor_atual = limpar_nome_fornecedor(linha_txt)
                if fornecedor_atual not in dict_fornecedores:
                    dict_fornecedores[fornecedor_atual] = []
                continue
            
            # Identifica linha de valores pela data
            data_orig = str(linha.iloc[0])
            if "/" in data_orig or (len(data_orig) >= 8 and "-" in data_orig):
                try:
                    # Formata a data para DD/MM/YY
                    dt_obj = pd.to_datetime(data_orig)
                    data_formatada = dt_obj.strftime('%d/%m/%y')
                except:
                    data_formatada = data_orig

                def limpar_num(v):
                    v = str(v).replace('.', '').replace(',', '.')
                    return pd.to_numeric(v, errors='coerce') or 0

                deb = limpar_num(linha.iloc[8]) if len(linha) > 8 else 0
                cre = limpar_num(linha.iloc[9]) if len(linha) > 9 else 0

                if (deb > 0 or cre > 0) and fornecedor_atual:
                    hist = str(linha.iloc[2])
                    dict_fornecedores[fornecedor_atual].append({
                        'Data': data_formatada,
                        'Nº NF': extrair_nfe(hist),
                        'Histórico': hist,
                        'Débito': deb,
                        'Crédito': cre
                    })

        if dict_fornecedores:
            lista_nomes = sorted([n for n in dict_fornecedores.keys() if dict_fornecedores[n]])
            fornecedor_sel = st.selectbox("Selecione o Fornecedor:", lista_nomes)

            # --- PROCESSAMENTO DO SELECIONADO ---
            df_razao = pd.DataFrame(dict_fornecedores[fornecedor_sel])
            
            # Soma Totais para o final do Razão
            total_deb = df_razao['Débito'].sum()
            total_cre = df_razao['Crédito'].sum()
            saldo_final = total_cre - total_deb

            # Tabela de Conciliação lateral
            df_concilia = df_razao.groupby('Nº NF').agg({'Débito': 'sum', 'Crédito': 'sum'}).reset_index()
            df_concilia['DIFERENÇA'] = df_concilia['Crédito'] - df_concilia['Débito']
            df_concilia['STATUS'] = df_concilia['DIFERENÇA'].apply(lambda x: "OK" if abs(x) < 0.05 else "DIVERGENTE")

            # EXIBIÇÃO
            col1, col2 = st.columns([2, 1])
            with col1:
                st.subheader(f"📋 Razão: {fornecedor_sel}")
                st.dataframe(df_razao, use_container_width=True)
                # Mostra o rodapé com somas
                st.info(f"**TOTAIS:** Débito: R$ {total_deb:,.2f} | Crédito: R$ {total_cre:,.2f} | Saldo: R$ {saldo_final:,.2f}")

            with col2:
                st.subheader("⚖️ Conciliação")
                st.dataframe(df_concilia, use_container_width=True)

            # --- GERAÇÃO DO EXCEL ---
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                for forn in lista_nomes:
                    df_f = pd.DataFrame(dict_fornecedores[forn])
                    if not df_f.empty:
                        df_c = df_f.groupby('Nº NF').agg({'Débito': 'sum', 'Crédito': 'sum'}).reset_index()
                        df_c['DIFERENÇA'] = df_c['Crédito'] - df_c['Débito']
                        df_c['STATUS'] = df_c['DIFERENÇA'].apply(lambda x: "OK" if abs(x) < 0.05 else "DIVERGENTE")
                        
                        # Nome da aba
                        nome_aba = (forn[:25]) if len(forn) > 25 else forn
                        
                        # Escreve Razão
                        df_f.to_excel(writer, sheet_name=nome_aba, index=False, startrow=1)
                        
                        # Escreve Totais abaixo do Razão
                        sheet = writer.sheets[nome_aba]
                        row_fim = len(df_f) + 3
                        sheet.cell(row=row_fim, column=4, value="TOTAL:")
                        sheet.cell(row=row_fim, column=5, value=df_f['Débito'].sum())
                        sheet.cell(row=row_fim, column=6, value=df_f['Crédito'].sum())
                        sheet.cell(row=row_fim+1, column=5, value="DIFERENÇA:")
                        sheet.cell(row=row_fim+1, column=6, value=df_f['Crédito'].sum() - df_f['Débito'].sum())

                        # Escreve Conciliação ao lado
                        df_c.to_excel(writer, sheet_name=nome_aba, index=False, startrow=1, startcol=8)
            
            st.download_button("📥 Baixar Excel Completo (Com Abas e Totais)", data=output.getvalue(), file_name="conciliacao_total.xlsx")
            
    except Exception as e:
        st.error(f"Erro: {e}")
