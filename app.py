import streamlit as st
import pandas as pd
import io
import re

st.set_page_config(page_title="Conciliador Contábil Pro", layout="wide")
st.title("🤖 Conciliador: Razão e Conciliação Integrados")

arquivo = st.file_uploader("Suba o Razão do Domínio aqui", type=["csv", "xlsx"])

def extrair_nfe(texto):
    match = re.search(r'(?:NFE|NF|NOTA|Nº)\s*(\d+)', str(texto).upper())
    return match.group(1) if match else "(vazio)"

def extrair_codigo_e_nome(linha_txt):
    # Pega o código (ex: 2113) e o nome
    codigo = ""
    match_cod = re.search(r'CONTA:\s*(\d+)', linha_txt)
    if match_cod:
        codigo = match_cod.group(1)
    
    nome = linha_txt.split("CONTA:")[-1]
    nome = re.sub(r'(\d+\.)+\d+', '', nome) # tira o 2.1.1...
    nome = nome.replace(codigo, '').replace('NOME:', '').strip()
    if nome.startswith('-'): nome = nome[1:].strip()
    
    # Retorna o formato para o rodapé: "2113 - NOME"
    return f"{codigo} - {nome[:20]}" if codigo else nome[:25]

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
                fornecedor_atual = extrair_codigo_e_nome(linha_txt)
                if fornecedor_atual not in dict_fornecedores:
                    dict_fornecedores[fornecedor_atual] = []
                continue
            
            data_orig = str(linha.iloc[0])
            if "/" in data_orig or (len(data_orig) >= 8 and "-" in data_orig):
                try:
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

            # --- PROCESSAMENTO ---
            df_razao = pd.DataFrame(dict_fornecedores[fornecedor_sel])
            total_deb = df_razao['Débito'].sum()
            total_cre = df_razao['Crédito'].sum()
            saldo_aberto = total_cre - total_deb # Crédito (+) e Débito (-)

            df_concilia = df_razao.groupby('Nº NF').agg({'Débito': 'sum', 'Crédito': 'sum'}).reset_index()
            df_concilia['DIFERENÇA'] = df_concilia['Crédito'] - df_concilia['Débito']
            df_concilia['STATUS'] = df_concilia['DIFERENÇA'].apply(lambda x: "OK" if abs(x) < 0.05 else "DIVERGENTE")

            # --- EXIBIÇÃO NO SITE ---
            col1, col2 = st.columns([2, 1])
            with col1:
                st.subheader(f"📋 Razão: {fornecedor_sel}")
                st.dataframe(df_razao, use_container_width=True)
                st.markdown(f"**TOTAIS DO RAZÃO:**")
                st.write(f"Soma Débito: R$ {total_deb:,.2f} | Soma Crédito: R$ {total_cre:,.2f}")
                st.warning(f"**SALDO EM ABERTO: R$ {saldo_aberto:,.2f}**")

            with col2:
                st.subheader("⚖️ Conciliação")
                st.metric("Total em Aberto (NF)", f"R$ {saldo_aberto:,.2f}")
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
                        
                        # Nome da aba (já com código e nome)
                        nome_aba = re.sub(r'[\\/*?:\[\]]', '', forn)[:30]
                        df_f.to_excel(writer, sheet_name=nome_aba, index=False, startrow=1)
                        
                        sheet = writer.sheets[nome_aba]
                        row_totais = len(df_f) + 3 # Pula uma linha
                        
                        # Colocando totais embaixo das colunas Débito (5) e Crédito (6)
                        sheet.cell(row=row_totais, column=4, value="TOTAL:")
                        sheet.cell(row=row_totais, column=5, value=df_f['Débito'].sum())
                        sheet.cell(row=row_totais, column=6, value=df_f['Crédito'].sum())
                        
                        sheet.cell(row=row_totais + 1, column=4, value="SALDO:")
                        sheet.cell(row=row_totais + 1, column=6, value=df_f['Crédito'].sum() - df_f['Débito'].sum())

                        # Conciliação ao lado com Totalizador no topo
                        sheet.cell(row=1, column=9, value="TOTAL EM ABERTO:")
                        sheet.cell(row=1, column=10, value=df_f['Crédito'].sum() - df_f['Débito'].sum())
                        df_c.to_excel(writer, sheet_name=nome_aba, index=False, startrow=2, startcol=8)
            
            st.download_button("📥 Baixar Excel Final", data=output.getvalue(), file_name="conciliacao_dominio.xlsx")
            
    except Exception as e:
        st.error(f"Erro: {e}")
