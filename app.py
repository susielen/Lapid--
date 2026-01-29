import streamlit as st
import pandas as pd
import re
from io import BytesIO

st.set_page_config(page_title="Conciliador Nota 10", layout="wide")

st.title("🤖 Robô Conciliador (Versão À Prova de Erros)")

arquivo = st.file_uploader("Suba o arquivo do Domínio", type=["xlsx", "csv"])

if arquivo:
    try:
        # 1. Lendo o arquivo com cuidado
        if arquivo.name.endswith('.csv'):
            df_bruto = pd.read_csv(arquivo, skip_blank_lines=True, header=None)
        else:
            df_bruto = pd.read_excel(arquivo, engine='openpyxl', header=None)
        
        # Tenta achar o nome da empresa (geralmente linha 0, coluna 2)
        try:
            nome_empresa = str(df_bruto.iloc[0, 2])
        except:
            nome_empresa = "MINHA EMPRESA"

        banco_fornecedores = {}
        fornecedor_atual = None
        dados_acumulados = []

        # 2. Processamento Linha por Linha
        for i in range(len(df_bruto)):
            linha = df_bruto.iloc[i]
            celula_0 = str(linha[0]).strip() if pd.notna(linha[0]) else ""
            
            # Identifica novo fornecedor
            if "Conta:" in celula_0:
                if fornecedor_atual and dados_acumulados:
                    banco_fornecedores[fornecedor_atual] = pd.DataFrame(dados_acumulados)
                
                # Pega o nome do fornecedor (tenta colunas 5, 2 ou 1)
                fornecedor_atual = "Desconhecido"
                for col_idx in [5, 2, 1]:
                    if len(linha) > col_idx and pd.notna(linha[col_idx]):
                        fornecedor_atual = str(linha[col_idx])
                        break
                dados_acumulados = []
                continue
            
            # Verifica se é uma linha de movimento (tem que ter algo na coluna de débito ou crédito)
            try:
                if len(linha) > 9 and (pd.notna(linha[8]) or pd.notna(linha[9])):
                    # Trata a Data (Abreviada ou como texto se der erro)
                    data_raw = str(linha[0])
                    try:
                        data_exibicao = pd.to_datetime(linha[0]).strftime('%d/%m/%y')
                    except:
                        data_exibicao = data_raw if data_raw != "nan" else ""

                    # Trata Débito e Crédito
                    def limpa_valor(v):
                        try:
                            return float(str(v).replace(',', '.')) if pd.notna(v) else 0.0
                        except: return 0.0

                    deb = limpa_valor(linha[8])
                    cre = limpa_valor(linha[9])
                    
                    if deb == 0 and cre == 0: continue # Pula linha vazia

                    # Busca NF no histórico
                    hist = str(linha[2])
                    nfe = re.findall(r'NFe\s?(\d+)', hist)
                    num_nota = nfe[0] if nfe else str(linha[1])
                    
                    dados_acumulados.append({
                        "Data": data_exibicao,
                        "NF": num_nota,
                        "Histórico": hist,
                        "Débito": deb,
                        "Crédito": cre
                    })
            except:
                continue # Se der erro em uma linha, pula para a próxima

        if fornecedor_atual and dados_acumulados:
            banco_fornecedores[fornecedor_atual] = pd.DataFrame(dados_acumulados)

        # 3. Gerar Excel e Mostrar na Tela
        if banco_fornecedores:
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                workbook = writer.book
                fmt_t = workbook.add_format({'bold': True, 'align': 'center', 'bg_color': '#EFEFEF', 'border': 1})
                
                abas = st.tabs(list(banco_fornecedores.keys()))
                
                for idx, (nome_forn, df_f) in enumerate(banco_fornecedores.items()):
                    # Conciliação
                    df_c = df_f.groupby("NF").agg({"Débito": "sum", "Crédito": "sum"}).reset_index()
                    df_c["Dif"] = df_c["Débito"] - df_c["Crédito"]
                    
                    with abas[idx]:
                        c1, c_vazia, c2 = st.columns([1.5, 0.2, 1])
                        c1.dataframe(df_f, use_container_width=True, hide_index=True)
                        c2.dataframe(df_c, use_container_width=True, hide_index=True)

                    # Excel Formatação
                    aba_nome = "".join(x for x in nome_forn[:31] if x.isalnum() or x in " -")
                    df_f.to_excel(writer, sheet_name=aba_nome, startrow=3, startcol=1, index=False)
                    df_c.to_excel(writer, sheet_name=aba_nome, startrow=3, startcol=len(df_f.columns)+4, index=False)
                    
                    ws = writer.sheets[aba_nome]
                    ws.set_column('A:A', 2)
                    ws.set_row(0, 5)
                    ws.merge_range('B2:M2', f"EMPRESA: {nome_empresa} | FORNECEDOR: {nome_forn}", fmt_t)

            st.download_button("📥 Baixar Excel Sem Erros", output.getvalue(), "conciliacao.xlsx")
        else:
            st.warning("O robô leu o arquivo, mas não encontrou fornecedores. Verifique se é o arquivo do Razão mesmo!")

    except Exception as e:
        st.error(f"Ocorreu um erro inesperado: {e}")
