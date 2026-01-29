import streamlit as st
import pandas as pd
import re
from io import BytesIO

st.set_page_config(page_title="Conciliador Mestre", layout="wide")

st.title("🤖 Robô Conciliador (Visual Espaçado)")

arquivo = st.file_uploader("Suba o Razão (Excel ou CSV)", type=["xlsx", "csv"])

if arquivo:
    try:
        # 1. Leitura do arquivo
        if arquivo.name.endswith('.csv'):
            df_bruto = pd.read_csv(arquivo, header=None)
        else:
            df_bruto = pd.read_excel(arquivo, engine='openpyxl', header=None)
        
        # Busca nome da empresa no topo
        nome_empresa = "EMPRESA NÃO IDENTIFICADA"
        for i in range(min(10, len(df_bruto))):
            txt = str(df_bruto.iloc[i, 0])
            if "Empresa:" in txt or "EMPRESA:" in txt.upper():
                nome_empresa = str(df_bruto.iloc[i, 2])
                break

        banco_fornecedores = {}
        fornecedor_atual = None
        dados_acumulados = []

        # 2. Processamento dos Dados
        for i in range(len(df_bruto)):
            linha = df_bruto.iloc[i]
            col0 = str(linha[0]).strip() if pd.notna(linha[0]) else ""

            if "Conta:" in col0:
                if fornecedor_atual and dados_acumulados:
                    banco_fornecedores[fornecedor_atual] = pd.DataFrame(dados_acumulados)
                
                # Pega código e nome
                codigo = str(linha.iloc[1]).strip() if pd.notna(linha.iloc[1]) else "000"
                nome_forn = "Sem Nome"
                for c in [5, 6, 2]:
                    if len(linha) > c and pd.notna(linha.iloc[c]) and str(linha.iloc[c]).strip() != "":
                        nome_forn = str(linha.iloc[c]).strip()
                        break
                
                fornecedor_atual = f"{codigo} - {nome_forn}"
                dados_acumulados = []
                continue
            
            try:
                if len(linha) > 9 and (pd.notna(linha[8]) or pd.notna(linha[9])):
                    def conv(v):
                        try: return float(str(v).replace(',', '.')) if pd.notna(v) else 0.0
                        except: return 0.0
                    
                    d_orig, c_orig = conv(linha[8]), conv(linha[9])
                    if d_orig == 0 and c_orig == 0: continue

                    hist = str(linha[2])
                    nfe = re.findall(r'NFe\s?(\d+)', hist)
                    num_nota = nfe[0] if nfe else str(linha[1])

                    dados_acumulados.append({
                        "Data": str(linha[0]), "NF": num_nota, "Histórico": hist, 
                        "Débito": -d_orig, "Crédito": c_orig
                    })
            except: continue

        if fornecedor_atual and dados_acumulados:
            banco_fornecedores[fornecedor_atual] = pd.DataFrame(dados_acumulados)

        # 3. Geração do Excel com Nova Formatação
        if banco_fornecedores:
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                workbook = writer.book
                
                # Formatos
                fmt_empresa = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#D3D3D3', 'border': 1, 'text_wrap': True})
                fmt_forn_esq = workbook.add_format({'bold': True, 'align': 'left'})
                fmt_moeda = workbook.add_format({'num_format': '_-R$ * #,##0.00_-;-R$ * #,##0.00_-;_-R$ * "-"??_-;_-@_-', 'border': 1})
                fmt_verde = workbook.add_format({'num_format': '_-R$ * #,##0.00_-', 'font_color': 'green', 'bold': True, 'border': 1})
                fmt_vermelho = workbook.add_format({'num_format': '-R$ * #,##0.00_-', 'font_color': 'red', 'bold': True, 'border': 1})
                fmt_negrito = workbook.add_format({'bold': True, 'border': 1})

                for f_id_nome, df_f in banco_fornecedores.items():
                    aba = "".join(c for c in f_id_nome[:31] if c.isalnum() or c in " -")
                    
                    # TABELAS NA LINHA 10 (startrow=9)
                    df_f.to_excel(writer, sheet_name=aba, startrow=9, startcol=1, index=False)
                    df_res = df_f.groupby("NF").agg({"Débito":"sum", "Crédito":"sum"}).reset_index()
                    df_res["Diferença"] = df_res["Débito"] + df_res["Crédito"]
                    
                    col_res_idx = len(df_f.columns) + 4
                    df_res.to_excel(writer, sheet_name=aba, startrow=9, startcol=col_res_idx, index=False)
                    
                    ws = writer.sheets[aba]
                    ws.set_column('A:A', 2)
                    ws.set_row(0, 5)
                    
                    # Título da Empresa em 2 linhas (Linha 2 e 3)
                    ws.merge_range('B2:M3', f"EMPRESA: {nome_empresa}", fmt_empresa)
                    
                    # --- FORNECEDOR: Alinhado à esquerda na linha 8 (2 linhas acima da tabela) ---
                    ws.write('B8', f"FORNECEDOR: {f_id_nome}", fmt_forn_esq)

                    # TOTAIS: Linha 9 (acima dos cabeçalhos)
                    ws.write('D9', 'Totais', fmt_negrito)
                    ws.write('E9', df_f['Débito'].sum(), fmt_moeda)
                    ws.write('F9', df_f['Crédito'].sum(), fmt_moeda)

                    # Saldo no Final
                    row_f = 10 + len(df_res)
                    col_m_idx = col_res_idx + 3
                    saldo = df_res['Diferença'].sum()
                    ws.write(row_f, col_m_idx-1, "Saldo Final:", fmt_negrito)
                    ws.write(row_f, col_m_idx, saldo, fmt_verde if saldo >= 0 else fmt_vermelho)

                    ws.set_column(1, 25, 15, fmt_moeda)

            st.success("✅ Relatório formatado e espaçado!")
            st.download_button("📥 Baixar Excel Luxo", output.getvalue(), "conciliacao_profissional.xlsx")

    except Exception as e:
        st.error(f"Erro na visualização: {e}")
