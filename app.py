import streamlit as st
import pandas as pd
import re
from io import BytesIO

st.set_page_config(page_title="Conciliador Mestre", layout="wide")

st.title("🤖 Robô Conciliador com Formatação Premium")

arquivo = st.file_uploader("Suba o Razão (Excel ou CSV)", type=["xlsx", "csv"])

if arquivo:
    try:
        # 1. Leitura e Identificação da Empresa
        if arquivo.name.endswith('.csv'):
            df_bruto = pd.read_csv(arquivo, header=None)
        else:
            df_bruto = pd.read_excel(arquivo, engine='openpyxl', header=None)
        
        nome_empresa = "EMPRESA NÃO IDENTIFICADA"
        for i in range(min(10, len(df_bruto))):
            txt = str(df_bruto.iloc[i, 0])
            if "Empresa:" in txt or "EMPRESA:" in txt.upper():
                nome_empresa = str(df_bruto.iloc[i, 2]) if pd.notna(df_bruto.iloc[i, 2]) else nome_empresa
                break

        banco_fornecedores = {}
        fornecedor_atual = None
        dados_acumulados = []

        # 2. Processamento dos Dados
        for i in range(len(df_bruto)):
            linha = df_bruto.iloc[i]
            col0 = str(linha[0]).strip() if pd.notna(linha[0]) else ""

            if any(x in col0 for x in ["Empresa:", "C.N.P.J.:", "Período:", "RAZÃO"]):
                continue

            if "Conta:" in col0:
                if fornecedor_atual and dados_acumulados:
                    banco_fornecedores[fornecedor_atual] = pd.DataFrame(dados_acumulados)
                
                fornecedor_atual = "Fornecedor Sem Nome"
                for c in [5, 6, 7, 2, 1]:
                    if len(linha) > c and pd.notna(linha[c]) and str(linha[c]).strip() != "":
                        fornecedor_atual = str(linha[c]).strip()
                        break
                dados_acumulados = []
                continue
            
            try:
                if len(linha) > 9 and (pd.notna(linha[8]) or pd.notna(linha[9])):
                    try:
                        data_fmt = pd.to_datetime(linha[0]).strftime('%d/%m/%y')
                    except:
                        data_fmt = str(linha[0]) if col0 != "nan" else ""

                    def conv(v):
                        try: return float(str(v).replace(',', '.')) if pd.notna(v) else 0.0
                        except: return 0.0

                    d, c = conv(linha[8]), conv(linha[9])
                    if d == 0 and c == 0: continue

                    hist = str(linha[2])
                    nfe = re.findall(r'NFe\s?(\d+)', hist)
                    num_nota = nfe[0] if nfe else str(linha[1])

                    dados_acumulados.append({"Data": data_fmt, "NF": num_nota, "Histórico": hist, "Débito": d, "Crédito": c})
            except: continue

        if fornecedor_atual and dados_acumulados:
            banco_fornecedores[fornecedor_atual] = pd.DataFrame(dados_acumulados)

        # 3. Geração do Excel com a Nova Formatação
        if banco_fornecedores:
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                workbook = writer.book
                fmt_titulo = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#D3D3D3', 'border': 1})
                fmt_negrito = workbook.add_format({'bold': True, 'border': 1})
                fmt_dinheiro = workbook.add_format({'num_format': '#,##0.00', 'border': 1})
                fmt_negrito_centralizado = workbook.add_format({'bold': True, 'align': 'center', 'border': 1})

                for f_nome, df_f in banco_fornecedores.items():
                    aba_nome = "".join(c for c in f_nome[:31] if c.isalnum() or c in " ")
                    
                    # Coloca a tabela detalhada (Razão) na LINHA 7 (startrow=6)
                    df_f.to_excel(writer, sheet_name=aba_nome, startrow=6, startcol=1, index=False)
                    
                    # Coloca a tabela de Conciliação na LINHA 7 (startrow=6)
                    df_resumo = df_f.groupby("NF").agg({"Débito":"sum", "Crédito":"sum"}).reset_index()
                    df_resumo["Dif"] = df_resumo["Débito"] - df_resumo["Crédito"]
                    df_resumo.to_excel(writer, sheet_name=aba_nome, startrow=6, startcol=len(df_f.columns)+4, index=False)
                    
                    ws = writer.sheets[aba_nome]
                    ws.set_column('A:A', 2) # A fina
                    ws.set_row(0, 5)        # 1 fina
                    
                    # --- LINHA 2: Nome da Empresa ---
                    ws.merge_range('B2:M2', f"EMPRESA: {nome_empresa} | FORNECEDOR: {f_nome}", fmt_titulo)
                    
                    # --- LINHA 6: TOTAIS (Coluna D, E, F) ---
                    ws.write('D6', 'Totais', fmt_negrito)
                    ws.write('E6', df_f['Débito'].sum(), fmt_dinheiro)
                    ws.write('F6', df_f['Crédito'].sum(), fmt_dinheiro)
                    
                    # --- LINHA 6: SALDO (Mescla J a L, Valor na M) ---
                    # Calculando a posição das colunas J, L e M (J=9, L=11, M=12)
                    ws.merge_range('J6:L6', 'Saldo em Aberto', fmt_negrito_centralizado)
                    ws.write('M6', df_resumo['Dif'].sum(), fmt_dinheiro)

                # Mostra no Streamlit (opcional, para conferir)
                st.success("Planilha organizada com sucesso!")
                
            st.download_button("📥 Baixar Excel na Linha 7", output.getvalue(), "conciliacao_organizada.xlsx")

    except Exception as e:
        st.error(f"Erro ao organizar as linhas: {e}")
