import streamlit as st
import pandas as pd
import re
from io import BytesIO

st.set_page_config(page_title="Conciliador Profissional", layout="wide")

st.title("🤖 Robô Conciliador com Formatação Especial")

arquivo = st.file_uploader("Suba o arquivo do Domínio", type=["xlsx", "csv"])

if arquivo:
    # 1. Lendo o arquivo e pegando o nome da empresa
    if arquivo.name.endswith('.csv'):
        df_bruto = pd.read_csv(arquivo)
    else:
        df_bruto = pd.read_excel(arquivo, engine='openpyxl')
    
    # Tenta pegar o nome da empresa (geralmente está na primeira linha)
    nome_empresa = str(df_bruto.iloc[0, 2]) if not df_bruto.empty else "CONCILIAÇÃO"

    banco_fornecedores = {}
    fornecedor_atual = None
    dados_acumulados = []

    # 2. Processamento (Ignorando erros de data e NF)
    for _, linha in df_bruto.iterrows():
        celula_0 = str(linha.iloc[0]).strip()
        
        if celula_0.startswith("Conta:"):
            if fornecedor_atual and dados_acumulados:
                banco_fornecedores[fornecedor_atual] = pd.DataFrame(dados_acumulados)
            fornecedor_atual = str(linha.iloc[5]) if len(linha) > 5 and pd.notna(linha.iloc[5]) else str(linha.iloc[2])
            dados_acumulados = []
            continue
        
        # O robô agora aceita qualquer linha que pareça ter valores, mesmo com erro na data
        if pd.notna(linha.iloc[8]) or pd.notna(linha.iloc[9]):
            try:
                # Tenta formatar a data, se der erro, deixa como está
                data_val = pd.to_datetime(linha.iloc[0], errors='ignore')
                data_exibicao = data_val.strftime('%d/%m/%y') if hasattr(data_val, 'strftime') else str(linha.iloc[0])
            except:
                data_exibicao = str(linha.iloc[0])

            hist = str(linha.iloc[2])
            nfe = re.findall(r'NFe\s?(\d+)', hist)
            num_nota = nfe[0] if nfe else str(linha.iloc[1]) # Se não achar no texto, pega da coluna de NF
            
            deb = float(str(linha.iloc[8]).replace(',', '.')) if pd.notna(linha.iloc[8]) else 0
            cre = float(str(linha.iloc[9]).replace(',', '.')) if pd.notna(linha.iloc[9]) else 0
            
            dados_acumulados.append({
                "Data": data_exibicao,
                "NF": num_nota,
                "Histórico": hist,
                "Débito": deb,
                "Crédito": cre
            })

    if fornecedor_atual and dados_acumulados:
        banco_fornecedores[fornecedor_atual] = pd.DataFrame(dados_acumulados)

    # 3. Geração do Excel com a formatação que você pediu
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        
        # Formatos
        formato_titulo = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'font_size': 14, 'bg_color': '#D3D3D3'})
        
        for nome_forn, df_f in banco_fornecedores.items():
            # Criar aba
            aba_nome = nome_forn[:31].replace('/', '-')
            df_f.to_excel(writer, sheet_name=aba_nome, startrow=3, startcol=1, index=False)
            worksheet = writer.sheets[aba_nome]
            
            # --- PEDIDO: Coluna A em branco e fina ---
            worksheet.set_column('A:A', 2) 
            
            # --- PEDIDO: Linha 1 em branco e fina ---
            worksheet.set_row(0, 5) 
            
            # --- PEDIDO: Nome da empresa na linha 2 mesclando até M ---
            worksheet.merge_range('B2:M2', f"EMPRESA: {nome_empresa} | FORNECEDOR: {nome_forn}", formato_titulo)
            worksheet.set_row(1, 30) # Linha 2 um pouco mais alta para o título
            
            # Ajustar colunas do Razão e Conciliação
            worksheet.set_column('B:G', 15)
            
            # Criar a Conciliação ao lado (pulando 3 colunas)
            df_conc = df_f.groupby("NF").agg({"Débito": "sum", "Crédito": "sum"}).reset_index()
            start_col_conc = len(df_f.columns) + 4 # Coluna B(1) + colunas do df + 3 de espaço
            df_conc.to_excel(writer, sheet_name=aba_nome, startrow=3, startcol=start_col_conc, index=False)

    # Botão de download
    st.success("Tudo pronto! Seu arquivo está formatado.")
    st.download_button("📥 Baixar Excel Formatado", output.getvalue(), "conciliacao_nota_dez.xlsx")
