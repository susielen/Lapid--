import streamlit as st
import pandas as pd
import re
from io import BytesIO

st.set_page_config(page_title="Conciliador Mágico", layout="wide")

st.title("🤖 Super Robô Conciliador")
st.write("Suba o Razão (Excel ou CSV) e baixe o resultado!")

arquivo = st.file_uploader("Escolha o arquivo do Domínio", type=["xlsx", "csv"])

if arquivo:
    # 1. Lendo o arquivo (Excel ou CSV)
    if arquivo.name.endswith('.csv'):
        df = pd.read_csv(arquivo, skip_blank_lines=True)
    else:
        df = pd.read_excel(arquivo, engine='openpyxl')
    
    banco_fornecedores = {}
    fornecedor_atual = None
    dados_acumulados = []

    # 2. O Robô Detetive separa e limpa os dados
    for _, linha in df.iterrows():
        primeira_celula = str(linha.iloc[0]).strip()
        
        if primeira_celula.startswith("Conta:"):
            if fornecedor_atual and dados_acumulados:
                banco_fornecedores[fornecedor_atual] = pd.DataFrame(dados_acumulados)
            fornecedor_atual = str(linha.iloc[5]) if len(linha) > 5 and pd.notna(linha.iloc[5]) else str(linha.iloc[2])
            dados_acumulados = []
            continue
        
        if pd.notna(linha.iloc[0]) and any(char.isdigit() for char in str(linha.iloc[0])):
            # Abreviando a data (deixa como 10/05/24)
            data_obj = pd.to_datetime(linha.iloc[0])
            data_abreviada = data_obj.strftime('%d/%m/%y')
            
            deb = float(str(linha.iloc[8]).replace(',', '.')) if pd.notna(linha.iloc[8]) else 0
            cre = float(str(linha.iloc[9]).replace(',', '.')) if pd.notna(linha.iloc[9]) else 0
            
            hist = str(linha.iloc[2])
            nfe = re.findall(r'NFe\s?(\d+)', hist)
            num_nota = nfe[0] if nfe else "S/N"
            
            dados_acumulados.append({
                "Data": data_abreviada,
                "NF": num_nota,
                "Histórico": hist,
                "Débito": deb,
                "Crédito": cre
            })

    if fornecedor_atual and dados_acumulados:
        banco_fornecedores[fornecedor_atual] = pd.DataFrame(dados_acumulados)

    # 3. Criando as Abas e Visualização
    if banco_fornecedores:
        nomes = list(banco_fornecedores.keys())
        abas = st.tabs(nomes)
        
        # Preparando o arquivo para download
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            for i, nome in enumerate(nomes):
                df_razao = banco_fornecedores[nome]
                
                # Conciliação
                df_conc = df_razao.groupby("NF").agg({"Débito": "sum", "Crédito": "sum"}).reset_index()
                df_conc["Diferença"] = df_conc["Débito"] - df_conc["Crédito"]
                df_conc["Status"] = df_conc["Diferença"].apply(lambda x: "OK" if abs(x) < 0.01 else "Divergente")

                # Mostrando na aba do Streamlit
                with abas[i]:
                    col_esq, col_pulo, col_dir = st.columns([1.5, 0.2, 1])
                    with col_esq:
                        st.write("**📄 Razão**")
                        st.dataframe(df_razao, use_container_width=True, hide_index=True)
                    with col_dir:
                        st.write("**⚖️ Conciliação**")
                        st.dataframe(df_conc, use_container_width=True, hide_index=True)
                
                # Salvando no arquivo Excel (Lado a Lado com pulo de 3 colunas)
                df_razao.to_excel(writer, sheet_name=nome[:31], index=False, startcol=0)
                df_conc.to_excel(writer, sheet_name=nome[:31], index=False, startcol=len(df_razao.columns) + 3)

        st.divider()
        st.download_button(
            label="📥 Baixar Conciliação em Excel",
            data=output.getvalue(),
            file_name="conciliacao_final.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
