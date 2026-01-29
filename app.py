import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Conciliador Domínio", page_icon="🤖")
st.title("🤖 Conciliador de Fornecedores")

arquivo = st.file_uploader("Arraste o arquivo Razão aqui", type=["csv", "xlsx"])

if arquivo is not None:
    try:
        if arquivo.name.endswith('.xlsx'):
            df = pd.read_excel(arquivo)
        else:
            df = pd.read_csv(arquivo, encoding='latin-1', sep=None, engine='python')

        dados = []
        fornecedor_atual = "Não Identificado"

        for i, linha in df.iterrows():
            linha_lista = [str(val).strip().upper() for val in linha.values]
            linha_texto = " ".join(linha_lista)
            
            # 1. Identifica o Fornecedor
            if "CONTA:" in linha_texto:
                fornecedor_atual = linha_texto.split("CONTA:")[-1].strip()
                continue

            # 2. Procura por valores em QUALQUER coluna da linha
            # Mas só faz isso se a linha tiver uma data (00/00/0000)
            if any("/" in s and len(s) >= 8 for s in linha_lista[:4]):
                
                valores_da_linha = []
                for val in linha.values:
                    try:
                        # Limpa o valor (tira pontos de milhar e muda vírgula para ponto)
                        v_limpo = str(val).replace('.', '').replace(',', '.')
                        num = pd.to_numeric(v_limpo, errors='coerce')
                        if pd.notna(num) and num > 0:
                            valores_da_linha.append(num)
                    except:
                        continue
                
                # Se achamos dois números, o primeiro é Débito e o segundo é Crédito
                # Se achamos só um, precisamos decidir qual é (baseado na posição)
                if len(valores_da_linha) >= 1:
                    # No Razão, Débito vem antes de Crédito
                    # Vamos pegar os maiores valores encontrados na linha
                    deb = valores_da_linha[0] if len(valores_da_linha) >= 1 else 0
                    cre = valores_da_linha[1] if len(valores_da_linha) >= 2 else 0
                    
                    # Se só achou um valor, vamos checar em qual lado da linha ele estava
                    if len(valores_da_linha) == 1:
                        # Se o valor estava mais para o fim da linha, é crédito
                        posicao = list(linha.values).index(valores_da_linha[0])
                        if posicao > len(linha)/2:
                            cre = deb
                            deb = 0

                    dados.append({
                        'Fornecedor': fornecedor_atual,
                        'Débito': deb,
                        'Crédito': cre
                    })

        if dados:
            df_resumo = pd.DataFrame(dados)
            resumo = df_resumo.groupby('Fornecedor').agg({'Débito': 'sum', 'Crédito': 'sum'}).reset_index()
            resumo['Saldo Final'] = resumo['Crédito'] - resumo['Débito']

            st.success("✅ Consegui! Encontrei os valores.")
            st.dataframe(resumo.style.format({'Débito': 'R$ {:.2f}', 'Crédito': 'R$ {:.2f}', 'Saldo Final': 'R$ {:.2f}'}))

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                resumo.to_excel(writer, index=False)
            st.download_button("📥 Baixar Resultado", data=output.getvalue(), file_name="resumo.xlsx")
        else:
            st.error("❌ Ainda não encontrei valores. O arquivo parece não ter lançamentos de débito/crédito reconhecíveis.")

    except Exception as e:
        st.error(f"Erro: {e}")
