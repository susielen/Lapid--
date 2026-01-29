import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Conciliador Domínio", page_icon="🤖")
st.title("🤖 Conciliador de Fornecedores")
st.write("Suba o arquivo do seu Razão (CSV ou Excel) e eu organizo tudo!")

arquivo = st.file_uploader("Arraste o arquivo Razão aqui", type=["csv", "xlsx"])

if arquivo is not None:
    try:
        # Lê o arquivo ignorando as linhas iniciais de cabeçalho do sistema
        if arquivo.name.endswith('.xlsx'):
            df = pd.read_excel(arquivo, skiprows=5)
        else:
            df = pd.read_csv(arquivo, skiprows=5, encoding='latin-1')

        # Limpa os nomes das colunas
        df.columns = [str(c).strip() for c in df.columns]
        
        dados = []
        fornecedor_atual = None

        for i, linha in df.iterrows():
            # Transforma a linha em texto para facilitar a busca
            conteudo_linha = " ".join([str(val) for val in linha.values])
            
            # Identifica a linha que contém o Fornecedor (geralmente começa com o código da conta)
            if "Conta:" in conteudo_linha or ("1.01." in conteudo_linha and "Nome:" not in conteudo_linha):
                # Tenta pegar o nome que vem após o código ou palavra Conta
                fornecedor_atual = conteudo_linha.split("-")[-1].strip() if "-" in conteudo_linha else conteudo_linha
                continue
            
            # Pega os valores das colunas de Débito e Crédito (ajustado para as colunas do seu arquivo)
            # No seu arquivo, Débito costuma ser a 4ª ou 5ª coluna preenchida
            try:
                data = str(linha.get('Data', ''))
                # Só processa se houver uma data válida (evita linhas de totais)
                if "/" in data:
                    debito = pd.to_numeric(linha.get('Débito', 0), errors='coerce')
                    credito = pd.to_numeric(linha.get('Crédito', 0), errors='coerce')
                    
                    if (pd.notna(debito) and debito > 0) or (pd.notna(credito) and credito > 0):
                        dados.append({
                            'Fornecedor': fornecedor_atual if fornecedor_atual else "Outros",
                            'Débito': debito if pd.notna(debito) else 0,
                            'Crédito': credito if pd.notna(credito) else 0
                        })
            except:
                continue

        if dados:
            df_final = pd.DataFrame(dados)
            # Agrupa e soma
            resumo = df_final.groupby('Fornecedor').agg({'Débito': 'sum', 'Crédito': 'sum'}).reset_index()
            resumo['Saldo Final'] = resumo['Crédito'] - resumo['Débito']
            
            # Remove linhas que ficaram sem nome ou vazias
            resumo = resumo[resumo['Fornecedor'].str.len() > 3]

            st.success("✅ Agora sim! Veja os resultados:")
            st.dataframe(resumo.style.format({'Débito': 'R$ {:.2f}', 'Crédito': 'R$ {:.2f}', 'Saldo Final': 'R$ {:.2f}'}))

            # Botão para baixar
            saida = io.BytesIO()
            with pd.ExcelWriter(saida, engine='openpyxl') as writer:
                resumo.to_excel(writer, index=False)
            st.download_button("📥 Baixar Resumo Consolidado", data=saida.getvalue(), file_name="conciliacao_dominio.xlsx")
        else:
            st.warning("⚠️ O arquivo foi lido, mas não encontramos lançamentos de Débito/Crédito. Verifique se o arquivo está correto.")
            
    except Exception as e:
        st.error(f"Erro ao processar: {e}")
