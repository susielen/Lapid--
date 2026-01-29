import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Conciliador Domínio", page_icon="🤖")
st.title("🤖 Conciliador de Fornecedores")
st.write("Suba o arquivo do seu Razão e eu organizo tudo!")

arquivo = st.file_uploader("Arraste o arquivo Razão aqui", type=["csv", "xlsx"])

if arquivo is not None:
    try:
        # Tenta ler o arquivo de várias formas para não dar erro
        if arquivo.name.endswith('.xlsx'):
            df = pd.read_excel(arquivo)
        else:
            df = pd.read_csv(arquivo, encoding='latin-1', sep=None, engine='python')

        dados = []
        fornecedor_atual = "Não Identificado"

        # O robô vai percorrer linha por linha como se estivesse lendo um livro
        for i, linha in df.iterrows():
            linha_texto = " ".join([str(val) for val in linha.values]).upper()
            
            # 1. Se encontrar a palavra "CONTA" ou "NOME", ele guarda o nome do fornecedor
            if "CONTA:" in linha_texto or "NOME:" in linha_texto:
                partes = linha_texto.split("CONTA:")[-1]
                if "NOME:" in partes:
                    fornecedor_atual = partes.split("NOME:")[-1].strip()
                else:
                    fornecedor_atual = partes.strip()
                continue

            # 2. Só soma se a linha tiver uma DATA (evita somar os totais do sistema)
            tem_data = any("/20" in str(val) for val in linha.values[:3])
            
            if tem_data:
                try:
                    def limpar_valor(val):
                        if pd.isna(val): return 0
                        v = str(val).replace('.', '').replace(',', '.')
                        return pd.to_numeric(v, errors='coerce') or 0

                    # No seu arquivo, Débito e Crédito costumam ser as colunas 4 e 5
                    deb = limpar_valor(linha.iloc[4]) if len(linha) > 4 else 0
                    cre = limpar_valor(linha.iloc[5]) if len(linha) > 5 else 0
                    
                    if deb > 0 or cre > 0:
                        dados.append({
                            'Fornecedor': fornecedor_atual,
                            'Débito': deb,
                            'Crédito': cre
                        })
                except:
                    continue

        if dados:
            df_resumo = pd.DataFrame(dados)
            resumo = df_resumo.groupby('Fornecedor').agg({'Débito': 'sum', 'Crédito': 'sum'}).reset_index()
            resumo['Saldo Final'] = resumo['Crédito'] - resumo['Débito']

            st.success("✅ Agora sim! Consegui ler tudo!")
            st.dataframe(resumo.style.format({'Débito': 'R$ {:.2f}', 'Crédito': 'R$ {:.2f}', 'Saldo Final': 'R$ {:.2f}'}))

            # Prepara o botão de baixar
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                resumo.to_excel(writer, index=False)
            
            st.download_button(
                label="📥 Baixar Resultado em Excel",
                data=output.getvalue(),
                file_name="conciliacao_final.xlsx"
            )
        else:
            st.error("❌ O arquivo abriu, mas não encontrei nenhum valor de Débito ou Crédito nas linhas com data.")

    except Exception as e:
        st.error(f"Erro técnico: {e}")
