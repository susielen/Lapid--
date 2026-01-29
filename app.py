import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Conciliador Domínio", page_icon="🤖")
st.title("🤖 Conciliador de Fornecedores")
st.write("Suba o arquivo do seu Razão e eu organizo tudo!")

arquivo = st.file_uploader("Arraste o arquivo Razão aqui", type=["csv", "xlsx"])

if arquivo is not None:
    try:
        # Lê o arquivo. Se for CSV do Domínio, geralmente usa encoding latin-1
        if arquivo.name.endswith('.xlsx'):
            df = pd.read_excel(arquivo)
        else:
            df = pd.read_csv(arquivo, encoding='latin-1', sep=None, engine='python')

        dados = []
        fornecedor_atual = "Não Identificado"

        # O robô vai percorrer linha por linha
        for i, linha in df.iterrows():
            linha_texto = " ".join([str(val) for val in linha.values]).upper()
            
            # 1. Identifica o Fornecedor (Linha que contém 'CONTA:' ou o código '1.01')
            if "CONTA:" in linha_texto or "NOME:" in linha_texto:
                fornecedor_atual = linha_texto.split("CONTA:")[-1].strip()
                # Limpa excessos como códigos numéricos no final
                if "NOME:" in fornecedor_atual:
                    fornecedor_atual = fornecedor_atual.split("NOME:")[-1].strip()
                continue

            # 2. Só processa valores se a linha tiver uma DATA (evita lixo e totais)
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

            st.success("✅ Agora sim! Processado com sucesso!")
            st.dataframe(resumo.style.format({'Débito': 'R$ {:.2f}', 'Crédito': 'R$ {:.2f}', 'Saldo Final': 'R$ {:.2f}'}))

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                resumo.to_excel(writer, index=False)
            
            st.download_button(
                label="📥 Baixar Resultado em Excel",
                data=output.getvalue(),
                file_name="conciliacao_final.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("❌ Não encontrei lançamentos válidos. Verifique se o arquivo está no formato correto.")

    except Exception as e:
        st.error(f"Erro técnico: {e}")
