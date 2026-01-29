import streamlit as st
import pandas as pd
import io
import re
from openpyxl.styles import Font

st.set_page_config(page_title="Conciliador Contábil Pro", layout="wide")
st.title("🤖 Conciliador: Organização de Topo Personalizada")

arquivo = st.file_uploader("Suba o Razão do Domínio aqui", type=["csv", "xlsx"])

def extrair_nfe(texto):
    match = re.search(r'(?:NFE|NF|NOTA|Nº)\s*(\d+)', str(texto).upper())
    return match.group(1) if match else "(vazio)"

def limpar_nome_simples(linha_txt):
    match_cod = re.search(r'CONTA:\s*(\d+)', linha_txt)
    codigo = match_cod.group(1) if match_cod else ""
    nome = linha_txt.split("CONTA:")[-1]
    nome = re.sub(r'(\d+\.)+\d+', '', nome) 
    nome = nome.replace(codigo, '').replace('NOME:', '').strip()
    return f"{codigo} - {nome}" if codigo else nome

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
                fornecedor_atual = limpar_nome_simples(linha_txt)
                dict_fornecedores[fornecedor_atual] = []
                continue
            
            data_orig = str(linha.iloc[0])
            if "/" in data_orig or (len(data_orig) >= 8 and "-" in data_orig):
                def limpar_num(v):
                    if pd.isna(v): return 0
                    v = str(v).replace('.', '').replace(',', '.')
                    return pd.to_numeric(v, errors='coerce') or 0
                
                deb = limpar_num(linha.iloc[8])
                cre = limpar_num(linha.iloc[9])

                if (deb > 0 or cre > 0) and fornecedor_atual:
                    hist = str(linha.iloc[2])
                    dict_fornecedores[fornecedor_atual].append({
                        'Data': data_orig, 'Nº NF': extrair_nfe(hist),
                        'Histórico': hist, 'Débito': deb, 'Crédito': cre
                    })

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            for forn, lancamentos in dict_fornecedores.items():
                if not lancamentos: continue
                
                df_f = pd.DataFrame(lancamentos)
                df_c = df_f.groupby('Nº NF').agg({'Débito': 'sum', 'Crédito': 'sum'}).reset_index()
                df_c['DIFERENÇA'] = df_c['Crédito'] - df_c['Débito']
                df_c['STATUS'] = df_c['DIFERENÇA'].apply(lambda x: "OK" if abs(x) < 0.05 else "DIVERGENTE")
                
                nome_aba = re.sub(r'[\\/*?:\[\]]', '', forn)[:31]
                
                # Inicia as tabelas na linha 5 para dar espaço ao cabeçalho novo
                df_f.to_excel(writer, sheet_name=nome_aba, index=False, startrow=4)
                df_c.to_excel(writer, sheet_name=nome_aba, index=False, startrow=4, startcol=8)
                
                sheet = writer.sheets[nome_aba]
                fmt_contabil = '_-R$ * #,##0.00_-;-R$ * #,##0.00_-;_-R$ * "-"??_-;_-@_-'
                negrito = Font(bold=True)
                cor_vermelha = Font(bold=True, color="FF0000")
                cor_verde = Font(bold=True, color="00B050")

                # --- LINHA 1: NOME DO FORNECEDOR ---
                sheet.cell(row=1, column=1, value=forn).font = negrito

                # --- LINHA 3: TOTAIS EM CIMA DAS COLUNAS ---
                # Totais do Razão
                sheet.cell(row=3, column=4, value="TOT. DÉBITO:").font = negrito
                c_deb_topo = sheet.cell(row=3, column=5, value=df_f['Débito'].sum())
                c_deb_topo.number_format = fmt_contabil
                c_deb_topo.font = negrito

                sheet.cell(row=3, column=6, value="TOT. CRÉDITO:").font = negrito
                c_cre_topo = sheet.cell(row=3, column=7, value=df_f['Crédito'].sum())
                c_cre_topo.number_format = fmt_contabil
                c_cre_topo.font = negrito

                # Saldo/Totalizador da Conciliação (Em cima da coluna Diferença - Col 12)
                saldo = df_f['Crédito'].sum() - df_f['Débito'].sum()
                sheet.cell(row=3, column=11, value="ABERTO:").font = negrito
                c_aberto = sheet.cell(row=3, column=12, value=saldo)
                c_aberto.number_format = fmt_contabil
                c_aberto.font = cor_vermelha if saldo < 0 else cor_verde

                # Formatação das colunas de valores no corpo
                for r in range(6, len(df_f) + 6):
                    sheet.cell(row=r, column=5).number_format = fmt_contabil
                    sheet.cell(row=r, column=6).number_format = fmt_contabil

                for r in range(6, len(df_c) + 6):
                    sheet.cell(row=r, column=10).number_format = fmt_contabil
                    sheet.cell(row=r, column=11).number_format = fmt_contabil
                    c_dif = sheet.cell(row=r, column=12)
                    c_dif.number_format = fmt_contabil
                    c_dif.font = Font(color="FF0000") if (c_dif.value or 0) < 0 else Font(color="00B050")
                    
                    c_status = sheet.cell(row=r, column=13)
                    c_status.font = Font(color="00B050") if c_status.value == "OK" else Font(color="FF0000")

        st.success("✅ Relatório formatado com sucesso!")
        st.download_button("📥 Baixar Planilha Final Ajustada", data=output.getvalue(), file_name="conciliacao_alinhada_topo.xlsx")
            
    except Exception as e:
        st.error(f"Erro: {e}")
