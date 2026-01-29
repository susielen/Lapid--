import streamlit as st
import pandas as pd
import io
import re
from openpyxl.styles import Font, PatternFill

st.set_page_config(page_title="Conciliador Contábil Colorido", layout="wide")
st.title("🤖 Conciliador: Padrão Contábil com Cores")

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
                df_f.to_excel(writer, sheet_name=nome_aba, index=False, startrow=2)
                df_c.to_excel(writer, sheet_name=nome_aba, index=False, startrow=2, startcol=8)
                
                sheet = writer.sheets[nome_aba]
                fmt_contabil = '_-R$ * #,##0.00_-;-R$ * #,##0.00_-;_-R$ * "-"??_-;_-@_-'
                cor_vermelha = Font(color="FF0000")
                cor_verde = Font(color="00B050")
                negrito = Font(bold=True)

                # Formatação Razão
                for r in range(4, len(df_f) + 4):
                    sheet.cell(row=r, column=5).number_format = fmt_contabil
                    sheet.cell(row=r, column=6).number_format = fmt_contabil

                # Totais e Saldo
                row_tot = len(df_f) + 5
                sheet.cell(row=row_tot, column=4, value="TOTAIS:").font = negrito
                c_deb = sheet.cell(row=row_tot, column=5, value=df_f['Débito'].sum())
                c_cre = sheet.cell(row=row_tot, column=6, value=df_f['Crédito'].sum())
                c_deb.number_format = c_cre.number_format = fmt_contabil
                
                sheet.cell(row=row_tot+1, column=4, value="SALDO:").font = negrito
                saldo = df_f['Crédito'].sum() - df_f['Débito'].sum()
                c_saldo = sheet.cell(row=row_tot+1, column=6, value=saldo)
                c_saldo.number_format = fmt_contabil
                c_saldo.font = Font(bold=True, color="FF0000" if saldo < 0 else "00B050")

                # Totalizador Topo
                sheet.cell(row=1, column=11, value="TOTAL ABERTO:").font = negrito
                c_topo = sheet.cell(row=1, column=12, value=saldo)
                c_topo.number_format = fmt_contabil
                c_topo.font = Font(bold=True, color="FF0000" if saldo < 0 else "00B050")

                # Formatação Tabela Conciliação (Cores no Status)
                for r in range(4, len(df_c) + 4):
                    sheet.cell(row=r, column=10).number_format = fmt_contabil
                    sheet.cell(row=r, column=11).number_format = fmt_contabil
                    
                    c_dif = sheet.cell(row=r, column=12)
                    c_dif.number_format = fmt_contabil
                    c_dif.font = cor_vermelha if c_dif.value and c_dif.value < 0 else cor_verde
                    
                    c_status = sheet.cell(row=r, column=13)
                    c_status.font = cor_verde if c_status.value == "OK" else cor_vermelha

        st.success("✅ Relatório Colorido Gerado!")
        st.download_button("📥 Baixar Excel Colorido", data=output.getvalue(), file_name="conciliacao_cores.xlsx")
            
    except Exception as e:
        st.error(f"Erro: {e}")
