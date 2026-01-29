import streamlit as st
import pandas as pd
import io
import re
from openpyxl.styles import Font

st.set_page_config(page_title="Conciliador Contábil Pro", layout="wide")
st.title("🤖 Conciliador: Limpeza Total e Totais Alinhados")

arquivo = st.file_uploader("Suba o Razão do Domínio aqui", type=["csv", "xlsx"])

def extrair_nfe(texto):
    match = re.search(r'(?:NFE|NF|NOTA|Nº)\s*(\d+)', str(texto).upper())
    return match.group(1) if match else "(vazio)"

def limpar_nome_simples(linha_txt):
    # Remove o termo 'nan' que aparece em campos vazios
    linha_txt = str(linha_txt).replace('nan', '').replace('NAN', '')
    match_cod = re.search(r'CONTA:\s*(\d+)', linha_txt)
    codigo = match_cod.group(1) if match_cod else ""
    
    nome = linha_txt.split("CONTA:")[-1]
    nome = re.sub(r'(\d+\.)+\d+', '', nome) 
    nome = nome.replace(codigo, '').replace('NOME:', '').strip()
    # Remove traços ou caracteres estranhos que sobraram
    nome = re.sub(r'^[ \-_]+', '', nome)
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
            # Limpa 'nan' de toda a linha antes de processar
            valores_limpos = [str(v).replace('nan', '').strip() for v in linha.values]
            linha_txt = " ".join(valores_limpos).upper()
            
            if "CONTA:" in linha_txt:
                fornecedor_atual = limpar_nome_simples(linha_txt)
                dict_fornecedores[fornecedor_atual] = []
                continue
            
            data_orig = str(linha.iloc[0])
            if "/" in data_orig or (len(data_orig) >= 8 and "-" in data_orig):
                def limpar_num(v):
                    if pd.isna(v) or str(v).lower() == 'nan': return 0
                    v = str(v).replace('.', '').replace(',', '.')
                    return pd.to_numeric(v, errors='coerce') or 0
                
                deb = limpar_num(linha.iloc[8])
                cre = limpar_num(linha.iloc[9])

                if (deb > 0 or cre > 0) and fornecedor_atual:
                    hist = str(linha.iloc[2]).replace('nan', '')
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
                df_c['DIFERENÇA'] = df_c['Crédito'] - df_f['Débito'].sum() # Regra do seu saldo
                df_c['DIFERENÇA'] = df_c['Crédito'] - df_c['Débito']
                df_c['STATUS'] = df_c['DIFERENÇA'].apply(lambda x: "OK" if abs(x) < 0.05 else "DIVERGENTE")
                
                nome_aba = re.sub(r'[\\/*?:\[\]]', '', forn)[:31]
                
                # Inicia as tabelas na linha 6 para dar espaço ao topo
                df_f.to_excel(writer, sheet_name=nome_aba, index=False, startrow=5)
                df_c.to_excel(writer, sheet_name=nome_aba, index=False, startrow=5, startcol=8)
                
                sheet = writer.sheets[nome_aba]
                fmt_contabil = '_-R$ * #,##0.00_-;-R$ * #,##0.00_-;_-R$ * "-"??_-;_-@_-'
                negrito = Font(bold=True)

                # --- LINHA 1: NOME DO FORNECEDOR ---
                sheet.cell(row=1, column=1, value=forn).font = negrito

                # --- LINHA 3: TÍTULO TOTAIS ---
                sheet.cell(row=3, column=1, value="TOTAIS").font = negrito

                # --- LINHA 4: APENAS OS VALORES ---
                # Valor Débito em cima da coluna E (5)
                c_deb_topo = sheet.cell(row=4, column=5, value=df_f['Débito'].sum())
                c_deb_topo.number_format = fmt_contabil
                c_deb_topo.font = negrito

                # Valor Crédito em cima da coluna F (6)
                c_cre_topo = sheet.cell(row=4, column=6, value=df_f['Crédito'].sum())
                c_cre_topo.number_format = fmt_contabil
                c_cre_topo.font = negrito

                # Valor Diferença em cima da coluna L (12)
                saldo = df_f['Crédito'].sum() - df_f['Débito'].sum()
                c_dif_topo = sheet.cell(row=4, column=12, value=saldo)
                c_dif_topo.number_format = fmt_contabil
                c_dif_topo.font = Font(bold=True, color="FF0000" if saldo < 0 else "00B050")

                # Formatação Contábil do corpo
                for r in range(7, len(df_f) + 7):
                    sheet.cell(row=r, column=5).number_format = fmt_contabil
                    sheet.cell(row=r, column=6).number_format = fmt_contabil

                for r in range(7, len(df_c) + 7):
                    sheet.cell(row=r, column=10).number_format = fmt_contabil
                    sheet.cell(row=r, column=11).number_format = fmt_contabil
                    sheet.cell(row=r, column=12).number_format = fmt_contabil

        st.success("✅ Relatório Processado!")
        st.download_button("📥 Baixar Planilha Limpa", data=output.getvalue(), file_name="conciliacao_limpa.xlsx")
            
    except Exception as e:
        st.error(f"Erro: {e}")
