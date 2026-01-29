import streamlit as st
import pandas as pd
import io
import re
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Conciliador Contábil Pro", layout="wide")
st.title("🤖 Conciliador: Formato Contábil e Ajuste Automático")

arquivo = st.file_uploader("Suba o Razão do Domínio aqui", type=["csv", "xlsx"])

def extrair_nfe(texto):
    match = re.search(r'(?:NFE|NF|NOTA|Nº)\s*(\d+)', str(texto).upper())
    return match.group(1) if match else "(vazio)"

def limpar_nome_simples(linha_txt):
    linha_txt = str(linha_txt).replace('nan', '').replace('NAN', '').replace('NaN', '')
    match_cod = re.search(r'CONTA:\s*(\d+)', linha_txt)
    codigo = match_cod.group(1) if match_cod else ""
    nome = linha_txt.split("CONTA:")[-1]
    nome = re.sub(r'(\d+\.)+\d+', '', nome) 
    nome = nome.replace(codigo, '').replace('NOME:', '').strip()
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
            valores_limpos = [str(v).replace('nan', '').strip() for v in linha.values]
            linha_txt = " ".join(valores_limpos).upper()
            
            if "CONTA:" in linha_txt:
                fornecedor_atual = limpar_nome_simples(linha_txt)
                dict_fornecedores[fornecedor_atual] = []
                continue
            
            data_orig = str(linha.iloc[0])
            if "/" in data_orig or (len(data_orig) >= 8 and "-" in data_orig):
                try:
                    data_dt = pd.to_datetime(data_orig)
                    data_formatada = data_dt.strftime('%d/%m/%y')
                except:
                    data_formatada = data_orig

                def limpar_num(v):
                    if pd.isna(v) or str(v).lower() == 'nan': return 0
                    v = str(v).replace('.', '').replace(',', '.')
                    return pd.to_numeric(v, errors='coerce') or 0
                
                deb = limpar_num(linha.iloc[8])
                cre = limpar_num(linha.iloc[9])

                if (deb > 0 or cre > 0) and fornecedor_atual:
                    hist = str(linha.iloc[2]).replace('nan', '')
                    dict_fornecedores[fornecedor_atual].append({
                        'Data': data_formatada, 'Nº NF': extrair_nfe(hist),
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
                df_f.to_excel(writer, sheet_name=nome_aba, index=False, startrow=5)
                df_c.to_excel(writer, sheet_name=nome_aba, index=False, startrow=5, startcol=8)
                
                sheet = writer.sheets[nome_aba]
                fmt_contabil = '_-R$ * #,##0.00_-;-R$ * #,##0.00_-;_-R$ * "-"??_-;_-@_-'
                negrito = Font(bold=True)
                cor_vermelha = Font(bold=True, color="FF0000")
                cor_verde = Font(bold=True, color="00B050")

                # --- TOPO ---
                sheet.cell(row=1, column=1, value=forn).font = negrito
                sheet.cell(row=3, column=4, value="TOTAIS").font = negrito
                sheet.cell(row=3, column=6, value="SALDO").font = negrito

                # VALORES TOPO
                v_deb = sheet.cell(row=4, column=4, value=df_f['Débito'].sum())
                v_deb.number_format = fmt_contabil
                v_deb.font = cor_vermelha

                v_cre = sheet.cell(row=4, column=5, value=df_f['Crédito'].sum())
                v_cre.number_format = fmt_contabil
                v_cre.font = cor_verde
                
                saldo = df_f['Crédito'].sum() - df_f['Débito'].sum()
                v_saldo = sheet.cell(row=4, column=6, value=saldo)
                v_saldo.number_format = fmt_contabil
                v_saldo.font = cor_vermelha if saldo < 0 else cor_verde

                # CONCILIAÇÃO TOPO
                sheet.cell(row=4, column=12, value="Saldo").font = negrito
                v_conc = sheet.cell(row=4, column=13, value=saldo)
                v_conc.number_format = fmt_contabil
                v_conc.font = cor_vermelha if saldo < 0 else cor_verde

                # FORMATAR COLUNAS E VALORES DO CORPO
                for r in range(7, len(df_f) + 7):
                    sheet.cell(row=r, column=5).number_format = fmt_contabil
                    sheet.cell(row=r, column=6).number_format = fmt_contabil
                
                for r in range(7, len(df_c) + 7):
                    for col in [10, 11, 12]:
                        sheet.cell(row=r, column=col).number_format = fmt_contabil
                    st_cell = sheet.cell(row=r, column=13)
                    st_cell.font = Font(color="00B050") if st_cell.value == "OK" else Font(color="FF0000")

                # AJUSTE AUTOMÁTICO DE LARGURA
                for column in sheet.columns:
                    max_length = 0
                    column_letter = get_column_letter(column[0].column)
                    for cell in column:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except: pass
                    sheet.column_dimensions[column_letter].width = max_length + 2

        st.success("✅ Relatório Contábil Profissional Gerado!")
        st.download_button("📥 Baixar Excel Final", data=output.getvalue(), file_name="conciliacao_final_perfeita.xlsx")
            
    except Exception as e:
        st.error(f"Erro: {e}")
