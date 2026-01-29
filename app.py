import streamlit as st
import pandas as pd
import io
import re
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Conciliador Contábil Pro", layout="wide")
st.title("🤖 Conciliador: Layout Premium sem Erros")

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
                
                # --- MESCLAR NOME NO TOPO ---
                sheet.merge_cells('A1:G1')
                celula_titulo = sheet['A1']
                celula_titulo.value = forn
                celula_titulo.font = Font(bold=True, size=14)
                celula_titulo.alignment = Alignment(horizontal='center')

                # --- TOTAIS ---
                sheet.cell(row=3, column=4, value="TOTAIS").font = Font(bold=True)
                sheet.cell(row=3, column=6, value="SALDO").font = Font(bold=True)

                # VALORES TOPO
                for col, val, cor in [(4, df_f['Débito'].sum(), "FF0000"), (5, df_f['Crédito'].sum(), "00B050")]:
                    c = sheet.cell(row=4, column=col, value=val)
                    c.number_format = fmt_contabil
                    c.font = Font(bold=True, color=cor)
                
                saldo = df_f['Crédito'].sum() - df_f['Débito'].sum()
                v_saldo = sheet.cell(row=4, column=6, value=saldo)
                v_saldo.number_format = fmt_contabil
                v_saldo.font = Font(bold=True, color="FF0000" if saldo < 0 else "00B050")

                # CONCILIAÇÃO TOPO
                sheet.cell(row=4, column=12, value="Saldo").font = Font(bold=True)
                v_conc = sheet.cell(row=4, column=13, value=saldo)
                v_conc.number_format = fmt_contabil
                v_conc.font = Font(bold=True, color="FF0000" if saldo < 0 else "00B050")

                # --- REMOVER ALERTAS E AJUSTAR COLUNAS ---
                # Ignorar erros de números armazenados como texto
                sheet.add_ignore_error('A1:Z500', numberStoredAsText=True)

                for column in sheet.columns:
                    column_letter = get_column_letter(column[0].column)
                    if column_letter == 'A': # Coluna Data
                        sheet.column_dimensions[column_letter].width = 12
                    elif column_letter == 'C': # Coluna Histórico (Mais larga)
                        sheet.column_dimensions[column_letter].width = 40
                    else:
                        sheet.column_dimensions[column_letter].width = 18

                # Formatação das cores de status no corpo
                for r in range(7, len(df_c) + 7):
                    st_cell = sheet.cell(row=r, column=13)
                    st_cell.font = Font(color="00B050") if st_cell.value == "OK" else Font(color="FF0000")

        st.success("✅ Relatório Premium Gerado com Sucesso!")
        st.download_button("📥 Baixar Excel Final", data=output.getvalue(), file_name="conciliacao_automatica.xlsx")
            
    except Exception as e:
        st.error(f"Erro: {e}")
