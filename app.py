import streamlit as st
import pandas as pd
import io
import re
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Conciliador Contábil Pro", layout="wide")
st.title("🤖 Conciliador: Layout e Erros Ajustados")

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
                    if pd.isna(v) or str(v).lower() == 'nan' or str(v).strip() == '': return 0.0
                    v = str(v).replace('.', '').replace(',', '.')
                    try: return float(v)
                    except: return 0.0
                
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
                sheet.sheet_view.showGridLines = False
                
                fmt_contabil = '_-R$ * #,##0.00_-;-R$ * #,##0.00_-;_-R$ * "-"??_-;_-@_-'
                preenchimento_cinza = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")
                borda_fina = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
                alinhar_centro = Alignment(horizontal='center')

                # --- 1. CABEÇALHO DO FORNECEDOR ---
                sheet.merge_cells('A1:M1')
                sheet['A1'] = forn
                sheet['A1'].font = Font(bold=True, size=14)
                sheet['A1'].alignment = alinhar_centro

                # --- 2. NOVA ORGANIZAÇÃO DO TOPO ---
                # "Saldo" na linha 4 coluna D, valor na coluna E linha 4
                sheet.cell(row=4, column=4, value="SALDO").font = Font(bold=True)
                saldo_val = df_f['Crédito'].sum() - df_f['Débito'].sum()
                c_saldo = sheet.cell(row=4, column=5, value=saldo_val)
                c_saldo.number_format = fmt_contabil
                c_saldo.font = Font(bold=True, color="FF0000" if saldo_val < 0 else "00B050")
                c_saldo.border = borda_fina

                # "Totais" na linha 5 coluna C, valores em D e E
                sheet.cell(row=5, column=3, value="TOTAIS").font = Font(bold=True)
                v_deb_t = sheet.cell(row=5, column=4, value=df_f['Débito'].sum())
                v_deb_t.number_format = fmt_contabil
                v_deb_t.font = Font(bold=True, color="FF0000")
                v_deb_t.border = borda_fina

                v_cre_t = sheet.cell(row=5, column=5, value=df_f['Crédito'].sum())
                v_cre_t.number_format = fmt_contabil
                v_cre_t.font = Font(bold=True, color="00B050")
                v_cre_t.border = borda_fina

                # Escrever "CONCILIAÇÃO" na linha 4 de I até K
                sheet.merge_cells('I4:K4')
                sheet['I4'] = "CONCILIAÇÃO"
                sheet['I4'].font = Font(bold=True)
                sheet['I4'].alignment = alinhar_centro

                # Saldo da Conciliação (Coluna L/M)
                sheet.cell(row=4, column=12, value="Saldo").font = Font(bold=True)
                v_conc_v = sheet.cell(row=4, column=13, value=saldo_val)
                v_conc_v.number_format = fmt_contabil
                v_conc_v.font = Font(bold=True, color="FF0000" if saldo_val < 0 else "00B050")
                v_conc_v.border = borda_fina

                # --- 3. IGNORAR ERROS ---
                # Adiciona as colunas de Data (A) e Nº NF (B e I) para o Excel não reclamar
                ultima_linha = max(len(df_f), len(df_c)) + 10
                r_range = f"A7:B{ultima_linha} I7:I{ultima_linha}"
                sheet.add_ignored_error(r_range, numberStoredAsText=True)

                # --- 4. CABEÇALHOS DA TABELA (LINHA 6) ---
                for col_idx in range(1, 14):
                    celula = sheet.cell(row=6, column=col_idx)
                    if celula.value:
                        celula.fill = preenchimento_cinza
                        celula.font = Font(bold=True)
                        celula.alignment = alinhar_centro
                        if col_idx != 6: celula.border = borda_fina

                # --- 5. CORPO COM BORDAS E CENTRALIZAÇÃO ---
                for r in range(7, len(df_f) + 7):
                    for c_idx in range(1, 7):
                        cel = sheet.cell(row=r, column=c_idx)
                        if c_idx < 6: cel.border = borda_fina
                        if c_idx in [1, 2]: cel.alignment = alinhar_centro
                        if c_idx in [5, 6]: cel.number_format = fmt_contabil
                
                for r in range(7, len(df_c) + 7):
                    for c_idx in range(9, 14):
                        cel = sheet.cell(row=r, column=c_idx)
                        cel.border = borda_fina
                        if c_idx == 9: cel.alignment = alinhar_centro
                        if c_idx in [10, 11, 12]: cel.number_format = fmt_contabil
                    st_cell = sheet.cell(row=r, column=13)
                    st_cell.alignment = alinhar_centro
                    st_cell.font = Font(color="00B050") if st_cell.value == "OK" else Font(color="FF0000")

                # --- 6. LARGURAS ---
                for column in sheet.columns:
                    col_letter = get_column_letter(column[0].column)
                    if col_letter == 'A': sheet.column_dimensions[col_letter].width = 12
                    elif col_letter in ['G', 'H']: sheet.column_dimensions[col_letter].width = 4
                    elif col_letter == 'C': sheet.column_dimensions[col_letter].width = 45
                    else: sheet.column_dimensions[col_letter].width = 18

        st.success("✅ Relatório com novo layout gerado!")
        st.download_button("📥 Baixar Excel Atualizado", data=output.getvalue(), file_name="conciliacao_layout_novo.xlsx")
            
    except Exception as e:
        st.error(f"Erro: {e}")
