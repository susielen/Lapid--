import streamlit as st
import pandas as pd
import io
import re
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Conciliador Contábil Pro", layout="wide")
st.title("🤖 Conciliador: Versão Final Ajustada")

arquivo = st.file_uploader("Suba o Razão do Domínio aqui", type=["csv", "xlsx"])

def extrair_nfe(texto):
    match = re.search(r'(?:NFE|NF|NOTA|Nº)\s*(\d+)', str(texto).upper())
    return match.group(1) if match else "(vazio)"

def limpar_nome_contabil(linha_txt):
    # Remove qualquer variação de NAN que venha do arquivo
    nome_limpo = str(linha_txt).replace('NAN', '').replace('nan', '').replace('NaN', '')
    
    # Busca o código da conta
    match_cod = re.search(r'CONTA:\s*(\d+)', nome_limpo)
    codigo = match_cod.group(1) if match_cod else ""
    
    # Pega o nome após "NOME:" ou "CONTA:"
    if "NOME:" in nome_limpo:
        nome = nome_limpo.split("NOME:")[-1]
    else:
        nome = nome_limpo.split("CONTA:")[-1]
    
    # Limpeza final de sujeira e números de classificação
    nome = re.sub(r'(\d+\.)+\d+', '', nome) 
    nome = nome.replace(codigo, '').strip()
    nome = re.sub(r'^[ \-_]+', '', nome) # Remove traços no início
    
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
            valores_linha = [str(v).upper() for v in linha.values]
            linha_completa = " ".join(valores_linha)
            
            if "CONTA:" in linha_completa:
                fornecedor_atual = limpar_nome_contabil(linha_completa)
                dict_fornecedores[fornecedor_atual] = []
                continue
            
            data_val = str(linha.iloc[0])
            if "/" in data_val or (len(data_val) >= 8 and "-" in data_val):
                try:
                    data_dt = pd.to_datetime(data_val)
                    data_short = data_dt.strftime('%d/%m/%y')
                except:
                    data_short = data_val

                def conv_num(v):
                    if pd.isna(v) or str(v).lower() == 'nan': return 0
                    return pd.to_numeric(str(v).replace('.', '').replace(',', '.'), errors='coerce') or 0
                
                deb = conv_num(linha.iloc[8])
                cre = conv_num(linha.iloc[9])

                if (deb > 0 or cre > 0) and fornecedor_atual:
                    hist = str(linha.iloc[2]).replace('nan', '').replace('NAN', '')
                    dict_fornecedores[fornecedor_atual].append({
                        'Data': data_short, 'Nº NF': extrair_nfe(hist),
                        'Histórico': hist, 'Débito': deb, 'Crédito': cre
                    })

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            for forn, dados in dict_fornecedores.items():
                if not dados: continue
                
                df_f = pd.DataFrame(dados)
                df_c = df_f.groupby('Nº NF').agg({'Débito': 'sum', 'Crédito': 'sum'}).reset_index()
                df_c['DIFERENÇA'] = df_c['Crédito'] - df_c['Débito']
                df_c['STATUS'] = df_c['DIFERENÇA'].apply(lambda x: "OK" if abs(x) < 0.05 else "DIVERGENTE")
                
                nome_aba = re.sub(r'[\\/*?:\[\]]', '', forn)[:31]
                df_f.to_excel(writer, sheet_name=nome_aba, index=False, startrow=5)
                df_c.to_excel(writer, sheet_name=nome_aba, index=False, startrow=5, startcol=8)
                
                sheet = writer.sheets[nome_aba]
                fmt_money = '_-R$ * #,##0.00_-;-R$ * #,##0.00_-;_-R$ * "-"??_-;_-@_-'
                fill_grey = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")
                f_bold = Font(bold=True)

                # 1. Nome Mesclado e Negrito
                sheet.merge_cells('A1:G1')
                top_name = sheet['A1']
                top_name.value = forn
                top_name.font = Font(bold=True, size=14)
                top_name.alignment = Alignment(horizontal='center')

                # 2. Cabeçalhos de Totais (Linha 3)
                sheet.cell(row=3, column=4, value="TOTAIS").font = f_bold
                sheet.cell(row=3, column=6, value="SALDO").font = f_bold

                # 3. Valores Coloridos (Linha 4)
                v_deb = sheet.cell(row=4, column=4, value=df_f['Débito'].sum())
                v_deb.number_format = fmt_money
                v_deb.font = Font(bold=True, color="FF0000")

                v_cre = sheet.cell(row=4, column=5, value=df_f['Crédito'].sum())
                v_cre.number_format = fmt_money
                v_cre.font = Font(bold=True, color="00B050")
                
                saldo = df_f['Crédito'].sum() - df_f['Débito'].sum()
                v_sal = sheet.cell(row=4, column=6, value=saldo)
                v_sal.number_format = fmt_money
                v_sal.font = Font(bold=True, color="FF0000" if saldo < 0 else "00B050")

                # Saldo Conciliação
                sheet.cell(row=4, column=12, value="Saldo").font = f_bold
                v_c_val = sheet.cell(row=4, column=13, value=saldo)
                v_c_val.number_format = fmt_money
                v_c_val.font = Font(bold=True, color="FF0000" if saldo < 0 else "00B050")

                # 4. Títulos das Colunas em Cinza (Linha 6)
                for c_idx in range(1, 14):
                    cell = sheet.cell(row=6, column=c_idx)
                    if cell.value:
                        cell.fill = fill_grey
                        cell.font = f_bold

                # 5. Ajuste de Colunas e Remover Alertas
                sheet.add_ignore_error('A1:Z500', numberStoredAsText=True)
                for col in sheet.columns:
                    col_let = get_column_letter(col[0].column)
                    if col_let == 'A': sheet.column_dimensions[col_let].width = 12
                    elif col_let == 'C': sheet.column_dimensions[col_let].width = 40
                    else: sheet.column_dimensions[col_let].width = 18

                # 6. Formatação do corpo
                for r in range(7, len(df_f) + 7):
                    sheet.cell(row=r, column=5).number_format = fmt_money
                    sheet.cell(row=r, column=6).number_format = fmt_money
                for r in range(7, len(df_c) + 7):
                    for c_i in [10, 11, 12]:
                        sheet.cell(row=r, column=c_i).number_format = fmt_money
                    st_c = sheet.cell(row=r, column=13)
                    st_c.font = Font(color="00B050") if st_c.value == "OK" else Font(color="FF0000")

        st.success("✅ Tudo corrigido! Baixe o arquivo abaixo.")
        st.download_button("📥 Baixar Excel Corrigido", data=output.getvalue(), file_name="conciliacao_final.xlsx")
            
    except Exception as e:
        st.error(f"Erro inesperado: {e}")
