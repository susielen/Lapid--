import streamlit as st
import pandas as pd
import io, re
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Conciliador Contábil Pro", layout="wide")
st.title("🤖 Conciliador: Tudo Limpo e Colorido")

arquivo = st.file_uploader("Suba o Razão aqui", type=["csv", "xlsx"])

def extrair_nf(t):
    m = re.search(r'(?:NFE|NF|NOTA|Nº)\s*(\d+)', str(t).upper())
    if m:
        try: return int(m.group(1))
        except: return m.group(1)
    return ""

def limpar_nome_correto(l):
    l = str(l).replace('nan', '').replace('NAN', '').upper()
    match_cod = re.search(r'CONTA:\s*(\d+)', l)
    codigo = match_cod.group(1) if match_cod else ""
    nome = l.split("CONTA:")[-1].replace('NOME:', '').strip()
    nome = re.sub(r'(\d+\.)+\d+', '', nome) 
    nome = nome.replace(codigo, '').strip()
    nome = re.sub(r'^[ \-_]+', '', nome)
    return f"{codigo} - {nome}" if codigo else nome

if arquivo is not None:
    try:
        df_raw = pd.read_excel(arquivo) if arquivo.name.endswith('.xlsx') else pd.read_csv(arquivo, encoding='latin-1', sep=None, engine='python')
        resumo = {}
        fornecedor = None
        
        for i, linha in df_raw.iterrows():
            texto = " ".join([str(v) for v in linha.values]).upper()
            if "CONTA:" in texto:
                fornecedor = limpar_nome_correto(texto)
                resumo[fornecedor] = []
            elif fornecedor and ("/" in str(linha.iloc[0]) or "-" in str(linha.iloc[0])):
                def n(v):
                    v = str(v).replace('.', '').replace(',', '.')
                    try: return float(v)
                    except: return 0.0
                d, c = n(linha.iloc[8]), n(linha.iloc[9])
                if d > 0 or c > 0:
                    h = str(linha.iloc[2]).replace('nan', '')
                    resumo[fornecedor].append({'Data': str(linha.iloc[0]), 'Nº NF': extrair_nf(h), 'Histórico': h, 'Débito': d, 'Crédito': c})

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            for nome, lancs in resumo.items():
                if not lancs: continue
                df_f = pd.DataFrame(lancs)
                df_c = df_f.groupby('Nº NF').agg({'Débito': 'sum', 'Crédito': 'sum'}).reset_index()
                df_c['DIFERENÇA'] = df_c['Crédito'] - df_c['Débito']
                df_c['STATUS'] = df_c['DIFERENÇA'].apply(lambda x: "OK" if abs(x) < 0.05 else "DIVERGENTE")
                
                aba = re.sub(r'[\\/*?:\[\]]', '', nome)[:31]
                df_f.to_excel(writer, sheet_name=aba, index=False, startrow=6)
                df_c.to_excel(writer, sheet_name=aba, index=False, startrow=6, startcol=8)
                ws = writer.sheets[aba]
                
                # --- TRUQUE PARA SUMIR AS GRADES (FUNDO BRANCO) ---
                fundo_branco = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
                for r in range(1, 200): # Pinta as primeiras 200 linhas
                    for c in range(1, 20): # Pinta as primeiras 20 colunas
                        ws.cell(row=r, column=c).fill = fundo_branco

                # Estilos
                b = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
                ac, ar = Alignment(horizontal='center'), Alignment(horizontal='right')
                fmt = '_-R$ * #,##0.00_-;-R$ * #,##0.00_-;_-R$ * "-"??_-;_-@_-'
                cor_vermelho = "FF0000"
                cor_verde = "00B050"

                # 1. TÍTULO
                ws.merge_cells('A1:M1')
                ws['A1'] = nome
                ws['A1'].font, ws['A1'].alignment = Font(bold=True, size=14), ac

                # 2. LINHA 3: SALDO
                ws.cell(row=3, column=4, value="SALDO").alignment = ar
                val_s = df_f['Crédito'].sum() - df_f['Débito'].sum()
                cs = ws.cell(row=3, column=5, value=val_s)
                cs.font = Font(bold=True, color=cor_vermelho if val_s < 0 else cor_verde)
                cs.border, cs.number_format = b, fmt

                # 3. LINHA 5: TOTAIS E CONCILIAÇÃO
                ws.cell(row=5, column=3, value="TOTAIS").alignment = ar
                # Débito Totais
                td = ws.cell(row=5, column=4, value=df_f['Débito'].sum())
                td.font, td.border, td.number_format = Font(bold=True, color=cor_vermelho), b, fmt
                # Crédito Totais
                tc = ws.cell(row=5, column=5, value=df_f['Crédito'].sum())
                tc.font, tc.border, tc.number_format = Font(bold=True, color=cor_verde), b, fmt

                ws.merge_cells('I5:K5')
                ws['I5'] = "CONCILIAÇÃO"
                ws['I5'].font, ws['I5'].alignment = Font(bold=True), ac
                ws.cell(row=5, column=12, value="Saldo").font = Font(bold=True)
                cc = ws.cell(row=5, column=13, value=val_s)
                cc.font, cc.border, cc.number_format = Font(bold=True, color=cor_vermelho if val_s < 0 else cor_verde), b, fmt

                # 4. CABEÇALHOS (LINHA 7)
                cinza = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")
                for c in range(1, 14):
                    cel = ws.cell(row=7, column=c)
                    if cel.value:
                        cel.font, cel.alignment, cel.fill = Font(bold=True), ac, cinza
                        if c != 6: cel.border = b

                # 5. CORPO
                for r in range(8, 8 + len(df_f)):
                    for c in range(1, 7):
                        cel = ws.cell(row=r, column=c)
                        if c < 6: cel.border = b
                        if c in [1, 2]: cel.alignment = ac
                        if c == 5: cel.font = Font(color=cor_vermelho)
                        if c == 6: cel.font = Font(color=cor_verde)
                        if c in [5, 6]: cel.number_format = fmt

                for r in range(8, 8 + len(df_c)):
                    for c in range(9, 14):
                        cel = ws.cell(row=r, column=c)
                        cel.border = b
                        if c == 9: cel.alignment = ac
                        if c in [10, 11, 12]: cel.number_format = fmt
                    st_c = ws.cell(row=r, column=13)
                    st_c.font = Font(color=cor_verde) if st_c.value == "OK" else Font(color=cor_vermelho)

                # 6. LARGURAS
                for c in range(1, 14):
                    letra = get_column_letter(c)
                    if letra in ['F', 'G', 'H']: ws.column_dimensions[letra].width = 4
                    elif letra == 'C': ws.column_dimensions[letra].width = 45
                    else: ws.column_dimensions[letra].width = 18

        st.success("✅ Tudo pronto! Sem grades, com cores e nomes limpos.")
        st.download_button("📥 Baixar Excel", output.getvalue(), "conciliacao_limpa.xlsx")
    except Exception as e:
        st.error(f"Erro: {e}")
