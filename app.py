import streamlit as st
import pandas as pd
import io, re
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Conciliador Contábil", layout="wide")
st.title("🤖 Conciliador Pro: Layout Final")

arquivo = st.file_uploader("Suba o Razão (Excel ou CSV)", type=["csv", "xlsx"])

def limpar_texto(s):
    if pd.isna(s) or str(s).lower() == 'nan': return ""
    return str(s).strip()

def extrair_nf(t):
    m = re.search(r'(?:NFE|NF|NOTA|Nº)\s*(\d+)', str(t).upper())
    return int(m.group(1)) if m else ""

def limpar_fornecedor(l):
    l = limpar_texto(l).upper()
    m_cod = re.search(r'CONTA:\s*(\d+)', l)
    cod = m_cod.group(1) if m_cod else ""
    nome = l.split("CONTA:")[-1].replace('NOME:', '').strip()
    nome = re.sub(r'(\d+\.)+\d+', '', nome).replace(cod, '').strip()
    nome = re.sub(r'^[ \-_]+', '', nome)
    return f"{cod} - {nome}" if cod else nome

if arquivo is not None:
    try:
        df_raw = pd.read_excel(arquivo) if arquivo.name.endswith('.xlsx') else pd.read_csv(arquivo, encoding='latin-1', sep=None, engine='python')
        
        # Identifica Nome da Empresa (limpa NAN)
        prim_linha = " ".join([str(v) for v in df_raw.iloc[0].values])
        nome_empresa = limpar_texto(prim_linha.upper().split("EMPRESA:")[-1].split("CNPJ:")[0])
        if not nome_empresa: nome_empresa = "EMPRESA NÃO IDENTIFICADA"

        resumo, forn_atual = {}, None
        for i, linha in df_raw.iterrows():
            txt = " ".join([str(v) for v in linha.values]).upper()
            if "CONTA:" in txt:
                forn_atual = limpar_fornecedor(txt)
                resumo[forn_atual] = []
            elif forn_atual and ("/" in str(linha.iloc[0]) or "-" in str(linha.iloc[0])):
                def conv(v):
                    v = str(v).replace('.', '').replace(',', '.')
                    try: return float(v)
                    except: return 0.0
                d, c = conv(linha.iloc[8]), conv(linha.iloc[9])
                if d > 0 or c > 0:
                    hist = limpar_texto(linha.iloc[2])
                    try: dt = pd.to_datetime(linha.iloc[0], dayfirst=True)
                    except: dt = str(linha.iloc[0])
                    resumo[forn_atual].append({'Data': dt, 'Nº NF': extrair_nf(hist), 'Histórico': hist, 'Débito': d, 'Crédito': c})

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            for nome_f, dados in resumo.items():
                if not dados: continue
                df_f = pd.DataFrame(dados)
                df_c = df_f.groupby('Nº NF').agg({'Débito': 'sum', 'Crédito': 'sum'}).reset_index()
                df_c['DIFERENÇA'] = df_c['Crédito'] - df_c['Débito']
                df_c['STATUS'] = df_c['DIFERENÇA'].apply(lambda x: "OK" if abs(x) < 0.05 else "DIVERGENTE")
                
                aba = re.sub(r'[\\/*?:\[\]]', '', nome_f)[:31]
                df_f.to_excel(writer, sheet_name=aba, index=False, startrow=9, startcol=1)
                df_c.to_excel(writer, sheet_name=aba, index=False, startrow=9, startcol=9)
                ws = writer.sheets[aba]
                
                # Fundo Branco e Margens
                br = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
                for r in range(1, 150):
                    for c in range(1, 20): ws.cell(row=r, column=c).fill = br
                
                ws.column_dimensions['A'].width = 2
                ws.row_dimensions[1].height = 10
                
                # Estilos
                b = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
                ac, ar, al = Alignment(horizontal='center'), Alignment(horizontal='right'), Alignment(horizontal='left')
                f_m = '_-R$ * #,##0.00_-;-R$ * #,##0.00_-;_-R$ * "-"??_-;_-@_-'

                # Cabeçalho
                ws.merge_cells('B2:N2'); ws['B2'] = nome_empresa
                ws['B2'].font, ws['B2'].alignment = Font(bold=True, size=12), al
                
                ws.merge_cells('B4:N4'); ws['B4'] = nome_f
                ws['B4'].font, ws['B4'].alignment = Font(bold=True, size=13), al

                # Saldo e Totais
                val_s = df_f['Crédito'].sum() - df_f['Débito'].sum()
                ws.cell(row=6, column=5, value="SALDO").alignment = ar
                cs = ws.cell(row=6, column=6, value=val_s)
                cs.font, cs.border, cs.number_format = Font(bold=True, color="FF0000" if val_s < 0 else "00B050"), b, f_m

                ws.cell(row=8, column=4, value="TOTAIS").alignment = ar
                for ci, v, cor in [(5, df_f['Débito'].sum(), "FF0000"), (6, df_f['Crédito'].sum(), "00B050")]:
                    cel = ws.cell(row=8, column=ci, value=v)
                    cel.font, cel.border, cel.number_format = Font(bold=True, color=cor), b, f_m

                # Conciliação
                ws.merge_cells('J8:L8'); ws['J8'] = "CONCILIAÇÃO POR NOTA"
                ws['J8'].font, ws['J8'].alignment = Font(bold=True), ac
                ws.cell(row=8, column=13, value="Saldo").font = Font(bold=True)
                cc = ws.cell(row=8, column=14, value=val_s)
                cc.font, cc.border, cc.number_format = Font(bold=True, color="FF0000" if val_s < 0 else "00B050"), b, f_m

                # Tabela
                cinza = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")
                for c in range(2, 15):
                    cel = ws.cell(row=10, column=c)
                    if cel.value:
                        cel.font, cel.alignment, cel.fill = Font(bold=True), ac, cinza
                        if c != 7: cel.border = b

                for r in range(11, 11 + len(df_f)):
                    for c in range(2, 8):
                        cel = ws.cell(row=r, column=c)
                        if c < 7: cel.border = b
                        if c == 2: cel.number_format = 'dd/mm/yyyy'
                        if c in [2, 3]: cel.alignment = ac
                        if c == 6: cel.font = Font(color="FF0000")
                        if c == 7: cel.font = Font(color="00B050")
                        if c in [6, 7]: cel.number_format = f_m

                for r in range(11, 11 + len(df_c)):
                    for c in range(10, 15):
                        cel = ws.cell(row=r, column=c)
                        cel.border = b
                        if c == 10: cel.alignment = ac
                        if c in [11, 12, 13]: cel.number_format = f_m
                    ws.cell(row=r, column=14).font = Font(color="00B050" if ws.cell(row=r, column=14).value == "OK" else "FF0000")

                for c in range(2, 15):
                    L = get_column_letter(c)
                    ws.column_dimensions[L].width = 4 if L in ['G','H','I'] else (45 if L=='D' else 18)

        st.success("✅ Relatório gerado com sucesso!")
        st.download_button("📥 Baixar Excel", output.getvalue(), "conciliacao_final.xlsx")
    except Exception as e:
        st.error(f"Erro no processamento: {e}")
