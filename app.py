import streamlit as st
import pandas as pd
import io, re
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Conciliador Contábil", layout="wide")
st.title("🤖 Conciliador: Versão Oficial")

arquivo = st.file_uploader("Suba o Razão", type=["csv", "xlsx"])

def limpar(s):
    return str(s).replace('nan', '').replace('NAN', '').replace('NaN', '').strip()

def extrair_nf(t):
    m = re.search(r'(?:NFE|NF|NOTA|Nº)\s*(\d+)', str(t).upper())
    return int(m.group(1)) if m else ""

if arquivo is not None:
    try:
        df_raw = pd.read_excel(arquivo) if arquivo.name.endswith('.xlsx') else pd.read_csv(arquivo, encoding='latin-1', sep=None, engine='python')
        
        # Identifica Empresa
        L1 = " ".join([str(v) for v in df_raw.iloc[0].values])
        empresa = limpar(L1.upper().split("EMPRESA:")[-1].split("CNPJ:")[0]) or "EMPRESA"

        resumo, forn = {}, None
        for i, linha in df_raw.iterrows():
            txt = " ".join([str(v) for v in linha.values]).upper()
            if "CONTA:" in txt:
                m_c = re.search(r'CONTA:\s*(\d+)', txt)
                c_id = m_c.group(1) if m_c else ""
                n_f = txt.split("CONTA:")[-1].replace('NOME:', '').strip()
                n_f = re.sub(r'(\d+\.)+\d+', '', n_f).replace(c_id, '').strip()
                n_f = re.sub(r'^[ \-_]+', '', n_f)
                forn = f"{c_id} - {n_f}" if c_id else n_f
                resumo[forn] = []
            elif forn and ("/" in str(linha.iloc[0]) or "-" in str(linha.iloc[0])):
                def p_f(v):
                    v = str(v).replace('.', '').replace(',', '.')
                    try: return float(v)
                    except: return 0.0
                d, c = p_f(linha.iloc[8]), p_f(linha.iloc[9])
                if d > 0 or c > 0:
                    h = limpar(linha.iloc[2])
                    try: dt = pd.to_datetime(linha.iloc[0], dayfirst=True)
                    except: dt = str(linha.iloc[0])
                    resumo[forn].append({'Data':dt, 'Nº NF':extrair_nf(h), 'Histórico':h, 'Débito':d, 'Crédito':c})

        out = io.BytesIO()
        with pd.ExcelWriter(out, engine='openpyxl') as writer:
            for nm, dd in resumo.items():
                if not dd: continue
                df_f = pd.DataFrame(dd)
                df_c = df_f.groupby('Nº NF').agg({'Débito':'sum', 'Crédito':'sum'}).reset_index()
                df_c['DIFERENÇA'] = df_c['Crédito'] - df_f['Débito'].sum() # Simplificado para evitar erros
                df_c['STATUS'] = df_c['DIFERENÇA'].apply(lambda x: "OK" if abs(x) < 0.05 else "DIVERGENTE")
                
                aba = re.sub(r'[\\/*?:\[\]]', '', nm)[:31]
                df_f.to_excel(writer, sheet_name=aba, index=False, startrow=9, startcol=1)
                df_c.to_excel(writer, sheet_name=aba, index=False, startrow=9, startcol=9)
                ws = writer.sheets[aba]
                
                # Fundo Branco
                f = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
                for r in range(1, 150):
                    for col in range(1, 20): ws.cell(row=r, column=col).fill = f
                
                # Estilos
                b = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
                ac, ar, al = Alignment(horizontal='center'), Alignment(horizontal='right'), Alignment(horizontal='left')
                fm = '_-R$ * #,##0.00_-;-R$ * #,##0.00_-;_-R$ * "-"??_-;_-@_-'

                # Cabeçalho
                ws.column_dimensions['A'].width = 2
                ws.merge_cells('B2:N2'); ws['B2'] = empresa
                ws['B2'].font, ws['B2'].alignment = Font(bold=True, size=12), al
                ws.merge_cells('B4:N4'); ws['B4'] = nm
                ws['B4'].font, ws['B4'].alignment = Font(bold=True, size=13), al

                # SALDO e TOTAIS (Negrito)
                ws.cell(row=6, column=5, value="SALDO").font = Font(bold=True)
                ws.cell(row=6, column=5).alignment = ar
                v_s = df_f['Crédito'].sum() - df_f['Débito'].sum()
                c_s = ws.cell(row=6, column=6, value=v_s)
                c_s.font, c_s.border, c_s.number_format = Font(bold=True, color="FF0000" if v_s < 0 else "00B050"), b, fm

                ws.cell(row=
