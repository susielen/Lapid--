import streamlit as st
import pandas as pd
import io, re
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Conciliador", layout="wide")
st.title("🤖 Conciliador Pro")

file = st.file_uploader("Suba o Razão", type=["csv", "xlsx"])

def clean(s): return str(s).replace('nan','').replace('NAN','').strip()

if file:
    try:
        df = pd.read_excel(file) if file.name.endswith('.xlsx') else pd.read_csv(file, encoding='latin-1', sep=None, engine='python')
        emp = clean(" ".join([str(v) for v in df.iloc[0].values]).upper().split("EMPRESA:")[-1].split("CNPJ:")[0]) or "EMPRESA"
        
        res, forn = {}, None
        for i, r in df.iterrows():
            txt = " ".join([str(v) for v in r.values]).upper()
            if "CONTA:" in txt:
                m = re.search(r'CONTA:\s*(\d+)', txt)
                cod = m.group(1) if m else ""
                nm = txt.split("CONTA:")[-1].replace('NOME:','').strip()
                nm = re.sub(r'(\d+\.)+\d+', '', nm).replace(cod, '').strip()
                forn = f"{cod} - {re.sub(r'^[ \-_]+', '', nm)}"
                res[forn] = []
            elif forn and ("/" in str(r.iloc[0]) or "-" in str(r.iloc[0])):
                def p(v):
                    try: return float(str(v).replace('.','').replace(',','.'))
                    except: return 0.0
                res[forn].append({'Data':r.iloc[0], 'Hist':clean(r.iloc[2]), 'D':p(r.iloc[8]), 'C':p(r.iloc[9])})

        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine='openpyxl') as writer:
            for nm, dd in res.items():
                if not dd: continue
                df_f = pd.DataFrame(dd)
                aba = re.sub(r'[\\/*?:\[\]]', '', nm)[:31]
                df_f.to_excel(writer, sheet_name=aba, index=False, startrow=9, startcol=1)
                ws = writer.sheets[aba]
                
                # Fundo e Margens
                for r in range(1, 100):
                    for c in range(1, 15): ws.cell(row=r, column=c).fill = PatternFill("solid", fgColor="FFFFFF")
                ws.column_dimensions['A'].width = 2
                ws.row_dimensions[1].height = 10
                
                # Estilos
                b = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
                f_m = '_-R$ * #,##0.00_-;-R$ * #,##0.00_-;_-R$ * "-"??_-;_-@_-'
                
                # Cabeçalho
                ws.merge_cells('B2:H2'); ws['B2'] = emp
                ws['B2'].font = Font(bold=True, size=12)
                ws.merge_cells('B4:H4'); ws['B4'] = nm
                ws['B4'].font = Font(bold=True, size=13)
                
                # Saldo e Totais (NEGRITO)
                ws.cell(row=6, column=4, value="SALDO").font = Font(bold=True)
                v_s = df_f['C'].sum() - df_f['D'].sum()
                ws.cell(row=6, column=5, value=v_s).font = Font(bold=True, color="00B050" if v_s >= 0 else "FF0000")
                ws.cell(row=6, column=5).number_format = f_m
                
                ws.cell(row=8, column=3, value="TOTAIS").font = Font(bold=True)
                for ci, v, cr in [(4, df_f['D'].sum(), "FF0000"), (5, df_f['C'].sum(), "00B050")]:
                    c = ws.cell(row=8, column=ci, value=v)
                    c.font, c.border, c.number_format = Font(bold=True, color=cr), b, f_m

                # Ajuste Colunas
                for c in range(2, 7):
                    ws.column_dimensions[get_column_letter(c)].width = 40 if c==3 else 18
                    ws.cell(row=10, column=c).font = Font(bold=True)
        
        st.success("✅ Relatório gerado!")
        st.download_button("📥 Baixar Excel", buf.getvalue(), "conciliacao.xlsx")
    except Exception as e:
        st.error(f"Erro: {e}")
