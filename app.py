import streamlit as st
import pandas as pd
import io, re
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Conciliador Pro", layout="wide")
st.title("🤖 Conciliador: Versão Corrigida")

arquivo = st.file_uploader("Suba o Razão aqui", type=["csv", "xlsx"])

def extrair_nf(t):
    m = re.search(r'(?:NFE|NF|NOTA|Nº)\s*(\d+)', str(t).upper())
    return int(m.group(1)) if m else ""

def limpar_nome(l):
    l = str(l).replace('nan', '').upper()
    cod = re.search(r'CONTA:\s*(\d+)', l)
    c = cod.group(1) if cod else ""
    n = l.split("CONTA:")[-1].replace('NOME:', '').strip()
    n = re.sub(r'(\d+\.)+\d+', '', n).strip()
    return f"{c} - {n}" if c else n

if arquivo is not None:
    try:
        df_raw = pd.read_excel(arquivo) if arquivo.name.endswith('.xlsx') else pd.read_csv(arquivo, encoding='latin-1', sep=None, engine='python')
        resumo = {}
        fornecedor = None
        for i, linha in df_raw.iterrows():
            texto = " ".join([str(v) for v in linha.values]).upper()
            if "CONTA:" in texto:
                fornecedor = limpar_nome(texto)
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
                df_c['STATUS'] = df_c['DIFERENÇA'].apply(lambda x: "OK" if abs(x) < 0.05 else "ERRO")
                
                aba = re.sub(r'[\\/*?:\[\]]', '', nome)[:31]
                df_f.to_excel(writer, sheet_name=aba, index=False, startrow=6)
                df_c.to_excel(writer, sheet_name=aba, index=False, startrow=6, startcol=8)
                ws = writer.sheets[aba]
                
                b = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
                ac, ar = Alignment(horizontal='center'), Alignment(horizontal='right')
                
                # Topo
                ws.merge_cells('A1:M1')
                ws['A1'] = nome
                ws['A1'].font, ws['A1'].alignment = Font(bold=True, size=12), ac
                
                # Linha 3: Saldo
                ws.cell(row=3, column=4, value="SALDO").alignment = ar
                val_s = df_f['Crédito'].sum() - df_f['Débito'].sum()
                cs = ws.cell(row=3, column=5, value=val_s)
                cs.font, cs.border, cs.number_format = Font(bold=True), b, '#,##0.00'
                
                # Linha 5: Totais e Conciliação
                ws.cell(row=5, column=3, value="TOTAIS").alignment = ar
                for ci, v in [(4, df_f['Débito'].sum()), (5, df_f['Crédito'].sum())]:
                    cel = ws.cell(row=5, column=ci, value=v)
                    cel.border, cel.number_format = b, '#,##0.00'
                
                ws.merge_cells('I5:K5')
                ws['I5'] = "CONCILIAÇÃO"
                ws['I5'].font, ws['I5'].alignment = Font(bold=True), ac
                ws.cell(row=5, column=12, value="Saldo").font = Font(bold=True)
                cc = ws.cell(row=5, column=13, value=val_s)
                cc.border, cc.number_format = b, '#,##0.00'
                
                # Estilo Tabelas
                for r in range(7, 8 + max(len(df_f), len(df_c))):
                    for c in range(1, 14):
                        if ws.cell(row=r, column=c).value is not None:
                            if c != 6: ws.cell(row=r, column=c).border = b
                            if r == 7: ws.cell(row=r, column=c).fill = PatternFill("solid", fgColor="D3D3D3")
                
                for c in range(1, 14):
                    letra = get_column_letter(c)
                    ws.column_dimensions[letra].width = 4 if letra in ['F', 'G', 'H'] else (40 if letra == 'C' else 15)

        st.success("✅ Relatório gerado!")
        st.download_button("📥 Baixar Excel", output.getvalue(), "conciliacao.xlsx")
    except Exception as e:
        st.error(f"Erro no processamento: {e}")
