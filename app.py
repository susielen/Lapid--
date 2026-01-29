import streamlit as st
import pandas as pd
import io, re
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Conciliador Contábil", layout="wide")
st.title("🤖 Conciliador: Nomes Limpos e Negritos")

arquivo = st.file_uploader("Suba o Razão (Excel ou CSV)", type=["csv", "xlsx"])

def limpar_nan(s):
    # Deixa o texto limpo, sem os erros de "nan"
    txt = str(s).replace('nan', '').replace('NAN', '').replace('NaN', '').strip()
    return txt

def extrair_nf(t):
    m = re.search(r'(?:NFE|NF|NOTA|Nº)\s*(\d+)', str(t).upper())
    return int(m.group(1)) if m else ""

def limpar_fornecedor_total(l):
    l = limpar_nan(l).upper()
    m_cod = re.search(r'CONTA:\s*(\d+)', l)
    cod = m_cod.group(1) if m_cod else ""
    nome = l.split("CONTA:")[-1].replace('NOME:', '').strip()
    # Tira os pontos e o código que se repete no nome
    nome = re.sub(r'(\d+\.)+\d+', '', nome).replace(cod, '').strip()
    nome = re.sub(r'^[ \-_]+', '', nome)
    return f"{cod} - {nome}" if cod else nome

if arquivo is not None:
    try:
        df_raw = pd.read_excel(arquivo) if arquivo.name.endswith('.xlsx') else pd.read_csv(arquivo, encoding='latin-1', sep=None, engine='python')
        
        # Pega o nome da empresa e limpa
        prim_linha = " ".join([str(v) for v in df_raw.iloc[0].values])
        nome_empresa = limpar_nan(prim_linha.upper().split("EMPRESA:")[-1].split("CNPJ:")[0])
        if not nome_empresa: nome_empresa = "EMPRESA NÃO IDENTIFICADA"

        resumo, forn_atual = {}, None
        for i, linha in df_raw.iterrows():
            txt = " ".join([str(v) for v in linha.values]).upper()
            if "CONTA:" in txt:
                forn_atual = limpar_fornecedor_total(txt)
                resumo[forn_atual] = []
            elif forn_atual and ("/" in str(linha.iloc[0]) or "-" in str(linha.iloc[0])):
                def conv(v):
                    v = str(v).replace('.', '').replace(',', '.')
                    try: return float(v)
                    except: return 0.0
                d, c = conv(linha.iloc[8]), conv(linha.iloc[9])
                if d > 0 or c > 0:
                    hist = limpar_nan(linha.iloc[2])
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
