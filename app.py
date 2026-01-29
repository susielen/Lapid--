import streamlit as st
import pandas as pd
import io, re
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Conciliador Contábil Pro", layout="wide")
st.title("🤖 Conciliador: Cabeçalho Limpo e Profissional")

arquivo = st.file_uploader("Suba o Razão aqui", type=["csv", "xlsx"])

def extrair_nf(t):
    m = re.search(r'(?:NFE|NF|NOTA|Nº)\s*(\d+)', str(t).upper())
    if m:
        try: return int(m.group(1))
        except: return m.group(1)
    return ""

def limpar_string(s):
    res = str(s).replace('nan', '').replace('NAN', '').replace('NaN', '').strip()
    return res

def limpar_nome_correto(l):
    l = limpar_string(l).upper()
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
        
        # BUSCANDO NOME DA EMPRESA
        primeira_linha = " ".join([str(v) for v in df_raw.iloc[0].values])
        nome_empresa = limpar_string(primeira_linha.upper().split("EMPRESA:")[-1].split("CNPJ:")[0])
        if not nome_empresa:
            nome_empresa = "EMPRESA NÃO IDENTIFICADA"

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
                    try: dt_obj = pd.to_datetime(linha.iloc[0], dayfirst=True)
                    except: dt_obj = str(linha.iloc[0])
                    resumo[fornecedor].append({'Data': dt_obj, 'Nº NF': extrair_nf(h), 'Histórico': h, 'Débito': d, 'Crédito': c})

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            for nome_forn, lancs in resumo.items():
                if not lancs: continue
                df_f = pd.DataFrame(lancs)
                df_c = df_f.groupby('Nº NF').agg({'Débito': 'sum', 'Crédito': 'sum'}).reset_index()
                df_c['DIFERENÇA'] = df_c['Crédito'] - df_c['Débito']
                df_c['STATUS'] = df_c['DIFERENÇA'].apply(lambda x: "OK" if abs(x) < 0.05 else "DIVERGENTE")
                
                aba = re.sub(r'[\\/*?:\[\]]', '', nome_forn)[:31]
                df_f.to_excel(writer, sheet_name=aba, index=False, startrow=9, startcol=1)
                df_c.to_excel(writer, sheet_name=aba, index=False, startrow=9, startcol=9)
                ws = writer.sheets[aba]
                
                # Fundo Branco
                fundo_br = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
                for r in range(1, 201):
                    for c in range(1, 21):
                        ws.cell(row=r, column=c).fill = fundo_br

                b = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
                ac, ar, al = Alignment(horizontal='center'), Alignment(horizontal='right'), Alignment(horizontal='left')
                fmt_moeda = '_-R$ * #,##0.0
