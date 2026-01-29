import streamlit as st
import pandas as pd
import io, re
from openpyxl.styles import Font, Alignment, PatternFill

st.set_page_config(page_title="Conciliador Datas Corretas", layout="wide")
st.title("📅 Conciliador (Agora com Datas Limpas)")

arquivo = st.file_uploader("Suba o arquivo Razão aqui", type=["csv", "xlsx"])

def formatar_data(valor):
    """Transforma qualquer formato de data em DD/MM/AAAA"""
    try:
        dt = pd.to_datetime(valor)
        return dt.strftime('%d/%m/%Y')
    except:
        return str(valor).split(' ')[0] # Se falhar, tenta apenas cortar o horário

def extrair_nota(texto):
    texto = str(texto).upper()
    busca = re.search(r'(?:NFE|NF|NOTA|Nº)\s*(\d+)', texto)
    return busca.group(1) if busca else "S/N"

def limpar_texto(v):
    if pd.isna(v) or str(v).lower() in ['nan', 'none', '']: return ""
    return " ".join(str(v).split()).strip()

if arquivo:
    try:
        df = pd.read_excel(arquivo) if arquivo.name.endswith('.xlsx') else pd.read_csv(arquivo, encoding='latin-1', sep=None, engine='python')
        
        linha_topo = " ".join([str(x) for x in df.iloc[0:2].values.flatten() if pd.notna(x)])
        empresa = limpar_texto(linha_topo.upper().split("EMPRESA:")[-1].split("CNPJ:")[0]) if "EMPRESA:" in linha_topo.upper() else "EMPRESA"

        dados_finais = {}
        forn_atual = None

        for i, linha in df.iterrows():
            texto_linha = " ".join([str(v) for v in linha.values if pd.notna(v)]).upper()
            
            if "CONTA:" in texto_linha:
                cod = re.search(r'CONTA:\s*(\d+)', texto_linha)
                cod_val = cod.group(1) if cod else ""
                nome_forn = limpar_texto(texto_linha.split("CONTA:")[-1].replace('NOME:', '').replace(cod_val, ''))
                forn_atual = f"{cod_val} - {nome_forn}"
                dados_finais[forn_atual] = []
            
            elif forn_atual and ("/" in str(linha.iloc[0]) or "-" in str(linha.iloc[0])):
                def num(v):
                    try: return float(str(v).replace('.','').replace(',','.'))
                    except: return 0.0
                
                v_deb = num(linha.iloc[8])
                v_cre = num(linha.iloc[9])
                historico = limpar_texto(linha.iloc[2])
                
                if v_deb > 0 or v_cre > 0:
                    dados_finais[forn_atual].append({
                        'Data': formatar_data(linha.iloc[0]), # <--- DATA CORRIGIDA AQUI
                        'Nota Fiscal': extrair_nota(historico),
                        'Histórico': historico,
                        'Débito': v_deb,
                        'Crédito': v_cre
                    })

        saida = io.BytesIO()
        with pd.ExcelWriter(saida, engine='openpyxl') as writer:
            for forn, itens in dados_finais.items():
                if not itens: continue
                df_f = pd.DataFrame(itens)
                nome_aba = re.sub(r'[\\/*?:\[\]]', '', forn)[:31]
                df_f.to_excel(writer, sheet_name=nome_aba, index=False, startrow=9, startcol=1)
                
                ws = writer.sheets[nome_aba]
                for r in range(1, 100):
                    for c in range(1, 15): ws.cell(row=r, column=c).fill = PatternFill("solid", fgColor="FFFFFF")

                ws['B2'] = f"EMPRESA: {empresa}"
                ws['B4'] = f"FORNECEDOR: {forn}"
                
                # Ajuste de colunas
                ws.column_dimensions['B'].width = 15 # Data (espaço para DD/MM/AAAA)
                ws.column_dimensions['C'].width = 15 # Nota
                ws.column_dimensions['D'].width = 50 # Histórico
                
        st.success("✅ Datas formatadas e notas extraídas!")
        st.download_button("📥 Baixar Planilha com Datas Certas", saida.getvalue(), "conciliacao_datas_corretas.xlsx")
    
    except Exception as e:
        st.error(f"Ocorreu um erro: {e}")
