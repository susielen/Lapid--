import streamlit as st
import pandas as pd
import io, re
from openpyxl.styles import Font, Alignment, PatternFill

st.set_page_config(page_title="Conciliador Estável", layout="wide")
st.title("🛠️ Conciliador: Versão Corrigida")

arquivo = st.file_uploader("Suba o arquivo Razão aqui", type=["csv", "xlsx"])

def limpar_data_simples(valor):
    """Apenas remove o horário se ele existir, sem bagunçar a coluna"""
    if pd.isna(valor): return ""
    txt = str(valor).strip()
    # Se tiver espaço (ex: 2025-01-01 00:00:00), pega só a primeira parte
    return txt.split(' ')[0]

def extrair_nota(texto):
    texto = str(texto).upper()
    busca = re.search(r'(?:NFE|NF|NOTA|Nº)\s*(\d+)', texto)
    return busca.group(1) if busca else ""

def limpar_texto(v):
    if pd.isna(v) or str(v).lower() in ['nan', 'none', '']: return ""
    return " ".join(str(v).split()).strip()

if arquivo:
    try:
        # Carrega o arquivo respeitando a estrutura de vírgulas do seu CSV
        if arquivo.name.endswith('.xlsx'):
            df = pd.read_excel(arquivo)
        else:
            df = pd.read_csv(arquivo, encoding='latin-1', sep=None, engine='python')
        
        # Pega o nome da empresa (Linhas iniciais)
        topo = " ".join([str(x) for x in df.iloc[0:3].values.flatten() if pd.notna(x)])
        nome_empresa = "EMPRESA"
        if "EMPRESA:" in topo.upper():
            nome_empresa = limpar_texto(topo.upper().split("EMPRESA:")[-1].split("CNPJ:")[0])

        dados_por_forn = {}
        forn_atual = None

        for i in range(len(df)):
            linha = df.iloc[i]
            # Transforma a linha em texto para procurar a conta
            texto_linha = " ".join([str(v) for v in linha.values if pd.notna(v)]).upper()
            
            if "CONTA:" in texto_linha:
                cod = re.search(r'CONTA:\s*(\d+)', texto_linha)
                cod_val = cod.group(1) if cod else ""
                nome_bruto = texto_linha.split("CONTA:")[-1].replace('NOME:', '').strip()
                forn_atual = f"{cod_val} - {limpar_texto(nome_bruto.replace(cod_val, ''))}"
                dados_por_forn[forn_atual] = []
            
            # Identifica linhas de movimento (onde a primeira coluna tem algo que parece data)
            elif forn_atual and pd.notna(linha.iloc[0]) and any(c in str(linha.iloc[0]) for c in ['/', '-']):
                try:
                    # Ajuste de colunas baseado no seu CSV (D e C estão no final)
                    v_deb = float(str(linha.iloc[-2]).replace('.','').replace(',','.')) if pd.notna(linha.iloc[-2]) else 0
                    v_cre = float(str(linha.iloc[-1]).replace('.','').replace(',','.')) if pd.notna(linha.iloc[-1]) else 0
                    
                    if v_deb > 0 or v_cre > 0:
                        hist = limpar_texto(linha.iloc[2]) # Coluna do Histórico
                        dados_por_forn[forn_atual].append({
                            'Data': limpar_data_simples(linha.iloc[0]),
                            'Nota': extrair_nota(hist),
                            'Histórico': hist,
                            'Débito': v_deb,
                            'Crédito': v_cre
                        })
                except: continue

        # Salva o Excel
        saida = io.BytesIO()
        with pd.ExcelWriter(saida, engine='openpyxl') as writer:
            for forn, lista in dados_por_forn.items():
                if not lista: continue
                pd.DataFrame(lista).to_excel(writer, sheet_name=re.sub(r'[\\/*?:\[\]]', '', forn)[:31], index=False, startrow=5)
                
                ws = writer.sheets[re.sub(r'[\\/*?:\[\]]', '', forn)[:31]]
                ws['A1'] = f"EMPRESA: {nome_empresa}"
                ws['A2'] = f"FORNECEDOR: {forn}"
                ws.column_dimensions['C'].width = 50 # Histórico largo
                
        st.success("✅ Voltamos ao trilho! Estrutura recuperada e datas limpas.")
        st.download_button("📥 Baixar Arquivo Corrigido", saida.getvalue(), "conciliacao_final.xlsx")

    except Exception as e:
        st.error(f"Erro técnico: {e}. Por favor, verifique se o arquivo está no formato original.")
