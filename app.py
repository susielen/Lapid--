import streamlit as st
import pandas as pd
import io, re
from openpyxl.styles import Font, Alignment, PatternFill

st.set_page_config(page_title="Conciliador Turbo", layout="wide")
st.title("🧼 Conciliador Profissional (Versão Limpa)")

arquivo = st.file_uploader("Suba o arquivo Razão aqui", type=["csv", "xlsx"])

def limpar_texto(v):
    """Remove NANs, Nones e espaços extras"""
    txt = str(v)
    # Remove termos nulos e limpa espaços
    if txt.lower() in ['nan', 'none', '', 'nan nan']: return ""
    txt = re.sub(r'\bnan\b', '', txt, flags=re.IGNORECASE)
    return " ".join(txt.split()).strip()

if arquivo:
    try:
        # Carregamento inteligente
        if arquivo.name.endswith('.xlsx'):
            df = pd.read_excel(arquivo)
        else:
            df = pd.read_csv(arquivo, encoding='latin-1', sep=None, engine='python')
        
        # Identifica a Empresa (Limpando NANs)
        linha1 = " ".join([str(x) for x in df.iloc[0:2].values.flatten() if pd.notna(x)])
        nome_empresa = limpar_texto(linha1.upper().split("EMPRESA:")[-1].split("CNPJ:")[0]) if "EMPRESA:" in linha1.upper() else "MINHA EMPRESA"

        dados_por_fornecedor = {}
        forn_atual = None

        for i, linha in df.iterrows():
            texto_linha = " ".join([str(v) for v in linha.values if pd.notna(v)]).upper()
            
            # Identifica início de novo fornecedor (CONTA: XXXX - NOME)
            if "CONTA:" in texto_linha:
                cod = re.search(r'CONTA:\s*(\d+)', texto_linha)
                cod_val = cod.group(1) if cod else ""
                nome_bruto = texto_linha.split("CONTA:")[-1].replace('NOME:', '').strip()
                # Limpa o nome de qualquer NAN residual
                nome_limpo = limpar_texto(nome_bruto.replace(cod_val, ''))
                forn_atual = f"{cod_val} - {nome_limpo}"
                dados_por_fornecedor[forn_atual] = []
            
            # Identifica linhas de valores (pela data na primeira coluna)
            elif forn_atual and ("/" in str(linha.iloc[0]) or "-" in str(linha.iloc[0])):
                def para_float(v):
                    try: return float(str(v).replace('.','').replace(',','.'))
                    except: return 0.0
                
                v_deb = para_float(linha.iloc[8]) # Coluna I
                v_cre = para_float(linha.iloc[9]) # Coluna J
                
                if v_deb > 0 or v_cre > 0:
                    dados_por_fornecedor[forn_atual].append({
                        'Data': linha.iloc[0],
                        'Histórico': limpar_texto(linha.iloc[2]),
                        'Débito': v_deb,
                        'Crédito': v_cre
                    })

        # Geração do arquivo Excel
        saida_excel = io.BytesIO()
        with pd.ExcelWriter(saida_excel, engine='openpyxl') as writer:
            for forn, itens in dados_por_fornecedor.items():
                if not itens: continue
                
                df_temp = pd.DataFrame(itens)
                nome_aba = re.sub(r'[\\/*?:\[\]]', '', forn)[:31]
                df_temp.to_excel(writer, sheet_name=nome_aba, index=False, startrow=9, startcol=1)
                
                ws = writer.sheets[nome_aba]
                align_left = Alignment(horizontal='left', vertical='center')
                
                # Deixar tudo com fundo branco
                for r in range(1, 100):
                    for c in range(1, 15):
                        ws.cell(row=r, column=c).fill = PatternFill("solid", fgColor="FFFFFF")

                # Cabeçalhos Alinhados à Esquerda
                ws['B2'] = f"EMPRESA: {nome_empresa}"
                ws['B2'].font = Font(bold=True, size=11)
                ws['B2'].alignment = align_left
                
                ws['B4'] = f"FORNECEDOR: {forn}"
                ws['B4'].font = Font(bold=True, size=13)
                ws['B4'].alignment = align_left
                
                # Formatação de colunas
                ws.column_dimensions['B'].width = 15
                ws.column_dimensions['C'].width = 65 # Histórico bem largo
                ws.column_dimensions['D'].width = 18
                ws.column_dimensions['E'].width = 18

        st.success("🎉 Erro resolvido! NANs eliminados e alinhamento corrigido.")
        st.download_button("📥 Baixar Planilha Corrigida", saida_excel.getvalue(), "conciliacao_limpa.xlsx")
    
    except Exception as erro:
        st.error(f"Ainda há um detalhe no arquivo: {erro}")
