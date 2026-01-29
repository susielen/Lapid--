import streamlit as st
import pandas as pd
import io, re
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Conciliador Turbo", layout="wide")
st.title("🧼 Conciliador Profissional (Versão Limpa)")

arquivo = st.file_uploader("Suba o arquivo Razão aqui", type=["csv", "xlsx"])

def limpar(v):
    """Apaga NANs e espaços extras"""
    txt = str(v)
    if txt.lower() in ['nan', 'none', '']: return ""
    return " ".join(txt.split()).strip()

if arquivo:
    try:
        df = pd.read_excel(arquivo) if arquivo.name.endswith('.xlsx') else pd.read_csv(arquivo, encoding='latin-1', sep=None, engine='python')
        
        # Pega o nome da empresa na primeira linha
        linha1 = " ".join([str(x) for x in df.iloc[0].values if pd.notna(x)])
        nome_empresa = limpar(linha1.upper().split("EMPRESA:")[-1].split("CNPJ:")[0]) if "EMPRESA:" in linha1.upper() else "MINHA EMPRESA"

        dados_finais = {}
        forn_atual = None

        for i, linha in df.iterrows():
            texto_linha = " ".join([str(v) for v in linha.values if pd.notna(v)]).upper()
            
            # Identifica início de novo fornecedor
            if "CONTA:" in texto_linha:
                cod = re.search(r'CONTA:\s*(\d+)', texto_linha)
                cod_val = cod.group(1) if cod else ""
                nome_bruto = texto_linha.split("CONTA:")[-1].replace('NOME:', '').strip()
                # Limpa o nome tirando números de conta repetidos e NANs
                nome_limpo = limpar(re.sub(r'(\d+\.)+\d+', '', nome_bruto).replace(cod_val, ''))
                forn_atual = f"{cod_val} - {nome_limpo}"
                dados_finais[forn_atual] = []
            
            # Identifica linhas de valores (pela data na primeira coluna)
            elif forn_atual and ("/" in str(linha.iloc[0]) or "-" in str(linha.iloc[0])):
                def num(v):
                    try: return float(str(v).replace('.','').replace(',','.'))
                    except: return 0.0
                
                v_deb = num(linha.iloc[8])
                v_cre = num(linha.iloc[9])
                
                if v_deb > 0 or v_cre > 0:
                    dados_finais[forn_atual].append({
                        'Data': linha.iloc[0],
                        'Histórico': limpar(linha.iloc[2]),
                        'Débito': v_deb,
                        'Crédito': v_cre
                    })

        saida_excel = io.BytesIO()
        with pd.ExcelWriter(saida_excel, engine='openpyxl') as writer:
            for forn, itens in dados_finais.items():
                if not itens: continue
                
                df_temp = pd.DataFrame(itens)
                nome_aba = re.sub(r'[\\/*?:\[\]]', '', forn)[:31]
                df_temp.to_excel(writer, sheet_name=nome_aba, index=False, startrow=9, startcol=1)
                
                ws = writer.sheets[nome_aba]
                
                # Limpa o fundo (tudo branco)
                for r in range(1, 100):
                    for c in range(1, 15): ws.cell(row=r, column=c).fill = PatternFill("solid", fgColor="FFFFFF")

                # Cabeçalhos Alinhados à Esquerda
                ws['B2'] = nome_empresa
                ws['B2'].font = Font(bold=True, size=12)
                ws['B4'] = forn
                ws['B4'].font = Font(bold=True, size=14)
                
                # Saldo e Totais em Negrito
                ws.cell(row=6, column=4, value="SALDO").font = Font(bold=True)
                ws.cell(row=8, column=4, value="TOTAIS").font = Font(bold=True)
                
                # Estilo de Moeda e Ajuste de Coluna
                fmt_moeda = '_-R$ * #,##0.00_-;-R$ * #,##0.00_-;_-R$ * "-"??_-;_-@_-'
                ws.column_dimensions['A'].width = 2
                ws.column_dimensions['B'].width = 15
                ws.column_dimensions['C'].width = 60 # Histórico largo
                ws.column_dimensions['D'].width = 18
                ws.column_dimensions['E'].width = 18

        st.success("🎉 Concluído! O erro dos 'NAN' foi resolvido.")
        st.download_button("📥 Baixar Planilha Limpa", saida_excel.getvalue(), "conciliacao_perfeita.xlsx")
    
    except Exception as erro:
        st.error(f"Ops, algo deu errado: {erro}")
