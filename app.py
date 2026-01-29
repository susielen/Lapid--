import streamlit as st
import pandas as pd
import io, re
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Conciliador Limpo", layout="wide")
st.title("🧼 Conciliador: Versão Sem Erros")

arquivo = st.file_uploader("Suba o Razão (Excel ou CSV)", type=["csv", "xlsx"])

def limpar_texto(txt):
    """Remove NANs, espaços extras e limpa o nome do fornecedor"""
    if pd.isna(txt): return ""
    t = str(txt).replace('nan', '').replace('NAN', '').replace('NaN', '')
    return " ".join(t.split()).strip()

if arquivo:
    try:
        df_raw = pd.read_excel(arquivo) if arquivo.name.endswith('.xlsx') else pd.read_csv(arquivo, encoding='latin-1', sep=None, engine='python')
        
        # 1. PEGAR O NOME DA EMPRESA (Linha 2)
        linha_empresa = " ".join([str(v) for v in df_raw.iloc[0].values if pd.notna(v)])
        empresa_nome = "EMPRESA"
        if "EMPRESA:" in linha_empresa.upper():
            empresa_nome = linha_empresa.upper().split("EMPRESA:")[-1].split("CNPJ:")[0].strip()

        resumo = {}
        fornecedor_atual = None

        for i, linha in df_raw.iterrows():
            texto_linha = " ".join([str(v) for v in linha.values if pd.notna(v)]).upper()
            
            # 2. IDENTIFICAR FORNECEDOR (Tira os NANs aqui)
            if "CONTA:" in texto_linha:
                codigo = re.search(r'CONTA:\s*(\d+)', texto_linha)
                cod_val = codigo.group(1) if codigo else ""
                
                # Limpa o nome removendo o código e palavras inúteis
                nome_bruto = texto_linha.split("CONTA:")[-1].replace('NOME:', '').strip()
                nome_limpo = limpar_texto(re.sub(r'(\d+\.)+\d+', '', nome_bruto).replace(cod_val, ''))
                
                fornecedor_atual = f"{cod_val} - {nome_limpo}" if cod_val else nome_limpo
                resumo[fornecedor_atual] = []
            
            # 3. PEGAR OS DADOS (Lançamentos)
            elif fornecedor_atual and ("/" in str(linha.iloc[0])):
                def para_float(val):
                    try: return float(str(val).replace('.', '').replace(',', '.'))
                    except: return 0.0
                
                debito = para_float(linha.iloc[8])
                credito = para_float(linha.iloc[9])
                
                if debito > 0 or credito > 0:
                    resumo[fornecedor_atual].append({
                        'Data': linha.iloc[0],
                        'Histórico': limpar_texto(linha.iloc[2]),
                        'Débito': debito,
                        'Crédito': credito
                    })

        # GERAR O EXCEL
        saida = io.BytesIO()
        with pd.ExcelWriter(saida, engine='openpyxl') as writer:
            for forn, dados in resumo.items():
                if not dados: continue
                
                df_f = pd.DataFrame(dados)
                aba = re.sub(r'[\\/*?:\[\]]', '', forn)[:31]
                df_f.to_excel(writer, sheet_name=aba, index=False, startrow=9, startcol=1)
                
                ws = writer.sheets[aba]
                
                # FORMATAÇÃO VISUAL
                branco = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
                for r in range(1, 100):
                    for c in range(1, 20): ws.cell(row=r, column=c).fill = branco
                
                # Cabeçalho Alinhado à Esquerda
                ws.merge_cells('B2:J2'); ws['B2'] = empresa_nome
                ws['B2'].font = Font(bold=True, size=12)
                
                ws.merge_cells('B4:J4'); ws['B4'] = forn
                ws['B4'].font = Font(bold=True, size=14)
                
                # Negrito no SALDO e TOTAIS
                ws.cell(row=6, column=4, value="SALDO").font = Font(bold=True)
                ws.cell(row=8, column=4, value="TOTAIS").font = Font(bold=True)
                
                # Ajuste de Colunas
                ws.column_dimensions['A'].width = 2
                ws.column_dimensions['D'].width = 50 # Histórico maior
                
        st.success("✅ Agora sim! O arquivo foi limpo e organizado.")
        st.download_button("📥 Baixar Planilha Corrigida", saida.getvalue(), "conciliacao_limpa.xlsx")

    except Exception as e:
        st.error(f"Erro ao processar: {e}")
