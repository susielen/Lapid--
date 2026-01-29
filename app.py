import streamlit as st
import pandas as pd
import io, re
from openpyxl.styles import Font, Alignment, PatternFill

st.set_page_config(page_title="Conciliador Limpo", layout="wide")
st.title("🧼 Conciliador: Versão Sem Erros (NAN Free)")

arquivo = st.file_uploader("Suba o arquivo Razão (Excel ou CSV)", type=["csv", "xlsx"])

def limpar_sujeira(texto):
    """Remove NANs, Nones e espaços extras de qualquer texto"""
    if pd.isna(texto) or str(texto).lower() in ['nan', 'none', '']:
        return ""
    # Remove a palavra 'nan' caso ela esteja no meio de uma frase
    limpo = re.sub(r'\bnan\b', '', str(texto), flags=re.IGNORECASE)
    return " ".join(limpo.split()).strip()

if arquivo:
    try:
        # Carregamento do arquivo
        if arquivo.name.endswith('.xlsx'):
            df_raw = pd.read_excel(arquivo)
        else:
            df_raw = pd.read_csv(arquivo, encoding='latin-1', sep=None, engine='python')

        # 1. PEGAR NOME DA EMPRESA (Geralmente no topo)
        linha_cabecalho = " ".join([str(v) for v in df_raw.iloc[0:2].values.flatten() if pd.notna(v)])
        empresa_nome = "EMPRESA NÃO IDENTIFICADA"
        if "EMPRESA:" in linha_cabecalho.upper():
            empresa_nome = linha_cabecalho.upper().split("EMPRESA:")[-1].split("CNPJ:")[0].strip()

        resumo_por_fornecedor = {}
        fornecedor_atual = None

        for i, linha in df_raw.iterrows():
            # Transforma a linha em texto para busca
            texto_da_linha = " ".join([str(v) for v in linha.values if pd.notna(v)]).upper()
            
            # 2. DETECTAR FORNECEDOR (E limpar os NANs do nome)
            if "CONTA:" in texto_da_linha:
                cod_match = re.search(r'CONTA:\s*(\d+)', texto_da_linha)
                cod_conta = cod_match.group(1) if cod_match else ""
                
                # Pega o que vem depois de CONTA: ou NOME:
                partes_nome = texto_da_linha.split("CONTA:")[-1]
                nome_limpo = limpar_sujeira(partes_nome.replace(cod_conta, "").replace("NOME:", ""))
                
                fornecedor_atual = f"{cod_conta} - {nome_limpo}" if cod_conta else nome_limpo
                resumo_por_fornecedor[fornecedor_atual] = []
            
            # 3. CAPTURAR LANÇAMENTOS (Linhas que começam com data)
            elif fornecedor_atual and ("/" in str(linha.iloc[0])):
                def converter_valor(v):
                    try: return float(str(v).replace('.', '').replace(',', '.'))
                    except: return 0.0
                
                deb = converter_valor(linha.iloc[8])
                cre = converter_valor(linha.iloc[9])
                
                if deb > 0 or cre > 0:
                    resumo_por_fornecedor[fornecedor_atual].append({
                        'Data': linha.iloc[0],
                        'Histórico': limpar_sujeira(linha.iloc[2]),
                        'Débito': deb,
                        'Crédito': cre
                    })

        # GERAR EXCEL FINAL
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            for forn, lancamentos in resumo_por_fornecedor.items():
                if not lancamentos: continue
                
                # Criar Aba (nome limitado a 31 caracteres)
                nome_aba = re.sub(r'[\\/*?:\[\]]', '', forn)[:31]
                df_f = pd.DataFrame(lancamentos)
                df_f.to_excel(writer, sheet_name=nome_aba, index=False, startrow=9, startcol=1)
                
                ws = writer.sheets[nome_aba]
                alinhamento_esquerda = Alignment(horizontal='left', vertical='center')
                
                # Aplicar Fundo Branco e Alinhamento
                for r in range(1, 50):
                    for c in range(1, 15):
                        ws.cell(row=r, column=c).fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

                # Escrever Cabeçalhos com Alinhamento à Esquerda
                ws['B2'] = f"EMPRESA: {empresa_nome}"
                ws['B2'].font = Font(bold=True, size=11)
                ws['B2'].alignment = alinhamento_esquerda
                
                ws['B4'] = f"FORNECEDOR: {forn}"
                ws['B4'].font = Font(bold=True, size=13)
                ws['B4'].alignment = alinhamento_esquerda
                
                # Títulos de Saldo
                ws.cell(row=6, column=4, value="SALDO ANTERIOR:").font = Font(bold=True)
                ws.cell(row=8, column=4, value="TOTAIS DO PERÍODO:").font = Font(bold=True)

                # Ajustar largura das colunas
                ws.column_dimensions['B'].width = 15 # Data
                ws.column_dimensions['C'].width = 65 # Histórico
                ws.column_dimensions['D'].width = 15 # Débito
                ws.column_dimensions['E'].width = 15 # Crédito

        st.success("✅ Processado com sucesso! Os 'NAN' foram removidos.")
        st.download_button("📥 Baixar Planilha Limpa", output.getvalue(), "conciliacao_organizada.xlsx")

    except Exception as e:
        st.error(f"Ocorreu um erro no processamento: {e}")
