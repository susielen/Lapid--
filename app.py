import streamlit as st
import pandas as pd
import re
from io import BytesIO

st.set_page_config(page_title="Conciliador Pro", layout="wide")
st.title("🤖 Robô Conciliador")

# Função para converter texto em número sem dar erro
def to_num(val):
    try:
        if pd.isna(val): return 0.0
        return float(str(val).replace('.', '').replace(',', '.'))
    except: return 0.0

arquivo = st.file_uploader("Suba o arquivo XLSX ou CSV", type=["xlsx", "csv"])

if arquivo:
    try:
        df_bruto = pd.read_excel(arquivo, header=None) if arquivo.name.endswith('xlsx') else pd.read_csv(arquivo, header=None)
        
        # Pega nome da empresa
        nome_emp = "EMPRESA"
        for i in range(min(15, len(df_bruto))):
            if "Empresa:" in str(df_bruto.iloc[i, 0]):
                nome_emp = str(df_bruto.iloc[i, 2])
                break

        banco = {}
        f_atual, dados = None, []

        for i in range(len(df_bruto)):
            lin = df_bruto.iloc[i]
            if "Conta:" in str(lin[0]):
                if f_atual and dados: banco[f_atual] = pd.DataFrame(dados)
                cod = str(lin[1]).strip()
                nom = str(lin[5]) if pd.notna(lin[5]) else str(lin[2])
                f_atual = f"{cod} - {nom}"
                dados = []
            elif len(lin) > 9:
                d, c = to_num(lin[8]), to_num(lin[9])
                if d != 0 or c != 0:
                    nf = re.findall(r'NFe\s?(\d+)', str(lin[2]))
                    dados.append({"Data": str(lin[0]), "NF": nf[0] if nf else str(lin[1]), "Hist": str(lin[2]), "Deb": -d, "Cred": c})

        if f_atual and dados: banco[f_atual] = pd.DataFrame(dados)

        if banco:
            out = BytesIO()
            with pd.ExcelWriter(out, engine='xlsxwriter') as writer:
                wb = writer.book
                # Estilos
                f_tit = wb.add_format({'bold':1,'align':'center','bg_color':'#D3D3D3','border':1})
                f_std = wb.add_format({'border':1})
                f_cur = wb.add_format({'num_format':'R$ #,##0.00','border':1})
                f_vde = wb.add_format({'num_format':'R$ #,##0.00','font_color':'green','bold':1,'border':1})
                f_vrm = wb.add_format({'num_format':'R$ #,##0.00','font_color':'red', 'bold':1,'border':1})

                for f, df in banco.items():
                    aba = re.sub(r'[\\/*?:\[\]]', '', f)[:31]
                    ws = wb.add_worksheet(aba)
                    ws.hide_gridlines(2)
                    
                    # Cabeçalho
                    ws.merge_range('B2:M3', f"EMPRESA: {nome_emp} | {f}", f_tit)
                    
                    # Tabela Razão (Linha 10)
                    df.to_excel(writer, sheet_name=aba, startrow=9, startcol=1, index=False)
                    
                    # Tabela Conciliação
                    res = df.groupby("NF").agg({"Deb":"sum", "Cred":"sum"}).reset_index()
                    res["Dif"] = res["Deb"] + res["Cred"]
                    res.to_excel(writer, sheet_name=aba, startrow=9, startcol=7, index=False)
                    
                    # Saldo no final
                    row = 10 + len(res)
                    saldo = res["Dif"].sum()
                    ws.write(row, 8, "Saldo Final:", f_std)
                    ws.write(row, 9, saldo, f_vde if saldo >= 0 else f_vrm)
                    ws.set_column('B:Z', 15, f_cur)

            st.success("✅ Relatório pronto!")
            st.download_button("📥 Baixar Excel", out.getvalue(), "conciliacao.xlsx")
    except Exception as e:
        st.error(f"Erro no processamento: {e}")
