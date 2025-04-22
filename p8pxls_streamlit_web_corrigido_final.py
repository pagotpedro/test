
import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

st.set_page_config(page_title="Conversão de TVI", layout="centered")
st.title("📄 Conversão de TVI em Auto de Infração")

arquivo = st.file_uploader("📎 Envie um arquivo .xlsx com estrutura padrão", type=["xlsx"])

if arquivo:
    try:
        df = pd.read_excel(arquivo)

        def formatar_planilha(df):
            if 'Data_5' in df.columns:
                df.rename(columns={'Data_5': 'Data do TVI'}, inplace=True)

            for col in ['CNPJ ou CPF', 'CNPJ ou CPF_2']:
                if col in df.columns:
                    df[col] = df[col].astype(str).str.extract(r'(\d+)')[0].fillna('').str.zfill(11)

            for col in ['Data', 'Data do TVI']:
                if col in df.columns:
                    df[col] = pd.to_datetime(df[col], errors='coerce')
                    df[col] = df[col].dt.strftime('%d/%m/%Y')

            df['Data do TVI_dt'] = pd.to_datetime(df['Data do TVI'], format='%d/%m/%Y', errors='coerce')
            df['BC + 50%'] = pd.to_numeric(df['Valor do Produto'], errors='coerce').fillna(0) * 1.5

            def calcular_aliquota(data):
                if pd.isnull(data): return None
                if data < datetime(2023, 3, 31): return 0.18
                elif data < datetime(2024, 2, 19): return 0.20
                elif data < datetime(2025, 3, 31): return 0.22
                return 0.23

            df['Aliq Interna'] = df['Data do TVI_dt'].map(calcular_aliquota)
            df['Valor do ICMS'] = pd.to_numeric(df['Valor do ICMS'], errors='coerce').fillna(0)
            df['ICMS Débito'] = df['BC + 50%'] * df['Aliq Interna'] - df['Valor do ICMS']
            df['Aliq Interna'] = df['Aliq Interna'].map(lambda x: f"{x*100:.0f}%" if pd.notnull(x) else "")
            df.drop(columns=['Data do TVI_dt'], inplace=True)

            for col in ['Base de Cálculo do ICMS ST', 'Valor do ICMS ST']:
                if col in df.columns:
                    df.drop(columns=[col], inplace=True)

            colunas_fin = ['Valor do Produto', 'Base de Cálculo ICMS', 'Valor do ICMS', 'BC + 50%', 'ICMS Débito']
            for col in colunas_fin:
                if col in df.columns:
                    df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

            df = df[df['Valor Débito TVI'] != 0]
            df_original = df.copy()

            # Inserir colunas calculadas após "Valor do ICMS"
            colunas = list(df.columns)
            if "Valor do ICMS" in colunas:
                idx = colunas.index("Valor do ICMS") + 1
                for col_insert in ["BC + 50%", "Aliq Interna", "ICMS Débito"]:
                    if col_insert in colunas:
                        colunas.remove(col_insert)
                        colunas.insert(idx, col_insert)
                        idx += 1
                df = df[colunas]

            # Calcular somatórios
            soma = df[colunas_fin].sum(numeric_only=True)
            icms_debito = soma.get('ICMS Débito', 0)
            multa = icms_debito / 2
            total_com_multa = icms_debito + multa

            linha_soma = {col: soma[col] if col in soma else '' for col in df.columns}
            linha_multa = {col: multa if col == 'ICMS Débito' else '' for col in df.columns}
            linha_total = {col: total_com_multa if col == 'ICMS Débito' else '' for col in df.columns}
            df = pd.concat([df, pd.DataFrame([linha_soma, linha_multa, linha_total])], ignore_index=True)

            return df

        df_final = formatar_planilha(df)
        st.success("✅ Arquivo processado com sucesso")
        st.dataframe(df_final)

        excel_bytes = BytesIO()
        with pd.ExcelWriter(excel_bytes, engine='xlsxwriter') as writer:
            df_final.to_excel(writer, sheet_name="Planilha Corrigida", index=False)
        excel_bytes.seek(0)

        st.download_button(
            "⬇️ Baixar Planilha de Cálculo Corrigida (.xlsx)",
            data=excel_bytes,
            file_name="Planilha Corrigida.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ Erro ao processar: {e}")
else:
    st.info("📤 Aguardando envio do arquivo Excel...")
