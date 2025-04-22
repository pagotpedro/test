
import streamlit as st
import pandas as pd
from datetime import datetime
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

st.set_page_config(page_title="Conversor P8Pxls", layout="centered")
st.title("Conversão de TVI em Auto de Infração")

st.write("📎 Envie um arquivo .xlsx com estrutura padrão para gerar os documentos.")

arquivo = st.file_uploader("🗂️ Escolha o arquivo Excel", type=["xlsx"])

if arquivo:
    try:
        df = pd.read_excel(arquivo)

        if 'Valor do Produto' not in df.columns:
            st.error("❌ A coluna 'Valor do Produto' não foi encontrada no arquivo.")
            st.stop()

        st.success("✅ Arquivo carregado com sucesso.")

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
            soma = df[colunas_fin].sum(numeric_only=True)
            icms_debito_sem_multa = soma.get('ICMS Débito', 0)
            multa = icms_debito_sem_multa / 2
            total_com_multa = icms_debito_sem_multa + multa

            nome_razao = df_original['Razão_4'].dropna().unique()[0]

            resumo = pd.DataFrame({
                'Descrição': [
                    f'Quadro Resumo - {nome_razao}', '',
                    'Valor total dos produtos',
                    'BC Aplicada - Base de Cálculo + 50%',
                    'ICMS Débito = Alíquota x BC',
                    'Crédito de ICMS destacado em NF-e',
                    'Valor ICMS a recolher',
                    'Multa de 50%',
                    'Total Devido (ICMS a recolher + Multa de 50%)'
                ],
                'Valor': [
                    '', '',
                    soma.get('Valor do Produto', 0),
                    soma.get('BC + 50%', 0),
                    icms_debito_sem_multa + soma.get('Valor do ICMS', 0),
                    soma.get('Valor do ICMS', 0),
                    icms_debito_sem_multa,
                    multa,
                    total_com_multa
                ]
            })

            return df, resumo, nome_razao

        df_formatado, resumo_df, razao_social = formatar_planilha(df)
        st.success(f"✅ Planilha processada para {razao_social}")

        st.dataframe(resumo_df)

        st.download_button(
            "⬇️ Baixar Planilha de Cálculo",
            df_formatado.to_csv(index=False).encode("utf-8"),
            file_name=f"Planilha de Cálculo - {razao_social}.csv",
            mime="text/csv"
        )

    except Exception as e:
        st.error(f"❌ Erro: {e}")
else:
    st.info("Aguardando envio de arquivo...")
