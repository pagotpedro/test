
import streamlit as st

st.set_page_config(
    page_title="SEFAZ-MA - Fiscalização de Mercadorias",
    page_icon="🛃",
    layout="centered"
)

# Cabeçalho com identidade visual
URL_BRASAO = "https://raw.githubusercontent.com/pedropagot/P8Pxls/main/brasao.png"

st.markdown(f'''
    <div style="text-align: center; margin-bottom: 15px;">
        <img src="{URL_BRASAO}" width="100" style="margin-bottom: 0;">
        <h2 style="margin-top: 0;">SECRETARIA DE ESTADO DA FAZENDA - MA</h2>
        <h3 style="color: #004080;">FISCALIZAÇÃO DE MERCADORIAS EM TRÂNSITO</h3>
    </div>
    <hr style="border: 1px solid #004080;">
''', unsafe_allow_html=True)


import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
import re

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

    if 'Valor do Produto' in df.columns:
        df['BC + 50%'] = pd.to_numeric(df['Valor do Produto'], errors='coerce').fillna(0) * 1.5

    def calcular_aliquota(data):
        if pd.isnull(data):
            return None
        elif data < datetime(2023, 3, 31):
            return 0.18
        elif data < datetime(2024, 2, 19):
            return 0.20
        elif data < datetime(2025, 3, 31):
            return 0.22
        else:
            return 0.23

    df['Aliq Interna'] = df['Data do TVI_dt'].map(calcular_aliquota)

    df['Valor do ICMS'] = pd.to_numeric(df['Valor do ICMS'], errors='coerce').fillna(0)
    df['ICMS Débito'] = df['BC + 50%'] * df['Aliq Interna'] - df['Valor do ICMS']
    df['Aliq Interna'] = df['Aliq Interna'].map(lambda x: f"{x*100:.0f}%" if pd.notnull(x) else "")
    df.drop(columns=['Data do TVI_dt'], inplace=True)

    for coluna_remover in ['Base de Cálculo do ICMS ST', 'Valor do ICMS ST']:
        if coluna_remover in df.columns:
            df.drop(columns=[coluna_remover], inplace=True)

    colunas_financeiras = [
        'Valor do Produto', 'Base de Cálculo ICMS', 'Valor do ICMS',
        'BC + 50%', 'Débito ICMS', 'ICMS Débito'
    ]
    colunas_presentes = [col for col in colunas_financeiras if col in df.columns]
    for col in colunas_presentes:
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    if 'Valor Débito TVI' in df.columns:
        df['Valor Débito TVI'] = pd.to_numeric(df['Valor Débito TVI'], errors='coerce').fillna(0)
        df = df[df['Valor Débito TVI'] != 0]

    df_original = df.copy()

    soma = df[colunas_presentes].sum(numeric_only=True)
    icms_debito_sem_multa = soma.get('ICMS Débito', 0)
    multa = icms_debito_sem_multa / 2
    total_com_multa = icms_debito_sem_multa + multa

    linha_soma = {col: soma[col] if col in soma else '' for col in df.columns}
    linha_multa = {col: multa if col == 'ICMS Débito' else '' for col in df.columns}
    linha_total_multa = {col: total_com_multa if col == 'ICMS Débito' else '' for col in df.columns}
    df = pd.concat([df, pd.DataFrame([linha_soma, linha_multa, linha_total_multa])], ignore_index=True)

    colunas_novas = ['BC + 50%', 'Aliq Interna', 'ICMS Débito']
    if 'Valor da NFe' in df.columns:
        cols = list(df.columns)
        idx = cols.index('Valor da NFe') + 1
        for col in reversed([c for c in colunas_novas if c in cols]):
            cols.insert(idx, cols.pop(cols.index(col)))
        df = df[cols]

    nome_razao = df_original['Razão_4'].dropna().unique()[0]
    nome_arquivo = f"GFIS_{re.sub(r'[^a-zA-Z0-9]+', '_', nome_razao)}.xls"

    resumo = pd.DataFrame({
        'Descrição': [
            f'Quadro Resumo - {nome_razao}',
            '',
            'Valor total dos produtos',
            'BC Aplicada - Base de Cálculo + 50%',
            'ICMS Débito = Alíquota x BC',
            'Crédito de ICMS destacado em NF-e',
            'Valor ICMS a recolher',
            'Multa de 50%',
            'Total Devido (ICMS a recolher + Multa de 50%)'
        ],
        'Valor': [
            '', '',  # título e espaço
            soma.get('Valor do Produto', 0),
            soma.get('BC + 50%', 0),
            icms_debito_sem_multa + soma.get('Valor do ICMS', 0),
            soma.get('Valor do ICMS', 0),
            icms_debito_sem_multa,
            multa,
            total_com_multa
        ]
    })

    gfis_df = pd.DataFrame()
    gfis_df['Inscrição Renavam_3'] = df_original['Inscrição Renavam_3'].astype(str).str.extract(r'(\d+)')[0].fillna('')
    gfis_df['CNPJ ou CPF_2'] = df_original['CNPJ ou CPF_2'].astype(str)
    gfis_df['Data do TVI'] = df_original['Data do TVI'].astype(str)
    gfis_df['ICMS Débito'] = df_original['ICMS Débito'].astype(float).round(2)

    return df, resumo, gfis_df, nome_arquivo

# Streamlit App
st.title("Conversor P8Pxls 📊")
st.write("Envie seu arquivo Excel `.xls` ou `.xlsx` e baixe os arquivos gerados")

arquivo = st.file_uploader("Selecione o arquivo Excel", type=["xls", "xlsx"])

if arquivo:
    try:
        extensao = arquivo.name.split('.')[-1].lower()
        if extensao == 'xls':
            df = pd.read_excel(arquivo, engine='xlrd')
        else:
            df = pd.read_excel(arquivo, engine='openpyxl')

        df_formatado, resumo, gfis, nome_arquivo_gfis = formatar_planilha(df)

        output_principal = BytesIO()
        with pd.ExcelWriter(output_principal, engine='xlsxwriter') as writer:
            df_formatado.to_excel(writer, index=False, sheet_name='Planilha Formatada')
            resumo.to_excel(writer, index=False, sheet_name='Quadro Resumo')

        output_gfis = BytesIO()
        with pd.ExcelWriter(output_gfis, engine='xlsxwriter') as writer:
            gfis.to_excel(writer, index=False, sheet_name='GFIS + Razão_4')

        st.success("✅ Arquivos gerados com sucesso!")

        st.download_button(
            label="📥 Baixar Planilha Formatada + Quadro Resumo",
            data=output_principal.getvalue(),
            file_name="P8Pxls_formatado_financeiro.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.download_button(
            label=f"📥 Baixar {nome_arquivo_gfis}",
            data=output_gfis.getvalue(),
            file_name=nome_arquivo_gfis,
            mime="application/vnd.ms-excel"
        )

    except Exception as e:
        st.error(f"❌ Erro ao processar o arquivo: {e}")



        # Gerar Quadro Resumo com título omitido e cabeçalho em A3/B3
        resumo_df = pd.DataFrame({
            'Descrição': [
                'Valor total dos produtos',
                'BC Aplicada - Base de Cálculo + 50%',
                'ICMS Débito = Alíquota x BC',
                'Crédito de ICMS destacado em NF-e',
                'Valor ICMS a recolher',
                'Multa de 50%',
                'Total Devido (ICMS a recolher + Multa de 50%)'
            ],
            'Valor': [
                total_produtos,
                total_bc_mais_50,
                total_icms_debito + total_valor_icms,
                total_valor_icms,
                total_icms_debito,
                multa,
                total_com_multa
            ]
        })

        resumo_df.to_excel(writer, sheet_name='Quadro Resumo', startrow=3, index=False, header=False)
        ws2 = writer.sheets['Quadro Resumo']

        # Estilos
        header_fmt = workbook.add_format({
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'bg_color': '#D9D9D9',
            'border': 1
        })
        valor_fmt = workbook.add_format({'num_format': 'R$ #,##0.00'})

        # Cabeçalho manual na linha 3
        ws2.write('A3', 'Descrição', header_fmt)
        ws2.write('B3', 'Valor', header_fmt)
        ws2.set_column('A:A', 55)
        ws2.set_column('B:B', 25, valor_fmt)


# Rodapé
st.markdown("<hr>", unsafe_allow_html=True)
st.markdown(
    "<div style='text-align: center; font-size: small; color: grey;'>"
    "Governo do Estado do Maranhão - SEFAZ/MA | © 2025</div>",
    unsafe_allow_html=True
)
