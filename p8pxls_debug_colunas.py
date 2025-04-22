
import os
import pandas as pd
from datetime import datetime
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

print("🚀 Iniciando execução...")

# Detectar o primeiro arquivo .xlsx na pasta
arquivos_excel = [f for f in os.listdir() if f.lower().endswith('.xlsx') and not f.startswith('~')]
if not arquivos_excel:
    raise FileNotFoundError("❌ Nenhum arquivo .xlsx encontrado na pasta.")

arquivo_excel = arquivos_excel[0]
print(f"📄 Arquivo detectado: {arquivo_excel}")

df = pd.read_excel(arquivo_excel)
print(f"✅ Arquivo carregado com {len(df)} linhas.")

# Confirmar existência de colunas
print(f"📋 Colunas disponíveis: {df.columns.tolist()}")

# Verificação de coluna crítica
if 'Valor do Produto' not in df.columns:
    raise KeyError("❌ A coluna 'Valor do Produto' não foi encontrada.")

# Formatar planilha
def formatar_planilha(df):
    print("🧮 Formatando planilha...")

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

    colunas_financeiras = ['Valor do Produto', 'Base de Cálculo ICMS', 'Valor do ICMS', 'BC + 50%', 'ICMS Débito']
    for col in colunas_financeiras:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    df = df[df['Valor Débito TVI'] != 0]
    print("✅ Planilha formatada.")

    return df

df_formatado = formatar_planilha(df)

print("✅ Processamento finalizado.")
