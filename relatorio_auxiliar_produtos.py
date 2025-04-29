
import pandas as pd
from tkinter import Tk
from tkinter.filedialog import askopenfilenames
import os

# Esconde a janela principal
Tk().withdraw()

# Selecionar de 1 a 5 arquivos
arquivos_selecionados = askopenfilenames(
    title='Selecione de 1 a 5 arquivos de estoque',
    filetypes=[('Arquivos Excel', '*.xls *.xlsx')]
)

if len(arquivos_selecionados) == 0:
    print("Nenhum arquivo selecionado. Programa encerrado.")
    exit()

# Dicionário para armazenar os DataFrames
abas = {}

# Processar cada arquivo
for caminho_arquivo in arquivos_selecionados:
    nome_loja = os.path.splitext(os.path.basename(caminho_arquivo))[0]
    try:
        # Lendo o arquivo a partir da linha 17 (skiprows=16 porque começa do 0)
        df = pd.read_excel(caminho_arquivo, skiprows=16)
    except Exception as e:
        print(f"Erro ao ler o arquivo {nome_loja}: {e}")
        continue

    if df.empty:
        continue

    # Renomeando as colunas principais manualmente para garantir
    colunas_renomeadas = {
        df.columns[0]: 'Codigo',
        df.columns[1]: 'Produto',
        df.columns[6]: 'Saldo'
    }
    df.rename(columns=colunas_renomeadas, inplace=True)

    # Selecionar apenas as colunas que importam
    if 'Saldo' not in df.columns:
        print(f"Atenção: O arquivo '{nome_loja}' não contém a coluna 'Saldo'. Pulando este arquivo.")
        continue

    # Filtrar produtos com saldo < 10
    df_filtrado = df[df['Saldo'] < 10]

    # Ordenar pelo saldo crescente
    df_filtrado = df_filtrado.sort_values(by='Saldo')

    # Mantém só as colunas principais para o relatório
    df_final = df_filtrado[['Codigo', 'Produto', 'Saldo']]

    if not df_final.empty:
        abas[nome_loja] = df_final

# Se nenhum produto foi encontrado
if not abas:
    print("Nenhum produto com saldo abaixo de 10 encontrado nos arquivos selecionados.")
    exit()

# Gerar o arquivo Excel final
with pd.ExcelWriter('relatorio_lojas_final.xlsx') as writer:
    for nome_aba, df_loja in abas.items():
        df_loja.to_excel(writer, sheet_name=nome_aba[:31], index=False)

print("Relatório gerado com sucesso como 'relatorio_lojas_final.xlsx'!")
