import pandas as pd
import os

codigos_validos = [
    202400000000669,
    202400000000809,
    202400000000914,
    202400000001125,
    202400000001156,
    202400000001267,
    202400000001295,
    202400000001336,
    202400000001582,
    202400000001593,
    202400000001595,
    202400000001627,
    202400000001701
]

# Caminho do arquivo Excel de entrada
arquivo_excel = r"C:\Users\Exatas\Desktop\Nova pasta\IMPORTAÇÃO\IMPORTACAO__RECEBIMENTOS - Julho a Dezembro 2024.xlsx"

# Nome da coluna onde os códigos estão
nome_coluna_codigo = 'Documento'  # Altere se necessário

# Lê o arquivo Excel
df = pd.read_excel(arquivo_excel)

# Filtra apenas os códigos desejados
df_filtrado = df[df[nome_coluna_codigo].isin(codigos_validos)]

# Define o caminho de saída no mesmo diretório do arquivo original
pasta_saida = os.path.dirname(arquivo_excel)
arquivo_saida = os.path.join(pasta_saida, 'FILTRO_CODIGOS.xlsx')

# Salva o resultado
df_filtrado.to_excel(arquivo_saida, index=False)

print(f'✅ Arquivo filtrado salvo em: {arquivo_saida}')
