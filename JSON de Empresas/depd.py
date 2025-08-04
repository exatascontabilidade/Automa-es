import pandas as pd
import json
import os

diretorio_base = os.path.dirname(os.path.abspath(__file__))

# Nome do arquivo Excel
nome_arquivo_excel = "gestta-clientes (4).xlsx"
caminho_arquivo_excel = os.path.join(diretorio_base, nome_arquivo_excel)

# Ler o arquivo Excel
df = pd.read_excel(caminho_arquivo_excel)

# Selecionar e renomear as colunas relevantes
df_filtrado = df[['Código', 'Nome', 'CNPJ']].rename(columns={
    'Código': 'codigo',
    'Nome': 'nome',
    'CNPJ': 'cnpj'
})

# Converter para uma lista de dicionários
dados_json = df_filtrado.to_dict(orient='records')

# Caminho de saída do JSON
caminho_arquivo_json = os.path.join(diretorio_base, "empresas.json")

# Salvar o JSON
with open(caminho_arquivo_json, "w", encoding="utf-8") as f:
    json.dump(dados_json, f, ensure_ascii=False, indent=2)

print("Arquivo 'empresas.json' gerado com sucesso!")
