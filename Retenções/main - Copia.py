import pandas as pd
import os

# 📥 Caminho da planilha original
caminho_arquivo = r"C:\Users\Exatas\Desktop\APP - Copia\Retencoes_Separadas\Retencao_CAIXA_ECONOMICA_FEDERAL.xlsx

# 📄 Lê a planilha
df = pd.read_excel(caminho_arquivo, header=None)

# 🧹 Remove linhas que contêm "Total:" ou "Cliente:" na primeira coluna
df_filtrado = df[~df[0].astype(str).str.strip().isin(["Total:", "Cliente:"])]

# 📁 Define caminho para salvar no mesmo diretório
diretorio = os.path.dirname(caminho_arquivo)
nome_arquivo_original = os.path.basename(caminho_arquivo)
nome_arquivo_sem_extensao = os.path.splitext(nome_arquivo_original)[0]
novo_arquivo = os.path.join(diretorio, f"{nome_arquivo_sem_extensao}_filtrado.xlsx")

# 💾 Salva o novo arquivo
df_filtrado.to_excel(novo_arquivo, index=False, header=False)

print(f"✅ Arquivo salvo em: {novo_arquivo}")
