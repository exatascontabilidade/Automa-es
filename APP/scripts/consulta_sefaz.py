import pandas as pd
import sys
from datetime import datetime
import os
import builtins

# Assuma que o script consulta_sefaz está no mesmo diretório
from solicitador_xml import main as consulta_sefaz_main

# Força o print a exibir na hora
print = lambda *args, **kwargs: builtins.print(*args, **kwargs, flush=True)

def formatar_data_excel(data):
    try:
        if not data or pd.isna(data):
            return ""
        return pd.to_datetime(data).strftime("%d/%m/%Y")
    except Exception as e:
        print(f" Erro ao converter data: {data} -> {e}")
        return ""

def main():
    if len(sys.argv) < 2:
        print(" Uso: python processa_planilha.py caminho_da_planilha.xlsx")
        sys.exit(1)

    caminho_planilha = sys.argv[1]
    print(f" Executando consulta_sefaz com planilha: {caminho_planilha}")

    colunas_necessarias = ["Inscrição Municipal", "Tipo de Arquivo", "Pesquisar Por", "Data Inicial", "Data Final"]

    try:
        df = pd.read_excel(caminho_planilha, usecols=colunas_necessarias, dtype=str).fillna("")
    except Exception as e:
        print(f" Erro ao ler a planilha: {e}")
        sys.exit(1)

    for idx, row in df.iterrows():
        inscricao = row["Inscrição Municipal"].strip()
        tipo_arquivo = row["Tipo de Arquivo"].strip().upper()
        pesquisar_por = row["Pesquisar Por"].strip()
        data_ini = formatar_data_excel(row["Data Inicial"].strip())
        data_fim = formatar_data_excel(row["Data Final"].strip())

        if not inscricao:
            print(f" Linha {idx + 2}: Inscrição Municipal ausente. Pulando...")
            continue

        print(f" Consultando Inscrição Municipal: {inscricao}")

        # Mock sys.argv para o script consulta_sefaz
        sys.argv = ["solicitador_xml.py", inscricao, tipo_arquivo, pesquisar_por, data_ini, data_fim]

        try:
            consulta_sefaz_main()
        except Exception as e:
            print(f" Erro ao processar {inscricao}: {e}")

    print(" Processamento concluído para todas as empresas.")

if __name__ == "__main__":
    main()
