import pandas as pd
import requests
import time

# === CONFIGURAÇÕES ===
ARQUIVO_EMPRESAS = "Relação Empresas - Nome - CNPJ.xls"
API_URL = "https://grupoexatas.bitrix24.com.br/rest/1/ipjqgc6n2nqiaui0/crm.company.add.json"

# === FUNÇÃO DE ENVIO ===
def cadastrar_empresa(nome, cnpj):
    payload = {
        "fields": {
            "TITLE": nome,
            "UF_CRM_1702820625": cnpj  # CNPJ → substitua se o ID do campo for outro
        }
    }

    response = requests.post(API_URL, json=payload)
    if response.status_code == 200:
        result = response.json()
        if "result" in result and isinstance(result["result"], int):
            print(f"✅ '{nome}' cadastrada com sucesso (ID: {result['result']})")
        else:
            print(f"⚠️ Erro ao cadastrar '{nome}': {result}")
    else:
        print(f"❌ Erro HTTP {response.status_code} ao cadastrar '{nome}'")

# === LEITURA DO ARQUIVO ===
def ler_empresas(caminho):
    df = pd.read_excel(caminho, dtype=str)
    df = df.rename(columns=lambda col: col.strip())  # Remove espaços invisíveis
    if not {"Empresa", "CNPJ"}.issubset(df.columns):
        raise Exception("O arquivo deve conter as colunas 'Empresa' e 'CNPJ'.")
    df = df[["Empresa", "CNPJ"]].dropna()
    df = df.rename(columns={"Empresa": "NOME"})
    return df

# === EXECUÇÃO PRINCIPAL ===
def main():
    try:
        empresas = ler_empresas(ARQUIVO_EMPRESAS)
        for _, linha in empresas.iterrows():
            nome = linha["NOME"].strip()
            cnpj = linha["CNPJ"].strip().replace(".", "").replace("/", "").replace("-", "")
            cadastrar_empresa(nome, cnpj)
            time.sleep(0.3)
    except Exception as e:
        print("❌ Erro geral:", e)

if __name__ == "__main__":
    main()
