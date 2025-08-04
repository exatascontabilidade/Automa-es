import os
import json
import re

def carregar_lista_empresas(temp_path):
    """
    Carrega a lista de empresas do arquivo JSON na pasta temp.
    
    :param temp_path: Caminho da pasta onde está o arquivo JSON.
    :return: Dicionário de empresas {CNPJ: código}.
    """
    json_path = os.path.join(temp_path, "empresas.json")

    if not os.path.exists(json_path):
        print("❌ Arquivo empresas.json não encontrado!")
        return {}

    with open(json_path, "r", encoding="utf-8") as file:
        try:
            lista_empresas = json.load(file)
            if not isinstance(lista_empresas, list):
                print("❌ O JSON não está no formato correto!")
                return {}

            # Criar dicionário com chave CNPJ (limpo) e valor código
            empresas_dict = {
                str(empresa["cnpj"]): str(empresa["codigo"])
                for empresa in lista_empresas if "cnpj" in empresa and "codigo" in empresa
            }
            return empresas_dict
        except json.JSONDecodeError:
            print("❌ Erro ao ler o JSON!")
            return {}

def extrair_cnpj_do_nome(nome_arquivo):
    """
    Extrai o CNPJ do nome do arquivo, se existir.
    :param nome_arquivo: Nome do arquivo PDF.
    :return: CNPJ em formato numérico (somente dígitos) ou None.
    """
    match = re.search(r'(\d{14})', nome_arquivo)
    if match:
        return match.group(1)
    return None

def analisar_pdfs_na_pasta(temp_path, empresas_por_cnpj):
    """
    Analisa os PDFs na pasta e retorna os códigos das empresas com CNPJ correspondente.
    
    :param temp_path: Caminho da pasta onde os PDFs estão armazenados.
    :param empresas_por_cnpj: Dicionário {CNPJ: código}.
    :return: Lista de códigos das empresas que possuem PDFs na pasta.
    """
    arquivos_pasta = os.listdir(temp_path)
    arquivos_pdf = [arquivo for arquivo in arquivos_pasta if arquivo.lower().endswith(".pdf")]

    empresas_com_pdf = set()

    for arquivo in arquivos_pdf:
        cnpj_encontrado = extrair_cnpj_do_nome(arquivo)
        if cnpj_encontrado and cnpj_encontrado in empresas_por_cnpj:
            empresas_com_pdf.add(empresas_por_cnpj[cnpj_encontrado])

    return list(empresas_com_pdf)

def obter_codigo_empresa():
    """
    Retorna uma lista de códigos das empresas que possuem PDFs na pasta temp com CNPJ válido.
    """
    temp_path = os.path.join(os.path.dirname(__file__), "temp")

    empresas_por_cnpj = carregar_lista_empresas(temp_path)
    if not empresas_por_cnpj:
        return []

    return analisar_pdfs_na_pasta(temp_path, empresas_por_cnpj)
