import os
import time
import json
import requests
import pandas as pd
from collections import defaultdict
from datetime import datetime
from requests.adapters import HTTPAdapter, Retry

# === CONFIGURAÇÕES GERAIS ===
APP_KEY = "2046507286824"
APP_SECRET = "cafa4a6f94f7eb3b1ae7fab5a710f1d6"
DATA_DESEJADA = "26/07/2025"
PASTA_RAIZ = os.path.dirname(os.path.abspath(__file__))
PASTA_PDFS = os.path.join(PASTA_RAIZ, "temp")
ARQUIVO_JSON_SAIDA = os.path.join(PASTA_RAIZ, "download.json")
ARQUIVO_EXCEL = os.path.join(PASTA_RAIZ, "Relação Empresas - Nome - CNPJ.xls")

# === URLS DAS APIS ===
URL_LISTAR = "https://app.omie.com.br/api/v1/financas/contareceber/"
URL_CLIENTE = "https://app.omie.com.br/api/v1/geral/clientes/"
URL_OBTER_BOLETO = "https://app.omie.com.br/api/v1/financas/contareceberboleto/"
URL_PIX = "https://app.omie.com.br/api/v1/financas/pix/"

# === CONTROLE DE NOMEAÇÃO E CACHE DE CLIENTES ===
cache_clientes = {}
os.makedirs(PASTA_PDFS, exist_ok=True)

# === SANITIZAÇÃO DE NOMES ===
def sanitize_nome(nome):
    return ''.join(c for c in nome if c.isalnum() or c in (' ', '-', '_')).strip()

def gerar_nome_pdf(data_vencimento, nome_empresa, cod_titulo):
    data_formatada = data_vencimento.replace("/", "-")
    competencia = datetime.now().strftime("%m-%Y")  # Competência com base na data atual
    nome_sanitizado = sanitize_nome(nome_empresa)
    return f"BOLETO - COMPETENCIA {competencia} - VENC {data_formatada} - {nome_sanitizado} - {cod_titulo}.pdf"



# === CARREGAR PLANILHA DOMÍNIO ===
def carregar_codigos_dominio():
    if not os.path.exists(ARQUIVO_EXCEL):
        print(f"❌ Arquivo de empresas não encontrado: {ARQUIVO_EXCEL}")
        return {}
    try:
        df = pd.read_excel(ARQUIVO_EXCEL, dtype=str, engine="xlrd")
        df.columns = [col.strip().lower() for col in df.columns]
        if "cód." not in df.columns or "cnpj" not in df.columns:
            print("❌ Colunas 'Cód.' ou 'CNPJ' não encontradas.")
            return {}
        df["cnpj"] = df["cnpj"].str.replace(r"\D", "", regex=True)
        df = df.rename(columns={"cód.": "codigo_empresa_dominio"})
        return dict(zip(df["cnpj"], df["codigo_empresa_dominio"]))
    except Exception as e:
        print(f"❌ Erro ao carregar planilha: {e}")
        return {}

codigos_dominio = carregar_codigos_dominio()

# === REQUISIÇÃO PADRÃO ===
def requisicao_omie(url, call, params):
    payload = {
        "call": call,
        "app_key": APP_KEY,
        "app_secret": APP_SECRET,
        "param": [params]
    }
    try:
        resposta = requests.post(url, json=payload)
        return resposta.json()
    except Exception as e:
        print(f"⚠️ Erro na requisição {call}: {e}")
        return {}

# === DADOS DO CLIENTE ===
def obter_dados_cliente(codigo_cliente):
    if codigo_cliente in cache_clientes:
        return cache_clientes[codigo_cliente]
    resposta = requisicao_omie(URL_CLIENTE, "ConsultarCliente", {
        "codigo_cliente_omie": codigo_cliente
    })
    cnpj = resposta.get("cnpj_cpf", "").replace(".", "").replace("-", "").replace("/", "")
    nome = resposta.get("razao_social", "").strip()
    if cnpj and nome:
        cache_clientes[codigo_cliente] = (cnpj, nome)
    return cache_clientes.get(codigo_cliente, ("", ""))

# === DOWNLOAD BOLETO COM ROBUSTEZ ===
def baixar_pdf(link, nome_arquivo, tentativas=3, timeout=10):
    session = requests.Session()
    retries = Retry(
        total=tentativas,
        backoff_factor=2,
        status_forcelist=[500, 502, 503, 504],
        allowed_methods=["GET"]
    )
    session.mount("https://", HTTPAdapter(max_retries=retries))

    try:
        print(f"⬇️ Baixando PDF: {nome_arquivo}")
        response = session.get(link, timeout=timeout)
        if response.status_code == 200:
            path = os.path.join(PASTA_PDFS, nome_arquivo)
            with open(path, 'wb') as f:
                f.write(response.content)
            print(f"✅ PDF salvo: {nome_arquivo}")
        else:
            print(f"❌ Erro {response.status_code} ao baixar {nome_arquivo} - Link: {link}")
    except requests.exceptions.RequestException as e:
        print(f"⚠️ Erro de rede ao baixar {nome_arquivo}: {e}")
    except Exception as e:
        print(f"⚠️ Erro inesperado ao salvar {nome_arquivo}: {e}")

# === OBTER LINK E BAIXAR (BOLETO) ===
def baixar_boleto_pdf_via_url(nCodTitulo, data_vencimento, nome_empresa):
    nome_pdf = gerar_nome_pdf(data_vencimento, nome_empresa, nCodTitulo)
    response = requisicao_omie(URL_OBTER_BOLETO, "ObterBoleto", {
        "nCodTitulo": nCodTitulo,
        "cCodIntTitulo": ""
    })
    link = response.get("cLinkBoleto")
    if link:
        baixar_pdf(link, nome_pdf)
    else:
        print(f"❌ Link não encontrado para {nCodTitulo}")


# === PROCESSO PRINCIPAL ===
def baixar_documentos():
    boletos = []
    erros = []

    primeira_pag = requisicao_omie(URL_LISTAR, "ListarContasReceber", {
        "pagina": 1, "registros_por_pagina": 50, "apenas_importado_api": "N"
    })
    if not primeira_pag:
        print("❌ Falha ao obter a primeira página.")
        return

    total_paginas = primeira_pag.get("total_de_paginas", 1)
    paginas_sem_data = 0

    for pagina in range(total_paginas, 0, -1):
        print(f"\n🔍 Página {pagina}")
        resposta = requisicao_omie(URL_LISTAR, "ListarContasReceber", {
            "pagina": pagina,
            "registros_por_pagina": 50,
            "apenas_importado_api": "N"
        })

        contas = resposta.get("conta_receber_cadastro", [])
        encontrou_na_data = False

        for conta in contas:
            if conta.get("data_emissao") != DATA_DESEJADA:
                continue

            encontrou_na_data = True
            tipo = conta.get("codigo_tipo_documento")
            cod_cliente = conta.get("codigo_cliente_fornecedor")
            cod_titulo = conta.get("codigo_lancamento_omie")

            if not cod_cliente or not cod_titulo:
                continue

            cnpj, nome_empresa = obter_dados_cliente(cod_cliente)
            if not cnpj or not nome_empresa:
                print(f"⚠️ Dados ausentes para cliente {cod_cliente}")
                nome_pdf = gerar_nome_pdf(conta.get("data_vencimento"), nome_empresa, cod_titulo)
                erros.append({
                    "id_boleto": cod_titulo,
                    "codigo_cliente_fornecedor": cod_cliente,
                    "motivo": "CNPJ ou nome da empresa não encontrado",
                    "nome_pdf": nome_pdf
                })
                continue

            codigo_dominio = codigos_dominio.get(cnpj, None)
            if not codigo_dominio:
                nome_pdf = gerar_nome_pdf(conta.get("data_vencimento"), nome_empresa, cod_titulo)
                erros.append({
                    "id_boleto": cod_titulo,
                    "cnpj": cnpj,
                    "nome_empresa": nome_empresa,
                    "codigo_cliente_fornecedor": cod_cliente,
                    "motivo": "CNPJ não encontrado na planilha",
                    "nome_pdf": nome_pdf
                })

            # === BOLETO ===
            if tipo == "BOL" and conta.get("boleto", {}).get("cGerado") == "S":
                nome_pdf = gerar_nome_pdf(conta.get("data_vencimento"), nome_empresa, cod_titulo)
                boletos.append({
                    "id_boleto": cod_titulo,
                    "cnpj": cnpj,
                    "nome_empresa": nome_empresa,
                    "codigo_cliente_fornecedor": cod_cliente,
                    "codigo_empresa_dominio": codigo_dominio or "NÃO ENCONTRADO",
                    "data_vencimento": conta.get("data_vencimento"),
                    "nome_pdf": nome_pdf
                })
                baixar_boleto_pdf_via_url(cod_titulo, conta.get("data_vencimento"), nome_empresa)

            # === PIX ===
            elif tipo == "PIX":
                pix = requisicao_omie(URL_PIX, "ObterPix", {
                    "nIdPix": 0,
                    "cCodIntPix": "",
                    "nCodTitulo": cod_titulo
                })
                if pix:
                    link = pix.get("cUrlPix") or pix.get("cLinkBoleto")
                    if link:
                        data_venc = conta.get("data_vencimento", "")
                        nome_pdf = gerar_nome_pdf(data_venc, nome_empresa, cod_titulo)
                        baixar_pdf(link, nome_pdf)
                        boletos.append({
                            "id_boleto": cod_titulo,
                            "cnpj": cnpj,
                            "nome_empresa": nome_empresa,
                            "codigo_cliente_fornecedor": cod_cliente,
                            "codigo_empresa_dominio": codigo_dominio or "NÃO ENCONTRADO",
                            "data_vencimento": conta.get("data_vencimento"),
                            "nome_pdf": nome_pdf
                        })
                    else:
                        print(f"⚠️ PIX sem link: {cod_titulo}")

            time.sleep(0.3)

        if not encontrou_na_data:
            paginas_sem_data += 1
            if paginas_sem_data >= 5:
                print("🛑 Encerrando após 5 páginas sem documentos.")
                break
        else:
            paginas_sem_data = 0

        time.sleep(1)

    # === SALVAR JSON COM BOLETOS E ERROS ===
    resultado = {
        "erros": erros,
        "boletos": boletos
    }
    with open(ARQUIVO_JSON_SAIDA, "w", encoding="utf-8") as f:
        json.dump(resultado, f, ensure_ascii=False, indent=2)
    print(f"\n✅ JSON salvo em: {ARQUIVO_JSON_SAIDA}")

# === EXECUÇÃO ===
if __name__ == "__main__":
    baixar_documentos()
