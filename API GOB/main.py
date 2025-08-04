import os
import re
import requests
import json

# === CONFIGURAÇÕES GERAIS ===
HEADERS = {
    "X-Api-Key": "2cb72bfa43fbd2ccf98e059a159c7dad"
}
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMP_DIR = os.path.join(BASE_DIR, "temp")
os.makedirs(TEMP_DIR, exist_ok=True)

PAGE_SIZE = 1000

# === DEFINIÇÃO DOS TIPOS DE PARCELAMENTO ===
TIPOS_PARCELAMENTO = {
    "ParcelamentoFederalSimplificado": {
        "url_empresas": "https://integracao.gob.ctributaria.com.br/api/v1/ParcelamentoFederalSimplificado",
        "url_parcelas": "https://integracao.gob.ctributaria.com.br/api/v1/ParcelamentoFederalSimplificadoParcelas?where[0][type]:=isNotNull&where[0][attribute]=darfId&where[1][type]=equals&where[1][attribute]=situacao&where[1][value]=A vencer",
        "id_empresa": "id",
        "id_parcelamento_ref": "parcelamentoFederalSimplificadoId",
        "id_arquivo": "darfId",
        "nome_arquivo": "darfName",
        "nome_pasta": "FEDERAL_SIMPLIFICADO"
    },
    "ParcelamentoPgfn": {
        "url_empresas": "https://integracao.gob.ctributaria.com.br/api/v1/ParcelamentoPgfn",
        "url_parcelas": "https://integracao.gob.ctributaria.com.br/api/v1/ParcelamentoPgfnParcelas?where[0][type]:=isNotNull&where[0][attribute]=darfId&where[1][type]=equals&where[1][attribute]=pago&where[1][value]=false",
        "id_empresa": "id",
        "id_parcelamento_ref": "parcelamentoPgfnId",
        "id_arquivo": "darfId",
        "nome_arquivo": "darfName",
        "nome_pasta": "PGFN"
    },
    "ParcelamentoPrevidenciario": {
        "url_empresas": "https://integracao.gob.ctributaria.com.br/api/v1/ParcelamentoPrevidenciario",
        "url_parcelas": "https://integracao.gob.ctributaria.com.br/api/v1/ParcelamentoPrevidenciarioParcelas?where[0][type]:=isNotNull&where[0][attribute]=gpsId&where[1][type]=equals&where[1][attribute]=situacaoParcela&where[1][value]=Devedora",
        "id_empresa": "id",
        "id_parcelamento_ref": "parcelamentoPrevidenciarioId",
        "id_arquivo": "gpsId",
        "nome_arquivo": "gpsName",
        "nome_pasta": "PREVIDENCIARIO"
    },
    "ParcelamentoSimplesNacional": {
        "url_empresas": "https://integracao.gob.ctributaria.com.br/api/v1/ParcelamentoSimplesNacional",
        "id_empresa": "id",
        "id_parcelamento_ref": "parcelamentoSimplesNacionalId",
        "id_arquivo": "darfId",
        "nome_arquivo": "darfName",
        "nome_pasta": "SIMPLES_NACIONAL",
        "personalizado": True
    },
    "ParcelamentoNaoPrevidenciario": {
        "url_empresas": "https://integracao.gob.ctributaria.com.br/api/v1/ParcelamentoNaoPrevidenciario",
        "url_parcelas": "https://integracao.gob.ctributaria.com.br/api/v1/ParcelamentoNaoPrevidenciarioParcelas?where[0][type]:=isNotNull&where[0][attribute]=darfId&where[1][type]=equals&where[1][attribute]=situacao&where[1][value]=Em aberto",
        "id_empresa": "id",
        "id_parcelamento_ref": "parcelamentoNaoPrevidenciarioTributosId",
        "id_arquivo": "darfId",
        "nome_arquivo": "darfName",
        "nome_pasta": "NAO_PREVIDENCIARIO"
    }
}

def paginar_requisicao(url_base):
    offset = 0
    resultados = []
    while True:
        paginada = f"{url_base}?maxSize={PAGE_SIZE}&offset={offset}" if "?" not in url_base else f"{url_base}&maxSize={PAGE_SIZE}&offset={offset}"
        resp = requests.get(paginada, headers=HEADERS)
        resp.raise_for_status()
        lista = resp.json().get("list", [])
        if not lista:
            break
        resultados.extend(lista)
        offset += PAGE_SIZE
    return resultados

def baixar_arquivo(config, empresa, arquivo_id, nome_arquivo, registros_baixados, registros_com_erros):
    nome = empresa["nome"]
    cnpj = empresa["cnpj"]
    pasta_empresa = os.path.join(TEMP_DIR, f"{nome} - {cnpj}")
    pasta_final = os.path.join(pasta_empresa, config["nome_pasta"])
    os.makedirs(pasta_final, exist_ok=True)
    caminho_arquivo = os.path.join(pasta_final, nome_arquivo)

    try:
        url_download = f"https://integracao.gob.ctributaria.com.br/api/v1/Attachment/file/{arquivo_id}"
        for tentativa in range(3):
            try:
                r = requests.get(url_download, headers=HEADERS, timeout=20)
                if r.status_code == 404:
                    raise FileNotFoundError(f"Arquivo não encontrado (404) para ID: {arquivo_id}")
                r.raise_for_status()
                break
            except (requests.exceptions.ConnectionError, requests.exceptions.Timeout) as err:
                if tentativa == 2:
                    raise err
                print(f"🔁 Tentando novamente ({tentativa + 1}/3)...")

        with open(caminho_arquivo, "wb") as f:
            f.write(r.content)
        registros_baixados.append({"empresa": nome, "cnpj": cnpj, "tipo_parcelamento": config["nome_pasta"], "nome_arquivo": nome_arquivo, "arquivo_id": arquivo_id, "pasta": pasta_final})
    except Exception as e:
        registros_com_erros.append({"empresa": nome, "cnpj": cnpj, "tipo_parcelamento": config["nome_pasta"], "nome_arquivo": nome_arquivo, "arquivo_id": arquivo_id, "pasta": pasta_final, "erro": str(e)})

def processar_simples_nacional(config, registros_baixados, registros_com_erros):
    empresas_raw = paginar_requisicao(config["url_empresas"])
    empresas = {}
    ids_validos = []

    for e in empresas_raw:
        situacao = e.get("situacao", "")
        if situacao in ["Em parcelamento", "Aguardando Pagamento da 1ª Parcela"]:
            empresa_id = e.get(config["id_empresa"])
            nome = re.sub(r'[\\/:*?"<>|]', '', e.get("accountName", "Empresa_Desconhecida")).strip()
            cnpj = e.get("cnpj", "00.000.000/0000-00").strip()
            empresas[empresa_id] = {"nome": nome, "cnpj": cnpj}
            ids_validos.append(empresa_id)

    for id_parcelamento in ids_validos:
        url = (
            "https://integracao.gob.ctributaria.com.br/api/v1/ParcelamentoSimplesNacionalParcelas"
            f"?where[0][type]:=isNotNull&where[0][attribute]=darfId"
            f"&where[1][type]=isNull&where[1][attribute]=valorPago"
            f"&where[2][type]=equals&where[2][attribute]=parcelamentoSimplesNacionalId"
            f"&where[2][value]={id_parcelamento}"
        )
        parcelas = paginar_requisicao(url)

        for p in parcelas:
            arquivo_id = p.get(config["id_arquivo"])
            nome_arquivo = p.get(config["nome_arquivo"], f"{arquivo_id}.pdf")
            empresa = empresas.get(id_parcelamento)
            if arquivo_id and empresa:
                baixar_arquivo(config, empresa, arquivo_id, nome_arquivo, registros_baixados, registros_com_erros)

def processar_parcelamento(tipo, config, registros_baixados, registros_com_erros):
    if config.get("personalizado"):
        return processar_simples_nacional(config, registros_baixados, registros_com_erros)

    empresas_raw = paginar_requisicao(config["url_empresas"])
    empresas = {}
    for e in empresas_raw:
        empresa_id = e.get(config["id_empresa"])
        nome = re.sub(r'[\\/:*?"<>|]', '', e.get("accountName", "Empresa_Desconhecida")).strip()
        cnpj = e.get("cnpj", "00.000.000/0000-00").strip()
        empresas[empresa_id] = {"nome": nome, "cnpj": cnpj}

    parcelas = paginar_requisicao(config["url_parcelas"])
    for p in parcelas:
        empresa_id = p.get(config["id_parcelamento_ref"])
        arquivo_id = p.get(config["id_arquivo"])
        nome_arquivo = p.get(config["nome_arquivo"], f"{arquivo_id}.pdf")
        empresa = empresas.get(empresa_id)
        if arquivo_id and empresa:
            baixar_arquivo(config, empresa, arquivo_id, nome_arquivo, registros_baixados, registros_com_erros)

# Execução principal
registros_baixados = []
registros_com_erros = []

for tipo, config in TIPOS_PARCELAMENTO.items():
    print(f"\n🔄 Processando: {tipo}...")
    processar_parcelamento(tipo, config, registros_baixados, registros_com_erros)

# Salvando arquivos de log
with open(os.path.join(BASE_DIR, "parcelamentos_baixados.json"), "w", encoding="utf-8") as f:
    json.dump(registros_baixados, f, indent=4, ensure_ascii=False)

with open(os.path.join(BASE_DIR, "parcelamentos_erros.json"), "w", encoding="utf-8") as f:
    json.dump(registros_com_erros, f, indent=4, ensure_ascii=False)

print(f"\n✅ JSON de baixados: {len(registros_baixados)} registros")
print(f"⚠️ JSON de erros: {len(registros_com_erros)} registros")
