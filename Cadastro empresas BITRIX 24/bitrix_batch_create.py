#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import os
import sys
import time
import json
import math
import typing as t

import requests
import pandas as pd

BITRIX_BASE = "https://grupoexatas.bitrix24.com.br/rest"
# Webhooks (adjust if needed)
WEBHOOK_ADD = "07d5hnaffmd7bsgh"       # used for company.add and requisite.add
WEBHOOK_LIST = "pwdu6a6tdwaeynut"      # used for company.list (as provided in your example)

# Endpoints
URL_COMPANY_ADD = f"{BITRIX_BASE}/1/{WEBHOOK_ADD}/crm.company.add.json"
URL_COMPANY_LIST = f"{BITRIX_BASE}/1/{WEBHOOK_LIST}/crm.company.list.json"
URL_REQUISITE_ADD = f"{BITRIX_BASE}/1/{WEBHOOK_ADD}/crm.requisite.add.json"

HEADERS = {
    "Content-Type": "application/json",
    "Accept": "application/json",
}

PAGE_SIZE = 50  # Bitrix default

def bitrix_company_add(title: str, codigo_dominio: str) -> t.Optional[int]:
    """Create a company and return its ID (if Bitrix returns it)."""
    payload = {
        "fields": {
            "TITLE": title,
            "COMPANY_TYPE": "CUSTOMER",
            "INDUSTRY": "IT",
            "EMPLOYEES": "EMPLOYEES_1",
            "CURRENCY_ID": "BRL",
            "OPENED": "Y",
            "ASSIGNED_BY_ID": 1,
            "PHONE": [{"VALUE": "555888", "VALUE_TYPE": "WORK"}],
            # Campo customizado -> Código - Domínio
            "UF_CRM_1755003469444": str(codigo_dominio),
        },
        "params": {"REGISTER_SONET_EVENT": "Y"},
    }
    resp = requests.post(URL_COMPANY_ADD, headers=HEADERS, data=json.dumps(payload), timeout=60)
    resp.raise_for_status()
    data = resp.json()
    if "error" in data:
        raise RuntimeError(f"Bitrix error on company.add: {data}")
    # Usually Bitrix returns the new ID in 'result'
    return data.get("result")

def bitrix_company_list_last_id() -> int:
    """Use pagination to find the ID of the most recently created company.
    Logic:
      1) First request to get 'total'
      2) Compute last page start: ((total-1)//PAGE_SIZE)*PAGE_SIZE
      3) Request that page and select last ID (max) from it
    """
    # Step 1: initial call to get total
    first_payload = {
        "order": {"ID": "ASC"},
        "select": ["ID"],
        "start": 0
    }
    resp = requests.post(URL_COMPANY_LIST, headers=HEADERS, data=json.dumps(first_payload), timeout=60)
    resp.raise_for_status()
    data = resp.json()
    if "error" in data:
        raise RuntimeError(f"Bitrix error on company.list[initial]: {data}")
    total = data.get("total")
    if total is None:
        # Try fallback if 'total' is missing
        items = data.get("result", [])
        if not items:
            raise RuntimeError("No companies found and 'total' missing in response.")
        last_id = max(int(i["ID"]) for i in items if "ID" in i)
        return last_id

    last_start = ((int(total) - 1) // PAGE_SIZE) * PAGE_SIZE

    # Step 2: request the last page
    last_payload = {
        "order": {"ID": "ASC"},
        "select": ["ID", "TITLE", "DATE_CREATE"],
        "start": last_start
    }
    resp2 = requests.post(URL_COMPANY_LIST, headers=HEADERS, data=json.dumps(last_payload), timeout=60)
    resp2.raise_for_status()
    data2 = resp2.json()
    if "error" in data2:
        raise RuntimeError(f"Bitrix error on company.list[last]: {data2}")

    items = data2.get("result", [])
    if not items:
        raise RuntimeError("Last page returned no items.")

    # Pick the highest ID from the last page
    last_id = max(int(i["ID"]) for i in items if "ID" in i)
    return last_id

def bitrix_requisite_add(entity_id: int, cnpj: str) -> int:
    """Create requisite with CNPJ for the given company ID. Returns new requisite ID."""
    payload = {
        "fields": {
            "ENTITY_TYPE_ID": 4,          # 4 = Company
            "ENTITY_ID": entity_id,
            "PRESET_ID": 5,               # Provided by you
            "NAME": "Pessoa jurídica",
            "RQ_CNPJ": cnpj
        }
    }
    resp = requests.post(URL_REQUISITE_ADD, headers=HEADERS, data=json.dumps(payload), timeout=60)
    resp.raise_for_status()
    data = resp.json()
    if "error" in data:
        raise RuntimeError(f"Bitrix error on requisite.add: {data}")
    return data.get("result")

def normalize_cnpj(cnpj: str) -> str:
    """Return CNPJ as provided; adjust here if you want to force a specific mask/digits-only."""
    return str(cnpj).strip()

def process_excel(path_excel: str):
    df = pd.read_excel(path_excel)
    # Expect columns: NOME, CNPJ, CODIGO_DOMINIO
    required_cols = {"NOME", "CNPJ", "CODIGO_DOMINIO"}
    missing = required_cols - set(map(str.upper, df.columns))
    # Attempt case-insensitive remap
    col_map = {}
    for col in df.columns:
        up = col.upper()
        if up in required_cols:
            col_map[col] = up
    if missing:
        # Remap succeeded?
        if set(col_map.values()) != required_cols:
            raise ValueError(f"Planilha deve conter colunas: {required_cols}. Encontradas: {df.columns.tolist()}")
        df = df.rename(columns=col_map)

    results = []
    for _, row in df.iterrows():
        nome = str(row["NOME"]).strip()
        cnpj = normalize_cnpj(row["CNPJ"])
        codigo = str(row["CODIGO_DOMINIO"]).strip()

        print(f"==> Criando empresa: {nome} | Código-Domínio: {codigo}")
        created_id = bitrix_company_add(nome, codigo)
        if created_id:
            print(f"ID retornado por company.add: {created_id}")
        else:
            print("company.add não retornou ID; será usado o método de paginação para localizar o último ID.")

        # Conforme solicitado, usar a listagem para encontrar a última criada
        last_id = bitrix_company_list_last_id()
        print(f"Última empresa (ID) identificada via paginação: {last_id}")

        # Cadastra o CNPJ nos requisitos
        req_id = bitrix_requisite_add(last_id, cnpj)
        print(f"Requisite criado (ID): {req_id} para empresa ID: {last_id}")

        results.append({
            "NOME": nome,
            "CNPJ": cnpj,
            "CODIGO_DOMINIO": codigo,
            "COMPANY_ID_LIST_LAST": last_id,
            "REQUISITE_ID": req_id,
            "COMPANY_ID_RETURNED": created_id
        })

        # Pausa curta para evitar rate limit
        time.sleep(0.5)

    return pd.DataFrame(results)

def main():
    if len(sys.argv) < 2:
        print("Uso: python bitrix_batch_create.py <caminho_para_planilha.xlsx>")
        sys.exit(1)
    excel_path = sys.argv[1]
    if not os.path.exists(excel_path):
        print(f"Arquivo não encontrado: {excel_path}")
        sys.exit(1)

    try:
        df_result = process_excel(excel_path)
        out_path = "resultado_criacao_empresas.csv"
        df_result.to_csv(out_path, index=False, encoding="utf-8-sig")
        print(f"\nConcluído. Resultado salvo em: {out_path}")
    except Exception as e:
        print(f"Erro durante a execução: {e}")
        sys.exit(2)

if __name__ == "__main__":
    main()
