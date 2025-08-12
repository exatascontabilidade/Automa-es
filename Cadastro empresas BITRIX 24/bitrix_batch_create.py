#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import os
import sys
import time
import json
import typing as t

import requests
import pandas as pd

BITRIX_BASE = "https://grupoexatas.bitrix24.com.br/rest"
WEBHOOK_ADD = "07d5hnaffmd7bsgh"
WEBHOOK_LIST = "pwdu6a6tdwaeynut"

URL_COMPANY_ADD = f"{BITRIX_BASE}/1/{WEBHOOK_ADD}/crm.company.add.json"
URL_COMPANY_LIST = f"{BITRIX_BASE}/1/{WEBHOOK_LIST}/crm.company.list.json"
URL_REQUISITE_ADD = f"{BITRIX_BASE}/1/{WEBHOOK_ADD}/crm.requisite.add.json"

HEADERS = {
    "Content-Type": "application/json",
    "Accept": "application/json",
}

PAGE_SIZE = 50


def bitrix_company_add(title: str, codigo_dominio: str) -> t.Optional[int]:
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
            "UF_CRM_1755003469444": str(codigo_dominio),
        },
        "params": {"REGISTER_SONET_EVENT": "Y"},
    }
    resp = requests.post(URL_COMPANY_ADD, headers=HEADERS, data=json.dumps(payload), timeout=60)
    resp.raise_for_status()
    data = resp.json()
    if "error" in data:
        raise RuntimeError(f"Bitrix error on company.add: {data}")
    return data.get("result")


def bitrix_company_list_last_id() -> int:
    first_payload = {"order": {"ID": "ASC"}, "select": ["ID"], "start": 0}
    resp = requests.post(URL_COMPANY_LIST, headers=HEADERS, data=json.dumps(first_payload), timeout=60)
    resp.raise_for_status()
    data = resp.json()
    if "error" in data:
        raise RuntimeError(f"Bitrix error on company.list[initial]: {data}")
    total = data.get("total")
    if total is None:
        items = data.get("result", [])
        if not items:
            raise RuntimeError("No companies found and 'total' missing in response.")
        return max(int(i["ID"]) for i in items if "ID" in i)

    last_start = ((int(total) - 1) // PAGE_SIZE) * PAGE_SIZE
    last_payload = {"order": {"ID": "ASC"}, "select": ["ID", "TITLE", "DATE_CREATE"], "start": last_start}
    resp2 = requests.post(URL_COMPANY_LIST, headers=HEADERS, data=json.dumps(last_payload), timeout=60)
    resp2.raise_for_status()
    data2 = resp2.json()
    if "error" in data2:
        raise RuntimeError(f"Bitrix error on company.list[last]: {data2}")

    items = data2.get("result", [])
    if not items:
        raise RuntimeError("Last page returned no items.")
    return max(int(i["ID"]) for i in items if "ID" in i)


def bitrix_requisite_add(entity_id: int, cnpj: str) -> int:
    payload = {
        "fields": {
            "ENTITY_TYPE_ID": 4,
            "ENTITY_ID": entity_id,
            "PRESET_ID": 5,
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
    return str(cnpj).strip()


def process_excel(path_excel: str):
    ext = os.path.splitext(path_excel)[1].lower()

    if ext == ".xls":
        try:
            df = pd.read_excel(path_excel, engine="xlrd")
        except ImportError:
            raise ImportError("O formato .xls requer a instalação do pacote 'xlrd'. Use: pip install xlrd")
    else:
        df = pd.read_excel(path_excel)

    required_cols = {"NOME", "CNPJ", "CODIGO_DOMINIO"}
    col_map = {}
    for col in df.columns:
        up = col.upper().strip()
        if up in required_cols:
            col_map[col] = up

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
            last_id = created_id
        else:
            print("company.add não retornou ID; será usado o método de paginação para localizar o último ID.")
            last_id = bitrix_company_list_last_id()

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

        time.sleep(0.5)

    return pd.DataFrame(results)


def main():
    excel_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Relação Empresas - Nome - CNPJ.xls")
    if not os.path.exists(excel_path):
        print(f"Arquivo não encontrado: {excel_path}")
        sys.exit(1)

    try:
        df_result = process_excel(excel_path)
        out_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "resultado_criacao_empresas.csv")
        df_result.to_csv(out_path, index=False, encoding="utf-8-sig")
        print(f"\nConcluído. Resultado salvo em: {out_path}")
    except Exception as e:
        print(f"Erro durante a execução: {e}")
        sys.exit(2)


if __name__ == "__main__":
    main()
