#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import os
import sys
import time
import json
import re
import unicodedata
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

# ---------------------- Normalização de colunas ----------------------
def _strip_accents(s: str) -> str:
    return ''.join(ch for ch in unicodedata.normalize('NFKD', s) if not unicodedata.combining(ch))

def _normalize_key(s: str) -> str:
    s = _strip_accents(str(s)).upper()
    s = re.sub(r'[^A-Z0-9]+', '', s)
    return s

COLUMN_ALIASES = {
    "NOME": "NOME",
    "EMPRESA": "NOME",
    "RAZAOSOCIAL": "NOME",
    "NOMEEMPRESA": "NOME",
    "CLIENTE": "NOME",
    "CNPJ": "CNPJ",
    "CNPJCPF": "CNPJ",
    "DOC": "CNPJ",
    "CODIGODOMINIO": "CODIGO_DOMINIO",
    "COD": "CODIGO_DOMINIO",
    "CODIGO": "CODIGO_DOMINIO",
    "CODIGODOCLIENTE": "CODIGO_DOMINIO",
    "CODCLIENTE": "CODIGO_DOMINIO",
    "CODIGODOMINIOEMPRESA": "CODIGO_DOMINIO",
    "CODDOMINIO": "CODIGO_DOMINIO",
    "CODIGOEMPRESA": "CODIGO_DOMINIO",
}

REQUIRED_TARGET_COLS = {"NOME", "CNPJ", "CODIGO_DOMINIO"}

# ---------------------- Bitrix helpers ----------------------
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

# ---------------------- Normalização de CNPJ ----------------------
def normalize_cnpj_with_flag(cnpj: t.Any) -> t.Tuple[str, bool]:
    """
    Retorna (cnpj_normalizado, houve_normalizacao)
    """
    raw = str(cnpj)
    digits = re.sub(r'\D+', '', raw)
    changed = (raw != digits)
    if len(digits) != 14:
        print(f"[AVISO] CNPJ '{raw}' normalizado para '{digits}', mas não possui 14 dígitos.")
    return digits, changed

# ---------------------- Excel helpers ----------------------
def read_excel_any(path_excel: str) -> pd.DataFrame:
    ext = os.path.splitext(path_excel)[1].lower()
    if ext == ".xls":
        try:
            return pd.read_excel(path_excel, engine="xlrd")
        except ImportError:
            raise ImportError("O formato .xls requer o pacote 'xlrd'. Instale com: pip install xlrd")
    else:
        return pd.read_excel(path_excel)

def build_column_map(df_cols: t.Iterable[str]) -> t.Dict[str, str]:
    col_map: t.Dict[str, str] = {}
    for orig in df_cols:
        norm = _normalize_key(orig)
        if norm in COLUMN_ALIASES:
            target = COLUMN_ALIASES[norm]
            if target not in col_map.values():
                col_map[orig] = target
    return col_map

# ---------------------- Processamento principal ----------------------
def process_excel(path_excel: str) -> t.List[t.Dict[str, str]]:
    df = read_excel_any(path_excel)

    col_map = build_column_map(df.columns)
    missing = REQUIRED_TARGET_COLS - set(col_map.values())
    if missing:
        raise ValueError(f"Planilha deve conter colunas: {REQUIRED_TARGET_COLS}. Encontradas: {df.columns.tolist()}")

    df = df.rename(columns=col_map)

    normalized_only = []

    for idx, row in df.iterrows():
        nome = str(row["NOME"]).strip()
        codigo = str(row["CODIGO_DOMINIO"]).strip()
        cnpj_norm, was_normalized = normalize_cnpj_with_flag(row["CNPJ"])

        if was_normalized:
            normalized_only.append({"NOME": nome, "CODIGO_DOMINIO": codigo, "CNPJ": cnpj_norm})

        print(f"==> Criando empresa: {nome} | Código-Domínio: {codigo}")
        created_id = bitrix_company_add(nome, codigo)
        if not created_id:
            created_id = bitrix_company_list_last_id()

        bitrix_requisite_add(created_id, cnpj_norm)
        time.sleep(0.5)

    return normalized_only

# ---------------------- main ----------------------
def main():
    if len(sys.argv) > 1:
        excel_path = sys.argv[1]
    else:
        excel_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Relação Empresas - Nome - CNPJ.xls")

    if not os.path.exists(excel_path):
        print(f"Arquivo não encontrado: {excel_path}")
        sys.exit(1)

    try:
        normalized_list = process_excel(excel_path)

        base_dir = os.path.dirname(os.path.abspath(__file__))
        out_txt = os.path.join(base_dir, "cnpjs_modificados.txt")

        if normalized_list:
            with open(out_txt, "w", encoding="utf-8") as f:
                f.write("NOME | CODIGO_DOMINIO | CNPJ_NORMALIZADO\n")
                for item in normalized_list:
                    f.write(f"{item['NOME']} | {item['CODIGO_DOMINIO']} | {item['CNPJ']}\n")
            print(f"Arquivo TXT com CNPJs modificados salvo em: {out_txt}")
        else:
            print("Nenhum CNPJ foi modificado; TXT não gerado.")

    except Exception as e:
        print(f"Erro durante a execução: {e}")
        sys.exit(2)

if __name__ == "__main__":
    main()
