import pandas as pd
import os
import re
from openpyxl import Workbook
import unidecode

def limpar_nome(texto):
    texto = unidecode.unidecode(texto.upper())  # Remove acentos e coloca em maiúsculas
    texto = re.sub(r'[^A-Z0-9 ]', '', texto)  # Remove pontuação
    texto = re.sub(r'\b(SA|S/A|LTDA|ME|EPP)\b', '', texto)  # Remove sufixos comuns
    texto = re.sub(r'\s+', '_', texto.strip())  # Espaços para underline
    return texto

def salvar_bloco(cliente_nome, linhas, pasta_saida):
    nome_arquivo = f"Retencao_{cliente_nome}.xlsx"
    caminho_saida = os.path.join(pasta_saida, nome_arquivo)

    wb = Workbook()
    ws = wb.active
    for linha in linhas:
        ws.append(linha)
    wb.save(caminho_saida)

def adicionar_bloco_acumulado(dicionario, cliente_nome, linhas):
    if cliente_nome not in dicionario:
        dicionario[cliente_nome] = []
    dicionario[cliente_nome].extend(linhas)

def separar_blocos_por_nome_normalizado(caminho_csv):
    df = pd.read_csv(caminho_csv, header=None, encoding="ISO-8859-1", sep=';')
    blocos_por_nome = {}
    i = 0

    while i < len(df):
        linha = df.iloc[i].astype(str).fillna("").tolist()

        if any("CLIENTE:" in str(col).upper() for col in linha):
            bloco_linhas = [df.iloc[i].tolist()]

            try:
                nome_raw = linha[11].strip()
                cliente_chave = limpar_nome(nome_raw)
            except IndexError:
                nome_raw = "CLIENTE_DESCONHECIDO"
                cliente_chave = limpar_nome(nome_raw)

            i += 1
            if i < len(df):
                bloco_linhas.append(df.iloc[i].tolist())
                i += 1

            while i < len(df):
                linha_atual = df.iloc[i].astype(str).fillna("").tolist()
                if any("CLIENTE:" in str(col).upper() for col in linha_atual):
                    break
                bloco_linhas.append(df.iloc[i].tolist())
                i += 1

            adicionar_bloco_acumulado(blocos_por_nome, cliente_chave, bloco_linhas)
        else:
            i += 1

    pasta_saida = "Retencoes_Separadas"
    os.makedirs(pasta_saida, exist_ok=True)

    for cliente, linhas in blocos_por_nome.items():
        salvar_bloco(cliente, linhas, pasta_saida)

    print(f"✅ {len(blocos_por_nome)} arquivos gerados agrupando nomes semelhantes em: {pasta_saida}/")

# Executar
separar_blocos_por_nome_normalizado("Retenções a Compensar.csv")
