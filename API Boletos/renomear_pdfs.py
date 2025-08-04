import os
import pandas as pd
from datetime import datetime
import re

def obter_diretorio_base():
    """Retorna o diretório onde o script está localizado."""
    return os.path.dirname(os.path.abspath(__file__))

def obter_diretorio_download():
    """Garante que a pasta 'temp' exista e retorna seu caminho."""
    diretorio_download = os.path.join(obter_diretorio_base(), "temp")
    if not os.path.exists(diretorio_download):
        os.makedirs(diretorio_download)
    return diretorio_download

def localizar_planilha_empresas():
    """Procura pela planilha na pasta 'temp'."""
    base = obter_diretorio_download()
    for nome in ["Relação Empresas - Nome - CNPJ.xls", "Relação Empresas - Nome - CNPJ.csv"]:
        caminho = os.path.join(base, nome)
        if os.path.exists(caminho):
            return caminho
    raise FileNotFoundError("❌ Nenhuma planilha chamada 'Relação Empresas - Nome - CNPJ.xls' ou 'Relação Empresas - Nome - CNPJ.csv' foi encontrada na pasta 'temp'.")

def carregar_planilha_empresas(caminho_planilha):
    """Lê a planilha e retorna um dicionário {cnpj: nome_empresa}, com verificação de colunas e CNPJs."""
    if caminho_planilha.endswith('.xls') or caminho_planilha.endswith('.xlsx'):
        df = pd.read_excel(caminho_planilha)
    else:
        df = pd.read_csv(caminho_planilha, encoding='latin1')

    df.columns = [col.lower().strip() for col in df.columns]
    print("🔎 Colunas encontradas na planilha:", df.columns)

    if 'empresa' not in df.columns or 'cnpj' not in df.columns:
        raise ValueError("❌ A planilha deve conter colunas chamadas 'Empresa' e 'CNPJ'.")

    dicionario = {}
    for _, row in df.iterrows():
        cnpj = re.sub(r'\D', '', str(row['cnpj'])).zfill(14)  # Garante 14 dígitos
        nome_completo = str(row['empresa']).strip()
        nome_sem_codigo = re.sub(r'^\d{2}\.\d{3}\.\d{3}\s+', '', nome_completo)
        nome_limpo = re.sub(r'[\/:*?"<>|]', '', nome_sem_codigo).strip()
        dicionario[cnpj] = nome_limpo

    print(f"✅ Total de empresas carregadas: {len(dicionario)}")
    print("🧾 Exemplos de CNPJs carregados:", list(dicionario.keys())[:5])
    return dicionario

def gerar_nome_disponivel(diretorio, base_nome):
    """Gera um nome de arquivo disponível com contador se necessário."""
    contador = 1
    nome_final = f"{base_nome}.pdf"
    caminho_final = os.path.join(diretorio, nome_final)

    while os.path.exists(caminho_final):
        contador += 1
        nome_final = f"{base_nome} ({contador}).pdf"
        caminho_final = os.path.join(diretorio, nome_final)

    return caminho_final, nome_final

def renomear_arquivos_por_cnpj():
    diretorio = obter_diretorio_download()
    planilha_path = localizar_planilha_empresas()
    empresas = carregar_planilha_empresas(planilha_path)
    data_execucao = datetime.today().strftime('%d-%m-%Y')
    total = 0

    for filename in os.listdir(diretorio):
        if filename.endswith('.pdf'):
            nome_base_limpo = re.sub(r'\(.*?\)', '', filename)  # remove sufixos como (2)
            cnpj_match = re.search(r'\d{11,14}', nome_base_limpo)
            if cnpj_match:
                cnpj = cnpj_match.group().zfill(14)
                nome_empresa = empresas.get(cnpj)
                if nome_empresa:
                    nome_limpo = re.sub(r'[\/:*?"<>|]', '', nome_empresa).strip()
                    base_nome = f"{cnpj}_{data_execucao}_{nome_limpo}"
                    caminho_antigo = os.path.join(diretorio, filename)
                    caminho_novo, nome_final = gerar_nome_disponivel(diretorio, base_nome)

                    os.rename(caminho_antigo, caminho_novo)
                    print(f"✅ Renomeado: {filename} → {nome_final}")
                    total += 1
                else:
                    print(f"❌ CNPJ {cnpj} não encontrado na planilha.")
            else:
                print(f"❌ CNPJ não identificado no nome do arquivo: {filename}")
    
    print(f"\n🚀 Renomeação finalizada! Total renomeados: {total}")
