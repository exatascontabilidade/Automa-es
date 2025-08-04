import glob
import os
import re
import pandas as pd

def carregar_lista_empresas(temp_path):
    """
    Carrega a lista de empresas do primeiro arquivo Excel (.xls ou .xlsx) encontrado na pasta temp.
    Espera colunas 'Cód.' e 'CNPJ'.
    
    :param temp_path: Caminho da pasta onde está o arquivo Excel.
    :return: Dicionário de empresas {CNPJ: código}.
    """
    arquivos_excel = glob.glob(os.path.join(temp_path, "*.xls")) + \
                     glob.glob(os.path.join(temp_path, "*.xlsx"))

    if not arquivos_excel:
        return {}

    excel_path = arquivos_excel[0]
    extensao = os.path.splitext(excel_path)[-1].lower()

    try:
        # Seleciona o engine com base na extensão do arquivo
        engine = "xlrd" if extensao == ".xls" else "openpyxl"

        df = pd.read_excel(excel_path, engine=engine)

        if "Cód." not in df.columns or "CNPJ" not in df.columns:
            return {}

        df["CNPJ"] = df["CNPJ"].astype(str).str.replace(r'\D', '', regex=True).str.zfill(14)
        df["Cód."] = df["Cód."].astype(str)

        empresas_dict = dict(zip(df["CNPJ"], df["Cód."]))
        return empresas_dict

    except Exception as e:
        print(f"❌ Erro ao ler o Excel: {e}")
        return {}

def extrair_cnpj_do_nome(nome_arquivo):
    """
    Extrai o CNPJ do nome do arquivo, apenas se contiver 'CNPJ_' seguido de 14 dígitos.
    :param nome_arquivo: Nome do arquivo PDF.
    :return: CNPJ em formato numérico ou None.
    """
    match = re.search(r'CNPJ_(\d{14})', nome_arquivo)
    return match.group(1) if match else None

def analisar_pdfs_na_pasta(temp_path, empresas_por_cnpj):
    """
    Analisa os PDFs e identifica os códigos das empresas com base no CNPJ no nome do arquivo.
    Salva um relatório 'comparativos.txt' na mesma pasta.
    :param temp_path: Caminho da pasta onde os PDFs estão.
    :param empresas_por_cnpj: Dicionário {CNPJ: código}.
    :return: Lista única de códigos das empresas com PDFs.
    """
    arquivos_pdf = [
        f for f in os.listdir(temp_path)
        if f.lower().endswith(".pdf")
    ]

    empresas_com_pdf = set()
    arquivos_reconhecidos = []
    arquivos_nao_correspondentes = []
    arquivos_sem_cnpj = []

    for arquivo in arquivos_pdf:
        cnpj = extrair_cnpj_do_nome(arquivo)
        if cnpj:
            if cnpj in empresas_por_cnpj:
                codigo = empresas_por_cnpj[cnpj]
                empresas_com_pdf.add(codigo)
                arquivos_reconhecidos.append(f"{arquivo} → empresa {codigo}")
            else:
                arquivos_nao_correspondentes.append(f"{arquivo} → CNPJ {cnpj} não encontrado na planilha")
        else:
            arquivos_sem_cnpj.append(arquivo)

    # Construção do conteúdo do relatório
    relatorio = []
    relatorio.append(f"📊 RELATÓRIO DE COMPARATIVO - PDF VS CNPJ (pasta temp)")
    relatorio.append(f"Total de PDFs encontrados: {len(arquivos_pdf)}")
    relatorio.append(f"Arquivos reconhecidos (CNPJ válido): {len(arquivos_reconhecidos)}")
    relatorio.append(f"Arquivos com CNPJ não correspondente: {len(arquivos_nao_correspondentes)}")
    relatorio.append(f"Arquivos sem CNPJ: {len(arquivos_sem_cnpj)}\n")

    if arquivos_reconhecidos:
        relatorio.append("✅ Arquivos reconhecidos:")
        relatorio.extend(arquivos_reconhecidos)
        relatorio.append("")

    if arquivos_nao_correspondentes:
        relatorio.append("⚠️ Arquivos com CNPJ não encontrado na planilha:")
        relatorio.extend(arquivos_nao_correspondentes)
        relatorio.append("")

    if arquivos_sem_cnpj:
        relatorio.append("⏩ Arquivos ignorados (sem CNPJ_ no nome):")
        relatorio.extend(arquivos_sem_cnpj)
        relatorio.append("")

    # Salva o relatório
    caminho_relatorio = os.path.join(temp_path, "comparativos.txt")
    with open(caminho_relatorio, "w", encoding="utf-8") as f:
        f.write("\n".join(relatorio))

    return list(empresas_com_pdf)

def obter_codigo_empresa():
    """
    Retorna a lista de códigos das empresas que possuem PDFs na pasta temp com CNPJ válido.
    """
    temp_path = os.path.join(os.path.dirname(__file__), "temp")

    empresas_por_cnpj = carregar_lista_empresas(temp_path)
    if not empresas_por_cnpj:
        return []

    return analisar_pdfs_na_pasta(temp_path, empresas_por_cnpj)
