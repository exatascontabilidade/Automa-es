import os
import re
import collections
import json
from pypdf import PdfReader

try:
    from pdf2image import convert_from_path
    import pytesseract
    OCR_DISPONIVEL = True
except ImportError:
    OCR_DISPONIVEL = False

def obter_diretorio_download():
    """Garante que a pasta 'temp' exista e retorna seu caminho."""
    diretorio_atual = os.path.dirname(os.path.abspath(__file__))
    diretorio_download = os.path.join(diretorio_atual, "temp")

    if not os.path.exists(diretorio_download):
        os.makedirs(diretorio_download)

    return diretorio_download

def gerar_prefixo():
    """Gera um prefixo no formato 'MM-AAAA - ' a partir do arquivo JSON."""
    diretorio_download = obter_diretorio_download()
    caminho_arquivo = os.path.join(diretorio_download, "mes_ano.json")

    if not os.path.exists(caminho_arquivo):
        return "INDEFINIDO - "

    try:
        with open(caminho_arquivo, "r", encoding="utf-8") as file:
            dados = json.load(file)
            ano_mes = dados.get("ano_mes", "INDEFINIDO")

        if ano_mes == "INDEFINIDO":
            raise ValueError("❌ Mês e ano não encontrados no arquivo JSON.")

        mes, ano = ano_mes.split("/")
        return f"{mes}-{ano} - "

    except Exception:
        return "INDEFINIDO - "

def extrair_texto_pdf(caminho_pdf):
    """Extrai texto de um PDF."""
    try:
        with open(caminho_pdf, "rb") as pdf_file:
            pdf = PdfReader(pdf_file)
            texto_pdf = "\n".join([page.extract_text() for page in pdf.pages if page.extract_text()])
        return texto_pdf

    except Exception as e:
        print(f"❌ Erro ao ler o PDF: {e}")
        return None

def identificar_modelo(texto):
    """Identifica o modelo do boleto com base na estrutura do texto."""
    if "Nome:" in texto and "CPF/CNPJ:" in texto and "Beneficiário" in texto and "Pagador" in texto:
        return "modelo_beneficiario_pagador"
    elif "Nome:" in texto and "CPF/CNPJ:" in texto:
        return "modelo_com_nome_cpf"
    elif "Recibo do Pagador" in texto:
        return "modelo_recibo"
    elif "Pagamento via Pix" in texto and "CNPJ do Pagador" in texto and "Pagador" in texto:
        return "modelo_pix_exatas"
    elif "Pagador" in texto:
        return "modelo_simples"
    else:
        return "desconhecido"

def extrair_nome_e_cnpj(texto):
    """Extrai o nome do pagador e o CNPJ com base no modelo identificado."""
    modelo = identificar_modelo(texto)
    linhas = texto.split("\n")
    nome_pagador = None
    cnpj = None

    if modelo == "modelo_beneficiario_pagador":
        match_pagador = re.search(r'Pagador\s+Nome:\s*(.+?)\n', texto)
        match_cnpj = re.search(r'Pagador.*?CPF\/CNPJ:\s*([\d./-]+)', texto, re.DOTALL)

        if match_pagador:
            nome_pagador = match_pagador.group(1).strip()
        if match_cnpj:
            cnpj = match_cnpj.group(1).strip()

    elif modelo == "modelo_com_nome_cpf":
        for i, linha in enumerate(linhas):
            if "Nome:" in linha:
                nome_pagador = linha.split("Nome:")[-1].strip()
            if "CPF/CNPJ:" in linha:
                cnpj = linha.split("CPF/CNPJ:")[-1].strip()
                
    elif modelo == "modelo_pix_exatas":
        try:
            linhas = texto.splitlines()
            for linha in linhas:
                if "CNPJ do Pagador" in linha:
                    match_cnpj = re.search(r"\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}", linha)
                    if match_cnpj:
                        cnpj = re.sub(r"[^\d]", "", match_cnpj.group())

                if "Pagador" in linha:
                    nome_pagador = linha.replace("Pagador", "").strip()

        except Exception as e:
            print(f"❌ Erro ao extrair dados do modelo_pix_exatas: {e}")
            
    elif modelo == "modelo_recibo" or modelo == "modelo_simples":
        for i, linha in enumerate(linhas):
            if "Pagador" in linha and i + 1 < len(linhas):
                nome_pagador = linhas[i + 1].strip()
            if "CPF/CNPJ" in linha:
                cnpj = linha.split("CPF/CNPJ:")[-1].strip()

    # Limpeza dos dados
    if nome_pagador:
        nome_pagador = re.sub(r'[\/:*?"<>|]', '', nome_pagador).strip()
        nome_pagador = re.sub(r'\s+', ' ', nome_pagador)

    if cnpj:
        cnpj = re.sub(r'[^\d]', '', cnpj).strip()

    if nome_pagador and cnpj:
        return f"{nome_pagador} - CNPJ_{cnpj}"
    elif nome_pagador:
        return nome_pagador
    elif cnpj:
        return cnpj
    else:
        print(f"⚠️ Nenhum nome de pagador encontrado ({modelo}). Verifique o layout do PDF.")
        return None



def extrair_vencimento(texto):
    """Tenta encontrar a data de vencimento no texto do boleto. Se não encontrar, imprime o texto para análise."""
    padroes = [
        r'vencimento em (\d{2}/\d{2}/\d{4})',
        r'data de vencimento\s*[:\-]?\s*(\d{2}/\d{2}/\d{4})',
        r'vencimento\s*[:\-]?\s*(\d{2}/\d{2}/\d{4})',
        r'vencimento\s*\n\s*(\d{2}/\d{2}/\d{4})',
        r'Pagável em qualquer banco\s+(\d{2}/\d{2}/\d{4})',  # mesma linha          
        r'Pagável em qualquer banco\s*\n\s*(\d{2}/\d{2}/\d{4})',
        r'Pagamento via Pix\s+(\d{2}/\d{2}/\d{4})',  # mesma linha          
        r'Pagamento via Pix\s*\n\s*(\d{2}/\d{2}/\d{4})', # linha seguinte
        r'(\d{4}-\d{2}-\d{2})'
    ]

    for padrao in padroes:
        match = re.search(padrao, texto, re.IGNORECASE | re.MULTILINE)
        if match:
            return match.group(1).replace("/", "-")

    # Verifica se a linha abaixo da palavra "vencimento" é uma data
    linhas = texto.splitlines()
    for i, linha in enumerate(linhas):
        if "vencimento" in linha.lower():
            if i + 1 < len(linhas):
                proxima_linha = linhas[i + 1].strip()
                data_match = re.match(r"\d{2}/\d{2}/\d{4}", proxima_linha)
                if data_match:
                    return proxima_linha.replace("/", "-")

    return "SEM-VENCIMENTO"


def contar_boletos_por_empresa():
    """Conta quantos boletos existem para cada empresa na pasta."""
    DIRETORIO = obter_diretorio_download()
    contador = collections.defaultdict(list)

    for filename in sorted(os.listdir(DIRETORIO)):
        if filename.endswith(".pdf"):
            file_path = os.path.join(DIRETORIO, filename)
            try:
                texto_pdf = extrair_texto_pdf(file_path)

                if texto_pdf:
                    nome_extraido = extrair_nome_e_cnpj(texto_pdf)

                    if nome_extraido:
                        contador[nome_extraido].append(file_path)

            except Exception as e:
                print(f"❌ Erro ao processar {filename}: {e}")

    return contador

def renomear_arquivos():
    """Renomeia os arquivos PDFs corretamente, incluindo vencimento no nome."""
    DIRETORIO = obter_diretorio_download()
    PREFIXO = gerar_prefixo()
    contador_boletos = contar_boletos_por_empresa()
    total_renomeados = 0

    for empresa, arquivos in contador_boletos.items():
        qtd_boletos = len(arquivos)

        for idx, file_path in enumerate(arquivos, start=1):
            try:
                # 🔍 Extrai texto e vencimento do boleto
                texto_pdf = extrair_texto_pdf(file_path)
                vencimento = extrair_vencimento(texto_pdf) if texto_pdf else "SEM-VENCIMENTO"
                vencimento = vencimento.replace("/", "-")  # evita / no nome do arquivo

                # 📄 Gera o novo nome do arquivo com vencimento no final
                if qtd_boletos > 1:
                    novo_nome = f"BOLETO MES {PREFIXO}({idx} de {qtd_boletos}) - {empresa} - VENC_{vencimento}.pdf"
                else:
                    novo_nome = f"BOLETO MES {PREFIXO}{empresa} - VENC_{vencimento}.pdf"

                novo_caminho = os.path.join(DIRETORIO, novo_nome)

                if os.path.exists(novo_caminho):
                    print(f"⚠️ Arquivo já existe: {novo_nome}. Pulando...")
                    continue

                os.rename(file_path, novo_caminho)
                total_renomeados += 1
                print(f"✅ Renomeado: {file_path} → {novo_caminho}")

            except Exception as e:
                print(f"❌ Erro ao processar o arquivo {file_path}: {e}")

    print(f"\n🚀 Todos os arquivos foram renomeados com sucesso! Total: {total_renomeados} arquivos.")
