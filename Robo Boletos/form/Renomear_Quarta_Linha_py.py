import os
import re
from PyPDF2 import PdfReader

# Diretório onde os arquivos estão localizados
diretorio = r"O:\EXATAS CONTABILIDADE\1 - DPTO FINANCEIRO\AUTOMACAO RECIBOS\recibos_onvio\BOLETOS MANUAIS"

def extrair_nome_pagador(texto):
    # Regex para capturar o conteúdo após "Pagador" e "Nome:"
    match_pagador = re.search(r'Pagador\s+Nome:\s*(.+)', texto, re.DOTALL)
    if match_pagador:
        # Mantém caracteres especiais e remove apenas quebras de linha
        return match_pagador.group(1).strip().split('\n')[0].replace('\n', '').strip()
    return None

def renomear_arquivos():
    for filename in os.listdir(diretorio):
        if filename.endswith(".pdf"):
            file_path = os.path.join(diretorio, filename)
            novo_nome = "Nome_Indefinido"

            try:
                # Ler o conteúdo do PDF
                with open(file_path, "rb") as pdf_file:
                    pdf = PdfReader(pdf_file)
                    texto_pdf = ""
                    for page in pdf.pages:
                        texto_pdf += page.extract_text() + "\n"

                # Tentar extrair o nome do pagador
                nome_extraido = extrair_nome_pagador(texto_pdf)
                if nome_extraido:
                    novo_nome = nome_extraido

                # Definir o novo nome do arquivo
                novo_nome_arquivo = f"{novo_nome}.pdf"
                novo_caminho = os.path.join(diretorio, novo_nome_arquivo)

                # Renomear o arquivo
                os.rename(file_path, novo_caminho)
                print(f"Renomeado: {file_path} -> {novo_caminho}")

            except Exception as e:
                print(f"Erro ao processar o arquivo {file_path}: {e}")

renomear_arquivos()
