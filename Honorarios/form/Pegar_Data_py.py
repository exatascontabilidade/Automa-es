import os
import sys
from PyPDF2 import PdfReader
import pyperclip

def extrair_data_vencimento(pdf_path):
    with open(pdf_path, 'rb') as file:
        reader = PdfReader(file)
        text = ''
        for page in reader.pages:
            text += page.extract_text() or ''  # Adiciona um fallback caso não extraia texto
        linhas = text.split('\n')
        for i, linha in enumerate(linhas):
            if 'vencimento' in linha.lower():
                if i + 1 < len(linhas):
                    return ''.join(filter(str.isdigit, linhas[i + 1]))
        return None

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Uso: python Pegar_Data_py.py <caminho_do_pdf>")
        sys.exit(1)

    pdf_path = sys.argv[1]
    data_vencimento = extrair_data_vencimento(pdf_path)
    if data_vencimento:
        pyperclip.copy(data_vencimento)
        print(f"Data de vencimento copiada para a área de transferência: {data_vencimento}")
    else:
        print("Não foi possível encontrar a data de vencimento no arquivo PDF.")
