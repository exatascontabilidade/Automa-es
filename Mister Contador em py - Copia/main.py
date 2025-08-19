import pdfplumber
import json
import subprocess
import os
import sys
from unidecode import unidecode

def carregar_configuracao_principal(caminho_json: str) -> dict:
    """Carrega a configuração principal que mapeia bancos a scripts."""
    try:
        with open(caminho_json, 'r', encoding='utf-8') as f:
            return json.load(f)
    except FileNotFoundError:
        print(f"Erro: Arquivo de configuração principal '{caminho_json}' não encontrado.")
        return {}
    except json.JSONDecodeError:
        print(f"Erro: Arquivo '{caminho_json}' contém um formato JSON inválido.")
        return {}

def normalizar_texto(texto: str) -> str:
    """Converte o texto para minúsculas e remove acentos."""
    return unidecode(texto.lower())

def identificar_banco_e_obter_script(texto_pdf: str, config: dict) -> str | None:
    """
    Identifica o banco com base no texto do PDF e retorna o nome do script
    especialista a ser chamado.
    """
    print("Iniciando identificação do banco para roteamento...")
    texto_pdf_normalizado = normalizar_texto(texto_pdf)
    
    for banco_info in config.get('bancos', []):
        for frase in banco_info.get('palavras_identificadoras', []):
            if normalizar_texto(frase) in texto_pdf_normalizado:
                script_especialista = banco_info.get('script')
                print(f"Banco '{banco_info['nome_banco']}' identificado. Roteando para o script: '{script_especialista}'")
                return script_especialista
                
    print("Nenhum banco correspondente encontrado na configuração principal.")
    return None

def main():
    """
    Função principal que solicita o caminho de um PDF, identifica o banco e
    chama o script de processamento apropriado.
    """
    caminho_do_arquivo_pdf = input("Por favor, cole ou digite o caminho para o arquivo PDF e pressione Enter: ")
    caminho_do_arquivo_pdf = caminho_do_arquivo_pdf.strip().strip('"').strip("'")

    if not os.path.exists(caminho_do_arquivo_pdf):
        print(f"Erro: O arquivo PDF '{caminho_do_arquivo_pdf}' não foi encontrado.")
        return

    # Constrói o caminho para o arquivo 'config.json' na mesma pasta do script.
    try:
        # __file__ é o caminho para o script atual
        script_dir = os.path.dirname(os.path.abspath(__file__))
    except NameError:
        # Fallback para o diretório de trabalho atual se __file__ não estiver definido
        script_dir = os.getcwd()
        
    caminho_config = os.path.join(script_dir, 'config.json')
    config_principal = carregar_configuracao_principal(caminho_config)

    if not config_principal:
        return

    try:
        with pdfplumber.open(caminho_do_arquivo_pdf) as pdf:
            primeira_pagina = pdf.pages[0]
            texto_primeira_pagina = primeira_pagina.extract_text()
            if not texto_primeira_pagina:
                print("Erro: Não foi possível extrair texto da primeira página do PDF.")
                return

    except Exception as e:
        print(f"Erro ao ler o arquivo PDF: {e}")
        return

    # --- INÍCIO DO BLOCO CORRIGIDO ---

    # 1. Identifica qual script deve ser chamado com base no texto do PDF
    script_a_chamar = identificar_banco_e_obter_script(texto_primeira_pagina, config_principal)

    # 2. Verifica se um script foi de fato identificado
    if script_a_chamar:
        # 3. Constrói o caminho COMPLETO e CORRETO para o script na subpasta 'sct'
        caminho_script_especialista = os.path.join(script_dir, 'sct', script_a_chamar)
        
        # 4. Verifica se o arquivo do script realmente existe no caminho esperado
        if not os.path.exists(caminho_script_especialista):
            print(f"Erro Crítico: O script especialista '{script_a_chamar}' foi encontrado na configuração, mas o arquivo não existe no diretório 'sct'.")
            print(f"Caminho verificado: {caminho_script_especialista}")
            return # Para a execução
            
        print(f"\n--- Executando especialista: {script_a_chamar} ---")
        
        try:
            # 5. Executa o script especialista
            resultado = subprocess.run(
                ['python', caminho_script_especialista, caminho_do_arquivo_pdf], 
                capture_output=True, 
                text=True, 
                encoding='utf-8', 
                errors='replace',
                check=True # Lança uma exceção se o script terminar com erro
            )
            
            # Imprime a saída do script especialista se tudo correu bem
            print("\n--- Saída do Especialista ---")
            print(resultado.stdout)
            print("--- Fim da execução do especialista ---")

        except subprocess.CalledProcessError as e:
            # Captura erros que ocorreram DENTRO do script especialista
            print(f"\nERRO: O script especialista '{script_a_chamar}' falhou durante a execução.")
            print(f"Código de Saída: {e.returncode}")
            if e.stdout:
                print("\n--- Saída Padrão (stdout) do erro ---")
                print(e.stdout)
            if e.stderr:
                print("\n--- Saída de Erro (stderr) do especialista ---")
                print(e.stderr)
    else:
        # Mensagem para o caso de nenhum banco ser identificado no PDF
        print("Não foi possível determinar um script especialista para processar este PDF.")

    # --- FIM DO BLOCO CORRIGIDO ---


if __name__ == "__main__":
    main()