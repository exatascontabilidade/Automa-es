import os
import sys
import zipfile
import json
import xml.etree.ElementTree as ET
import re
from datetime import datetime
from oletools.olevba import VBA_Parser

# ==============================================================================
# FUNÇÃO 1: Descompactação Completa (Extração Bruta)
# ==============================================================================
def descompactar_arquivo_completo(caminho_arquivo, diretorio_saida):
    """Descompacta o arquivo Office em sua estrutura de pastas e arquivos brutos."""
    print("\n--- Iniciando Extração Bruta da Estrutura ---")
    try:
        with zipfile.ZipFile(caminho_arquivo, 'r') as zip_ref:
            zip_ref.extractall(diretorio_saida)
        print(f"[+] Estrutura completa do arquivo extraída para: {diretorio_saida}")
        return True
    except Exception as e:
        print(f"[!] Erro na descompactação bruta: {e}")
        return False

# ==============================================================================
# FUNÇÃO 2: Extração de Código VBA
# ==============================================================================
def extrair_codigo_vba(caminho_arquivo, diretorio_saida):
    """Extrai código VBA usando a análise do olevba."""
    print("\n--- Iniciando Extração de Código VBA ---")
    try:
        vba_parser = VBA_Parser(caminho_arquivo)
        if not vba_parser.detect_vba_macros():
            print("[-] Nenhum macro VBA detectado.")
            return

        print("[+] Macros VBA detectadas. Extraindo...")
        arquivos_salvos = 0
        for (filename, stream_path, vba_filename, vba_code) in vba_parser.extract_macros():
            try:
                codigo_decodificado = vba_code.decode('utf-8', errors='replace')
                if codigo_decodificado.strip():
                    caminho_completo_saida = os.path.join(diretorio_saida, f"{vba_filename}.vba")
                    with open(caminho_completo_saida, 'w', encoding='utf-8') as f:
                        f.write(codigo_decodificado)
                    arquivos_salvos += 1
            except Exception:
                continue
        print(f"[+] {arquivos_salvos} módulo(s) VBA extraído(s) para: {diretorio_saida}")
    except Exception as e:
        print(f"[!] Erro na extração de VBA: {e}")
    finally:
        if 'vba_parser' in locals() and vba_parser:
            vba_parser.close()

# ==============================================================================
# FUNÇÃO 3: Extração de JSON Formatado
# ==============================================================================
def extrair_dados_json(caminho_arquivo, diretorio_saida):
    """Extrai e formata JSON de 'Custom XML Parts'."""
    print("\n--- Iniciando Extração de Dados JSON ---")
    arquivos_extraidos = 0
    try:
        with zipfile.ZipFile(caminho_arquivo, 'r') as zip_ref:
            partes_xml = [f for f in zip_ref.namelist() if f.startswith('customXml/item')]
            if not partes_xml:
                print("[-] Nenhum dado JSON (Custom XML Part) encontrado.")
                return

            print(f"[+] {len(partes_xml)} parte(s) de dados encontradas. Processando...")
            for i, caminho_parte in enumerate(partes_xml, 1):
                json_string = None
                try:
                    conteudo_bytes = zip_ref.read(caminho_parte)
                    try:
                        root = ET.fromstring(conteudo_bytes)
                        for element in root.iter():
                            if element.text and element.text.strip().startswith('{'):
                                json_string = element.text
                                break
                    except ET.ParseError:
                        raw_text = conteudo_bytes.decode('utf-8', errors='ignore')
                        match = re.search(r'<!\[CDATA\[(.*)\]\]>', raw_text, re.DOTALL)
                        if match:
                            json_string = match.group(1)
                    
                    if json_string and json_string.strip():
                        dados_json = json.loads(json_string)
                        json_formatado = json.dumps(dados_json, indent=4, ensure_ascii=False)
                        caminho_completo_saida = os.path.join(diretorio_saida, f"dados_extraidos_{i}.json")
                        with open(caminho_completo_saida, 'w', encoding='utf-8') as f_out:
                            f_out.write(json_formatado)
                        arquivos_extraidos += 1
                except Exception:
                    continue
            print(f"[+] {arquivos_extraidos} arquivo(s) JSON extraído(s) para: {diretorio_saida}")
    except Exception as e:
        print(f"[!] Erro na extração de JSON: {e}")

# ==============================================================================
# FUNÇÃO PRINCIPAL (ORQUESTRADOR)
# ==============================================================================
def extrator_forense(caminho_arquivo_excel):
    if not os.path.exists(caminho_arquivo_excel):
        print(f"Erro: Arquivo não encontrado -> {caminho_arquivo_excel}")
        return

    # --- Criação dos diretórios de saída ---
    script_dir = os.path.dirname(os.path.abspath(sys.argv[0])) if '__file__' in locals() else os.getcwd()
    nome_base = os.path.splitext(os.path.basename(caminho_arquivo_excel))[0]
    timestamp = datetime.now().strftime("%Y-%m-%d_%Hh%Mm%Ss")
    
    dir_principal_saida = os.path.join(script_dir, f"Extracao_Forense_{nome_base}_{timestamp}")
    dir_raw = os.path.join(dir_principal_saida, "_ESTRUTURA_BRUTA")
    dir_vba = os.path.join(dir_principal_saida, "_CODIGO_VBA")
    dir_json = os.path.join(dir_principal_saida, "_DADOS_JSON")

    print(f"[*] Iniciando Análise Forense do arquivo: {os.path.basename(caminho_arquivo_excel)}")
    print(f"[*] Os resultados serão salvos em: {dir_principal_saida}")
    os.makedirs(dir_principal_saida, exist_ok=True)
    os.makedirs(dir_raw, exist_ok=True)
    os.makedirs(dir_vba, exist_ok=True)
    os.makedirs(dir_json, exist_ok=True)
    
    # --- Execução das extrações ---
    descompactar_arquivo_completo(caminho_arquivo_excel, dir_raw)
    extrair_codigo_vba(caminho_arquivo_excel, dir_vba)
    extrair_dados_json(caminho_arquivo_excel, dir_json)

    print("\n[✓] Análise Forense Concluída!")

# --- Bloco de Execução ---
if __name__ == "__main__":
    try:
        caminho_arquivo = input("Insira o caminho completo para o arquivo Excel a ser desmontado: ")
        extrator_forense(caminho_arquivo.strip().strip('"'))
    except KeyboardInterrupt:
        print("\nOperação cancelada.")