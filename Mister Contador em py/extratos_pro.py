import pdfplumber
import re
import os
import json
from typing import Union, List, Dict, Tuple, Optional
from unidecode import unidecode # Importa a nova biblioteca

def carregar_configuracao(caminho_json: str) -> List[Dict]:
    """Carrega as configurações dos bancos a partir de um arquivo JSON."""
    try:
        with open(caminho_json, 'r', encoding='utf-8') as f:
            config = json.load(f)
        print("Arquivo de configuração 'config.json' carregado com sucesso.")
        return config.get('bancos', [])
    except FileNotFoundError:
        print("Erro: Arquivo 'config.json' não encontrado.")
        return []
    except json.JSONDecodeError:
        print("Erro: Arquivo 'config.json' contém um formato JSON inválido.")
        return []

def normalizar_texto(texto: str) -> str:
    """Converte o texto para minúsculas e remove acentos."""
    return unidecode(texto.lower())

def identificar_banco(texto_pdf: str, configs_bancos: List[Dict]) -> Optional[Dict]:
    """
    (Versão Definitiva) Identifica o banco normalizando os textos antes de comparar.
    """
    print("\nIniciando identificação do banco...")
    texto_pdf_normalizado = normalizar_texto(texto_pdf)
    
    for config in configs_bancos:
        for frase_identificadora in config['palavras_identificadoras']:
            # Normaliza também a frase da configuração para uma comparação justa
            frase_normalizada = normalizar_texto(frase_identificadora)
            if frase_normalizada in texto_pdf_normalizado:
                print(f"Banco identificado com base na frase '{frase_identificadora}': {config['nome_banco'].upper()}")
                return config
                
    print("Nenhum banco correspondente encontrado na configuração.")
    return None

def extrair_dados_banco(texto_completo: str, config_banco: Dict) -> Tuple[str, List]:
    """Processa o texto de um extrato usando as regras específicas do banco identificado."""
    config_extracao = config_banco['config_extracao']
    
    # Extrai CNPJ
    cnpj_encontrado = "Não encontrado"
    matches_cnpj = re.findall(config_extracao['padrao_cnpj'], texto_completo)
    ocorrencia = config_extracao.get('ocorrencia_cnpj', 1)
    if len(matches_cnpj) >= ocorrencia:
        cnpj_encontrado = matches_cnpj[ocorrencia - 1]

    # Isola o bloco de transações
    palavra_chave_inicio = config_extracao['palavra_chave_inicio_tabela']
    inicio_lancamentos = texto_completo.find(palavra_chave_inicio)
    if inicio_lancamentos == -1:
        print(f"Aviso: Palavra-chave de início '{palavra_chave_inicio}' não encontrada.")
        return cnpj_encontrado, []
    
    bloco_transacoes = texto_completo[inicio_lancamentos:]
    
    # Encontra datas como âncoras
    padrao_data = re.compile(config_extracao['padrao_data'])
    matches_data = list(padrao_data.finditer(bloco_transacoes))
    
    lancamentos_extraidos = []
    print(f"Encontradas {len(matches_data)} possíveis datas. Analisando...")

    for i, match_atual in enumerate(matches_data):
        data_str = match_atual.group(0)
        inicio_bloco = match_atual.end()
        fim_bloco = matches_data[i + 1].start() if (i + 1) < len(matches_data) else len(bloco_transacoes)
        
        texto_do_bloco = bloco_transacoes[inicio_bloco:fim_bloco]
        texto_do_bloco = texto_do_bloco.replace(palavra_chave_inicio, '')
        
        numeros_no_bloco = re.findall(r'[\d\.,-]+\b', texto_do_bloco)
        if not numeros_no_bloco: continue
        
        valor_str = numeros_no_bloco[-1]
        try:
            float(valor_str.replace('.', '').replace(',', '.'))
        except ValueError: continue

        texto_sem_numeros = re.sub(r'[\d\.,-]+', ' ', texto_do_bloco)
        descricao = ' '.join(texto_sem_numeros.split()).strip().replace("•", "")

        if not descricao or ("descrição" in descricao.lower() and "protocolo" in descricao.lower()): continue

        lancamentos_extraidos.append([data_str, descricao, valor_str])
        
    return cnpj_encontrado, lancamentos_extraidos

def formatar_para_txt_final(lista_lancamentos: list) -> str:
    """Formata uma lista de lançamentos em uma única string de texto bem alinhada."""
    tabela_para_formatar = [['Data', 'Descrição', 'Valor']] + lista_lancamentos
    larguras_colunas = [0, 0, 0]
    for linha in tabela_para_formatar:
        if len(linha[0]) > larguras_colunas[0]: larguras_colunas[0] = len(linha[0])
        if len(linha[1]) > larguras_colunas[1]: larguras_colunas[1] = len(linha[1])
        if len(linha[2]) > larguras_colunas[2]: larguras_colunas[2] = len(linha[2])
    
    if larguras_colunas[1] > 80:
        larguras_colunas[1] = 80

    string_final = ""
    cabecalho = tabela_para_formatar[0]
    string_final += f"{cabecalho[0].center(larguras_colunas[0])} | {cabecalho[1].center(larguras_colunas[1])} | {cabecalho[2].center(larguras_colunas[2])}\n"
    string_final += f"{'-' * larguras_colunas[0]}-+-{'-' * larguras_colunas[1]}-+-{'-' * larguras_colunas[2]}\n"
    
    for linha in tabela_para_formatar[1:]:
        data, desc, valor = linha
        desc_truncada = (desc[:larguras_colunas[1]-3] + '...') if len(desc) > larguras_colunas[1] else desc
        string_final += f"{data.ljust(larguras_colunas[0])} | {desc_truncada.ljust(larguras_colunas[1])} | {valor.rjust(larguras_colunas[2])}\n"
        
    return string_final

def main():
    """Função principal que orquestra todo o fluxo de extração."""
    caminho_do_arquivo_pdf = "58ed5fa9478df5f9af11344250202c2a.pdf"
    
    configs_bancos = carregar_configuracao('config.json')
    if not configs_bancos:
        return
        
    texto_pdf_completo = ""
    print(f"Lendo arquivo: {caminho_do_arquivo_pdf}")
    try:
        with pdfplumber.open(caminho_do_arquivo_pdf) as pdf:
            for pagina in pdf.pages:
                texto_pagina = pagina.extract_text()
                if texto_pagina:
                    texto_pdf_completo += " " + texto_pagina
    except Exception as e:
        print(f"Erro ao abrir ou ler o PDF: {e}")
        return

    config_banco_identificado = identificar_banco(texto_pdf_completo, configs_bancos)
    
    if config_banco_identificado:
        cnpj, lancamentos = extrair_dados_banco(texto_pdf_completo, config_banco_identificado)
        
        if lancamentos:
            texto_formatado = formatar_para_txt_final(lancamentos)
            nome_arquivo_saida = f"Extrato_{config_banco_identificado['nome_banco']}_{cnpj.replace('/', '-')}.txt"
            
            with open(nome_arquivo_saida, 'w', encoding='utf-8') as f:
                f.write(f"Extrato para o CNPJ: {cnpj}\n")
                f.write(f"Banco: {config_banco_identificado['nome_banco'].upper()}\n")
                f.write("=" * 40 + "\n\n")
                f.write(texto_formatado)

            print("\n--- DADOS GERAIS ---")
            print(f"CNPJ do Titular: {cnpj}")
            print("\nProcesso concluído!")
            print(f"{len(lancamentos)} lançamentos salvos em '{nome_arquivo_saida}'")
        else:
            print("Dados de lançamentos não puderam ser extraídos apesar de o banco ter sido identificado.")

if __name__ == "__main__":
    main()