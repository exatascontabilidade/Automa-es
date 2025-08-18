import pdfplumber
import re
import os
import json
import argparse
from typing import List, Dict, Tuple, Optional
from unidecode import unidecode

# As funções de `definir_tipo_por_cor` e `encontrar_cor_do_texto` permanecem as mesmas.
def definir_tipo_por_cor(cor: tuple) -> str:
    if not cor or len(cor) < 3: return "Padrão"
    r, g, b = cor[:3]
    if r > 0.5 and g < 0.4: return "Débito"
    if g > 0.5 and r < 0.4: return "Crédito"
    return "Padrão"

def encontrar_cor_do_texto(pagina, texto_procurado: str) -> Optional[tuple]:
    resultados = pagina.search(texto_procurado, case=False)
    if not resultados: return None
    bbox = (resultados[0]['x0'], resultados[0]['top'], resultados[0]['x1'], resultados[0]['bottom'])
    chars_na_caixa = [c for c in pagina.chars if c['x0'] >= bbox[0] and c['x1'] <= bbox[2] and c['top'] >= bbox[1] and c['bottom'] <= bbox[3]]
    return chars_na_caixa[0].get('non_stroking_color') if chars_na_caixa else None

def carregar_config_efi() -> Dict:
    """Carrega a configuração específica para o banco Efí."""
    try:
        # Garante que o arquivo de configuração seja encontrado na mesma pasta do script
        try:
            script_dir = os.path.dirname(os.path.abspath(__file__))
        except NameError:
            script_dir = os.getcwd()
        # Usa o nome 'configuracoes.json' para manter a consistência com o roteador
        caminho_config = os.path.join(script_dir, 'config.json')

        with open(caminho_config, 'r', encoding='utf-8') as f:
            config_total = json.load(f)
            for banco_info in config_total.get('bancos', []):
                if banco_info.get('nome_banco') == 'efi':
                    print("Configuração do modelo de banco 'EFI' carregada.")
                    return banco_info.get('config_extracao', {})
    except Exception as e:
        print(f"Erro ao carregar configuração do EFI: {e}")
    return {}

def extrair_dados_metodo_hibrido(pagina_obj, config_extracao: Dict) -> Tuple[List, List]:
    """
    Extrai os dados e retorna uma lista com 4 elementos:
    [data, descricao, tipo, valor]
    """
    config = config_extracao['config_hibrido']
    palavra_chave_inicio = config['palavra_chave_inicio']
    padrao_data = re.compile(config['padrao_data_transacao'])
    lancamentos_extraidos = []
    
    todas_as_palavras = pagina_obj.extract_words(x_tolerance=2, y_tolerance=2)
    pos_inicio = 0
    for i, palavra in enumerate(todas_as_palavras):
        if palavra_chave_inicio in palavra['text'].lower():
            pos_inicio = i + 1
            break
            
    palavras_relevantes = todas_as_palavras[pos_inicio:]
    blocos_de_transacao = []
    bloco_atual = []
    for palavra in palavras_relevantes:
        if padrao_data.match(palavra['text']):
            if bloco_atual: blocos_de_transacao.append(bloco_atual)
            bloco_atual = [palavra]
        elif bloco_atual:
            bloco_atual.append(palavra)
    if bloco_atual: blocos_de_transacao.append(bloco_atual)

    for bloco in blocos_de_transacao:
        if not bloco or len(bloco) < 2: continue
        
        data_str = bloco[0]['text']
        palavras_do_bloco_sem_data = bloco[1:]
        
        valor_str = None
        valor_index = -1

        # Lógica para encontrar o valor
        for i in range(len(palavras_do_bloco_sem_data) - 1, -1, -1):
            palavra_candidata = palavras_do_bloco_sem_data[i]['text']
            try:
                float(palavra_candidata.replace('.', '').replace(',', '.'))
                valor_str = palavra_candidata
                valor_index = i
                break
            except ValueError:
                continue

        if valor_str is None:
            continue

        # --- LÓGICA ATUALIZADA PARA DESCRIÇÃO COMPLETA ---
        # A descrição são TODAS as palavras do bloco, exceto a data (já removida) e o próprio valor.
        # Isso captura textos que aparecem antes e DEPOIS do valor.
        descricao_palavras = []
        for i, palavra_obj in enumerate(palavras_do_bloco_sem_data):
            if i == valor_index:
                continue # Pula a palavra que foi identificada como o valor
            descricao_palavras.append(palavra_obj['text'])
        descricao = ' '.join(descricao_palavras)
        # --- FIM DA LÓGICA ATUALIZADA ---

        # Remove qualquer sequência de 8 ou mais dígitos da descrição
        descricao = re.sub(r'\b\d{8,}\b', '', descricao)
        # Limpa espaços duplos que possam ter sido criados
        descricao = ' '.join(descricao.split())

        cor_valor = encontrar_cor_do_texto(pagina_obj, valor_str)
        tipo_transacao = definir_tipo_por_cor(cor_valor) if cor_valor else "Padrão"
        
        if tipo_transacao == "Débito":
            tipo_completo = "DÉBITO"
        elif tipo_transacao == "Crédito":
            tipo_completo = "CRÉDITO"
        else:
            tipo_completo = "" 

        lancamentos_extraidos.append([data_str, descricao, tipo_completo, valor_str])
        
    return lancamentos_extraidos, []

def formatar_para_txt_final(lista_lancamentos: list) -> str:
    """Formata os lançamentos em uma tabela de 4 colunas, sem truncar a descrição."""
    if not lista_lancamentos: return "Nenhum lançamento para formatar."
    
    tabela_para_formatar = [['Data', 'Descrição', 'Tipo', 'Valor']] + lista_lancamentos
    larguras_colunas = [0, 0, 0, 0]
    
    for linha in tabela_para_formatar:
        try:
            larguras_colunas[0] = max(larguras_colunas[0], len(linha[0]))
            larguras_colunas[1] = max(larguras_colunas[1], len(linha[1]))
            larguras_colunas[2] = max(larguras_colunas[2], len(linha[2]))
            larguras_colunas[3] = max(larguras_colunas[3], len(linha[3]))
        except IndexError:
            continue
    
    # LÓGICA DE TRUNCAMENTO REMOVIDA PARA MOSTRAR A DESCRIÇÃO COMPLETA

    string_final = ""
    cabecalho = tabela_para_formatar[0]
    string_final += f"{cabecalho[0].center(larguras_colunas[0])} | {cabecalho[1].center(larguras_colunas[1])} | {cabecalho[2].center(larguras_colunas[2])} | {cabecalho[3].center(larguras_colunas[3])}\n"
    string_final += f"{'-' * larguras_colunas[0]}-+-{'-' * larguras_colunas[1]}-+-{'-' * larguras_colunas[2]}-+-{'-' * larguras_colunas[3]}\n"
    
    for linha in tabela_para_formatar[1:]:
        data, desc, tipo, valor = linha
        # Usa a descrição completa (desc) em vez de uma versão truncada
        string_final += f"{data.ljust(larguras_colunas[0])} | {desc.ljust(larguras_colunas[1])} | {tipo.center(larguras_colunas[2])} | {valor.rjust(larguras_colunas[3])}\n"
        
    return string_final

def main():
    """Função principal do especialista Efí."""
    parser = argparse.ArgumentParser(description="Processador para o modelo extratos do banco EFI.")
    parser.add_argument("caminho_pdf", help="O caminho para o arquivo PDF do EFI.")
    args = parser.parse_args()
    
    config_extracao = carregar_config_efi()
    if not config_extracao:
        print("Nao foi possivel carregar a configuracao para o especialista EFI.")
        return

    todos_lancamentos = []
    try:
        with pdfplumber.open(args.caminho_pdf) as pdf:
            texto_primeira_pagina = pdf.pages[0].extract_text()
            padrao_cnpj = config_extracao.get('padrao_cnpj', '')
            ocorrencia = config_extracao.get('ocorrencia_cnpj', 1)
            matches_cnpj = re.findall(padrao_cnpj, texto_primeira_pagina) if padrao_cnpj else []
            cnpj = matches_cnpj[ocorrencia - 1] if len(matches_cnpj) >= ocorrencia else "Nao_Encontrado"
            
            for i, pagina_obj in enumerate(pdf.pages):
                lancamentos, _ = extrair_dados_metodo_hibrido(pagina_obj, config_extracao)
                todos_lancamentos.extend(lancamentos)
        
        if todos_lancamentos:
            nome_banco_fmt = "efi"
            cnpj_fmt = cnpj.replace('/', '-').replace('.', '')
            
            script_dir = os.path.dirname(os.path.abspath(__file__))
            pasta_processamento = os.path.join(script_dir, "Processamento")
            pasta_banco = os.path.join(pasta_processamento, nome_banco_fmt.upper())

            os.makedirs(pasta_banco, exist_ok=True)
            
            nome_arquivo = f"Extrato_{nome_banco_fmt}_{cnpj_fmt}.txt"
            caminho_arquivo_saida = os.path.join(pasta_banco, nome_arquivo)

            texto_formatado = formatar_para_txt_final(todos_lancamentos)
            with open(caminho_arquivo_saida, 'w', encoding='utf-8') as f:
                f.write(f"Extrato para o CNPJ: {cnpj}\nBanco: {nome_banco_fmt.upper()}\n{'='*40}\n\n{texto_formatado}")
            
            print(f"Processamento finalizado. {len(todos_lancamentos)} lançamentos salvos em:")
            print(f"'{caminho_arquivo_saida}'")
        else:
            print("Nenhum lançamento válido foi extraído pelo processamento do modelo de banco EFI")
            
    except Exception as e:
        print(f"Ocorreu um erro no processamento do modelo de banco EFI: {e}")

if __name__ == "__main__":
    main()
