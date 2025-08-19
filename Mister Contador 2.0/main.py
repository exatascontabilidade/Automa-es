import os
import json
import subprocess
import re
import sys

# --- Constantes de Configuração ---
ARQUIVO_CONTAS = 'contas.json'
ARQUIVO_CONFIG = 'config.json'
PASTA_ESPECIALISTAS = 'sct'

# --- Funções de Gerenciamento do "Banco de Dados" (JSON) ---

def carregar_dados(caminho_arquivo):
    """Carrega dados de um arquivo JSON. Cria o arquivo se não existir."""
    if not os.path.exists(caminho_arquivo):
        # Se for o contas.json, cria uma lista vazia. Se for o config, um objeto.
        dados_iniciais = [] if 'contas' in caminho_arquivo else {"roteamento": []}
        with open(caminho_arquivo, 'w', encoding='utf-8') as f:
            json.dump(dados_iniciais, f, indent=4)
        return dados_iniciais
    try:
        with open(caminho_arquivo, 'r', encoding='utf-8') as f:
            return json.load(f)
    except json.JSONDecodeError:
        print(f"Erro: O arquivo '{caminho_arquivo}' está mal formatado.")
        return None

def salvar_contas(contas):
    """Salva a lista de contas no arquivo JSON."""
    with open(ARQUIVO_CONTAS, 'w', encoding='utf-8') as f:
        json.dump(contas, f, indent=4, ensure_ascii=False)

# --- Funções do Menu ---

def exibir_menu():
    """Mostra as opções do menu principal."""
    print("\n--- Gerenciador de Processamento de Extratos ---")
    print("1. Adicionar Nova Empresa/Conta")
    print("2. Editar ou Excluir uma Empresa/Conta")
    print("3. Processar Arquivo PDF")
    print("S. Sair")
    return input("Escolha uma opção: ").strip().upper()

def listar_e_selecionar_conta(contas, finalidade="selecionar"):
    """Exibe uma lista numerada de contas e retorna a conta selecionada pelo usuário."""
    if not contas:
        print("\nNenhuma conta cadastrada.")
        return None

    print(f"\n--- Contas Cadastradas (para {finalidade}) ---")
    for i, conta in enumerate(contas):
        print(f"{i + 1}. {conta['nome_empresa']} - Banco: {conta['tipo_banco']})")
    
    while True:
        try:
            escolha = input("Digite o número da conta (ou 0 para cancelar): ")
            indice = int(escolha) - 1
            if indice == -1:
                return None
            if 0 <= indice < len(contas):
                return contas[indice]
            else:
                print("Número inválido. Tente novamente.")
        except ValueError:
            print("Entrada inválida. Por favor, digite um número.")

def adicionar_empresa():
    """Adiciona uma nova conta ao banco de dados."""
    print("\n--- Adicionar Nova Conta ---")
    contas = carregar_dados(ARQUIVO_CONTAS)
    
    nova_conta = {
        "nome_empresa": input("Nome da Empresa: "),
        "agencia": input("Número da Agência (ex: 1234-5): "),
        "conta": input("Número da Conta (ex: 98765-4): "),
        "tipo_banco": input("Tipo do Banco (ex: banco_a, santander_pf): ").lower()
    }
    
    contas.append(nova_conta)
    salvar_contas(contas)
    print(f"\nEmpresa '{nova_conta['nome_empresa']}' adicionada com sucesso!")

def editar_excluir_empresa():
    """Permite editar ou excluir uma conta existente."""
    contas = carregar_dados(ARQUIVO_CONTAS)
    conta_selecionada = listar_e_selecionar_conta(contas, finalidade="editar/excluir")

    if not conta_selecionada:
        return

    print("\nO que você deseja fazer?")
    print("[E] Editar")
    print("[D] Deletar")
    print("[C] Cancelar")
    acao = input("Escolha uma ação: ").strip().upper()

    if acao == 'D':
        contas.remove(conta_selecionada)
        salvar_contas(contas)
        print(f"\nConta de '{conta_selecionada['nome_empresa']}' deletada com sucesso.")
    elif acao == 'E':
        print("\n--- Editando Conta (deixe em branco para não alterar) ---")
        
        # Encontra o índice da conta para poder alterá-la na lista
        indice = contas.index(conta_selecionada)

        novo_nome = input(f"Nome da Empresa ({conta_selecionada['nome_empresa']}): ")
        if novo_nome: contas[indice]['nome_empresa'] = novo_nome

        nova_agencia = input(f"Agência ({conta_selecionada['agencia']}): ")
        if nova_agencia: contas[indice]['agencia'] = nova_agencia

        nova_conta = input(f"Conta ({conta_selecionada['conta']}): ")
        if nova_conta: contas[indice]['conta'] = nova_conta

        novo_tipo = input(f"Tipo do Banco ({conta_selecionada['tipo_banco']}): ")
        if novo_tipo: contas[indice]['tipo_banco'] = novo_tipo.lower()

        salvar_contas(contas)
        print("\nConta atualizada com sucesso.")
    else:
        print("\nAção cancelada.")

def processar_arquivo():
    """Processa um arquivo PDF para uma conta selecionada."""
    contas = carregar_dados(ARQUIVO_CONTAS)
    config = carregar_dados(ARQUIVO_CONFIG)

    conta_selecionada = listar_e_selecionar_conta(contas, finalidade="processar")
    if not conta_selecionada:
        return

    # Obter caminho do PDF
    caminho_pdf = input("\nCole o caminho para o arquivo PDF e pressione Enter: ").strip().strip('"\'')
    if not os.path.exists(caminho_pdf):
        print("Erro: Arquivo PDF não encontrado.")
        return

    # Encontrar o script especialista
    script_especialista = None
    for roteamento in config.get('roteamento', []):
        if roteamento['tipo_banco'] == conta_selecionada['tipo_banco']:
            script_especialista = roteamento['script']
            break

    if not script_especialista:
        print(f"Erro: Nenhum script especialista encontrado para o tipo de banco '{conta_selecionada['tipo_banco']}'.")
        print("Verifique seu arquivo config.json.")
        return

    script_path = os.path.join(PASTA_ESPECIALISTAS, script_especialista)
    if not os.path.exists(script_path):
        print(f"Erro: O script '{script_especialista}' não foi encontrado na pasta '{PASTA_ESPECIALISTAS}'.")
        return

    print(f"\n--- Processando PDF... ---")
    
    try:
        resultado = subprocess.run(
            [sys.executable, script_path, caminho_pdf], 
            capture_output=True, text=True, encoding='utf-8', errors='replace', check=True
        )
        
        # Imprime a saída completa do especialista
        print("\n--- Resultado ---")
        print(resultado.stdout)

        # --- VALIDAÇÃO FINAL ---
        agencia_pdf_match = re.search(r"Agencia:\s*(.*)", resultado.stdout)
        conta_pdf_match = re.search(r"Conta:\s*(.*)", resultado.stdout)

        if agencia_pdf_match and conta_pdf_match:
            agencia_pdf = agencia_pdf_match.group(1).strip()
            conta_pdf = conta_pdf_match.group(1).strip()

            print("\n--- Verificação de Consistência ---")
            if (agencia_pdf == conta_selecionada['agencia'] and conta_pdf == conta_selecionada['conta']):
                print(" SUCESSO: A conta processada no PDF corresponde à conta selecionada no menu.")
            else:
                print(" ATENÇÃO: A conta processada no PDF não corresponde à conta selecionada!")
                print(f"  - Conta Selecionada: Ag {conta_selecionada['agencia']}, CC {conta_selecionada['conta']}")
                print(f"  - Conta no PDF:      Ag {agencia_pdf}, CC {conta_pdf}")
        else:
            print("\nAVISO: Não foi possível encontrar as linhas 'Agencia/Conta Processada' na saída do especialista para validação.")

    except subprocess.CalledProcessError as e:
        print(f"\nERRO: O script especialista '{script_especialista}' falhou.")
        print("--- Saída de Erro (stderr) ---")
        print(e.stderr)

# --- Loop Principal ---
def main():
    while True:
        escolha = exibir_menu()
        if escolha == '1':
            adicionar_empresa()
        elif escolha == '2':
            editar_excluir_empresa()
        elif escolha == '3':
            processar_arquivo()
        elif escolha == 'S':
            print("Saindo...")
            break
        else:
            print("Opção inválida, tente novamente.")

if __name__ == "__main__":
    main()