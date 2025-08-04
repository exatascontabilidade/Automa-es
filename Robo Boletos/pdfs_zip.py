import os
import zipfile
from pathlib import Path

def extrair_zips_em_lote(pasta_origem, pasta_destino):
    """
    Extrai todos os arquivos .zip da pasta_origem em sequência,
    salvando os conteúdos diretamente em pasta_destino.
    Cada arquivo extraído recebe um prefixo numérico para evitar sobrescrita.
    Se ainda assim houver conflito, um sufixo (1), (2)... é adicionado.
    Após extração, o .zip é removido com segurança.
    """
    os.makedirs(pasta_destino, exist_ok=True)
    arquivos_zip = sorted(Path(pasta_origem).glob("*.zip"), key=lambda f: f.stat().st_mtime)

    if not arquivos_zip:
        print("🚫 Nenhum arquivo .zip encontrado.")
        return

    for idx, zip_path in enumerate(arquivos_zip, start=1):
        print(f"📦 ({idx}/{len(arquivos_zip)}) Extraindo: {zip_path.name}")
        prefixo = f"{idx:02d}_"  # Ex: 01_, 02_, 03_...

        try:
            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                for nome_arquivo in zip_ref.namelist():
                    if not nome_arquivo.endswith('/'):
                        conteudo = zip_ref.read(nome_arquivo)
                        nome_base = Path(nome_arquivo).name
                        nome_completo = prefixo + nome_base
                        destino_final = Path(pasta_destino) / nome_completo

                        # Se o arquivo já existe, cria um nome alternativo com sufixo (1), (2), etc.
                        contador = 1
                        while destino_final.exists():
                            nome_sem_ext = destino_final.stem
                            extensao = destino_final.suffix
                            destino_final = Path(pasta_destino) / f"{nome_sem_ext} ({contador}){extensao}"
                            contador += 1

                        with open(destino_final, 'wb') as f:
                            f.write(conteudo)

            print(f"✅ Extraído com sucesso com prefixo '{prefixo}' em: {pasta_destino}")
            zip_path.unlink()
            print(f"🗑️ Zip deletado: {zip_path.name}\n")

        except zipfile.BadZipFile:
            print(f"❌ Erro: {zip_path.name} está corrompido ou não é um arquivo zip válido. Arquivo não será removido.\n")
        except Exception as e:
            print(f"⚠️ Erro inesperado ao extrair {zip_path.name}: {e}\n")
