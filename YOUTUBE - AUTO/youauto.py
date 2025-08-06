import webbrowser
import re

def extrair_id_video(url: str) -> str | None:
    """
    Extrai o ID do vídeo de uma URL do YouTube.

    Args:
        url: A URL do vídeo do YouTube.

    Returns:
        O ID do vídeo se a URL for válida, caso contrário, None.
    """
    # Padrões de regex para os diferentes formatos de URL do YouTube
    padroes = [
        r"(?:v=|\/)([0-9A-Za-z_-]{11}).*",
        r"youtu\.be\/([0-9A-Za-z_-]{11}).*",
        r"youtube\.com\/embed\/([0-9A-Za-z_-]{11}).*"
    ]
    for padrao in padroes:
        match = re.search(padrao, url)
        if match:
            return match.group(1)
    return None

def abrir_video_youtube_em_tela_cheia(url_do_video: str):
    """
    Abre um vídeo do YouTube em modo de cinema (embed).

    Args:
        url_do_video: A URL completa do vídeo do YouTube.
    """
    video_id = extrair_id_video(url_do_video)

    if not video_id:
        print("URL do YouTube inválida.")
        return

    # Constrói a URL de embed para o vídeo
    url_embed = f"https://www.youtube.com/embed/{video_id}?autoplay=1"

    try:
        # Abre a URL de embed no navegador padrão
        webbrowser.open(url_embed)
        print(f"Abrindo o vídeo: {url_embed}")
    except Exception as e:
        print(f"Ocorreu um erro ao tentar abrir o vídeo: {e}")

if __name__ == "__main__":
    # Substitua pela URL do vídeo que você deseja abrir
    url_do_video_exemplo = "https://www.youtube.com/watch?v=tN3qmAzTrX4"
    abrir_video_youtube_em_tela_cheia(url_do_video_exemplo)