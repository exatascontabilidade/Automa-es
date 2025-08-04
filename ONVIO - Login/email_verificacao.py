import imaplib
import email
import re
from email.utils import parsedate_to_datetime

EMAIL = "automacao.gestta@exatascontabilidade.com.br"
SENHA = "Exatas@1010"
IMAP_SERVER = "mail.exatascontabilidade.com.br"
IMAP_PORT = 993
REMETENTE_NOME = "Thomson Reuters"

def extrair_codigo_do_email():
    try:
        mail = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
        mail.login(EMAIL, SENHA)
        mail.select("inbox")

        # Buscar todos os IDs
        status, dados = mail.search(None, "ALL")
        if status != "OK":
            print("Erro na busca de e-mails.")
            return None

        ids = dados[0].split()
        if not ids:
            print("Nenhum e-mail encontrado.")
            return None

        # Pega os últimos 30 e-mails (ou menos, se houver poucos)
        ultimos_ids = ids[-30:]

        # Lista com tuplas (data, id)
        emails_ordenados = []

        for msg_id in ultimos_ids:
            status, dados = mail.fetch(msg_id, "(RFC822)")
            if status != "OK":
                continue

            msg = email.message_from_bytes(dados[0][1])
            data = msg.get("Date")
            try:
                data_convertida = parsedate_to_datetime(data)
                emails_ordenados.append((data_convertida, msg_id))
            except Exception:
                continue

        # Ordena por data de recebimento (do mais novo para o mais antigo)
        emails_ordenados.sort(reverse=True)

        # Procura o e-mail mais recente da Thomson Reuters
        for _, msg_id in emails_ordenados:
            status, dados = mail.fetch(msg_id, "(RFC822)")
            if status != "OK":
                continue

            msg = email.message_from_bytes(dados[0][1])
            remetente = msg["From"]

            if REMETENTE_NOME.lower() not in remetente.lower():
                continue  # pula se não for da Thomson Reuters

            # Tenta extrair o corpo
            corpo = ""
            if msg.is_multipart():
                for parte in msg.walk():
                    if parte.get_content_type() == "text/html":
                        corpo = parte.get_payload(decode=True).decode(errors="ignore")
                        break
            else:
                corpo = msg.get_payload(decode=True).decode(errors="ignore")

            texto = re.sub(r"<[^>]+>", "", corpo)
            texto = texto.replace('\r', '').strip()

            padrao = r"Aqui está seu código de autenticação de dois fatores\s+para Onvio:\s+(\d{6})"
            match = re.search(padrao, texto)

            if match:
                codigo = match.group(1)
                print(f"Código encontrado: {codigo}")
                return codigo
            else:
                print("E-mail da Thomson Reuters localizado, mas código não encontrado.")

        print("Nenhum código da Thomson Reuters encontrado entre os e-mails mais recentes.")
        return None

    except Exception as e:
        print(f"Erro ao buscar código do e-mail: {e}")
        return None
