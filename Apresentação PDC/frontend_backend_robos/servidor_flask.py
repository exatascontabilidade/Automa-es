from flask import Flask, jsonify, request
from flask_cors import CORS
import requests
import base64
import os
from dotenv import load_dotenv

load_dotenv()

app = Flask(__name__)
CORS(app)  # ✅ Permite requisições de qualquer origem

GITHUB_TOKEN = os.getenv("GITHUB_TOKEN")
REPO = "exatascontabilidade/automacao-remota"
ARQUIVO = "comando.json"
API_URL = f"https://api.github.com/repos/{REPO}/contents/{ARQUIVO}"

HEADERS = {
    "Authorization": f"Bearer {GITHUB_TOKEN}",
    "Accept": "application/vnd.github.v3+json"
}

@app.route("/executar", methods=["POST"])
def executar_comando():
    try:
        res = requests.get(API_URL, headers=HEADERS)
        if res.status_code != 200:
            return jsonify(success=False, message="Erro ao obter SHA."), 500

        sha = res.json().get("sha")
        novo_conteudo = base64.b64encode(b'{"executar": true}').decode("utf-8")

        payload = {
            "message": "🚀 Acionar robô",
            "content": novo_conteudo,
            "sha": sha
        }

        res_update = requests.put(API_URL, headers=HEADERS, json=payload)
        if res_update.status_code in [200, 201]:
            return jsonify(success=True)
        else:
            return jsonify(success=False, message=res_update.text), 500

    except Exception as e:
        return jsonify(success=False, message=str(e)), 500

if __name__ == "__main__":
    app.run(debug=True)
