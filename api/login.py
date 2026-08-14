# -*- coding: utf-8 -*-
"""
POST /api/login — { "usuario": "...", "senha": "..." } -> { token, usuario, papel }

Credenciais NUNCA ficam no código — só o hash SHA-256 da senha, na variável
de ambiente USUARIOS_PAINEL. Veja gerar_credenciais.py pra criar as suas.
"""
import os
import time
import hmac
import base64
import hashlib
import json
from http.server import BaseHTTPRequestHandler

PAPEIS_VALIDOS = {"admin", "elite", "sniper", "olympus"}
TOKEN_HORAS = 12
AUTH_SECRET = (os.environ.get("AUTH_SECRET", "") or "troque-este-segredo").encode()


def _sha256(txt):
    return hashlib.sha256(txt.encode("utf-8")).hexdigest()


def _carregar_usuarios():
    raw = os.environ.get("USUARIOS_PAINEL", "") or ""
    raw = raw.strip().strip('"').strip("'").strip()
    usuarios = {}
    for par in raw.split(","):
        par = par.strip()
        if not par:
            continue
        partes = par.split(":")
        if len(partes) != 3:
            continue
        u, h, papel = partes
        papel = papel.strip().lower()
        if papel not in PAPEIS_VALIDOS:
            continue
        usuarios[u.strip().lower()] = {"hash": h.strip().lower(), "papel": papel}
    return usuarios


def _gera_token(usuario, papel):
    exp = int(time.time()) + TOKEN_HORAS * 3600
    corpo = f"{usuario}|{papel}|{exp}"
    assinatura = hmac.new(AUTH_SECRET, corpo.encode(), hashlib.sha256).hexdigest()
    bruto = f"{corpo}|{assinatura}"
    return base64.urlsafe_b64encode(bruto.encode()).decode()


class handler(BaseHTTPRequestHandler):
    def do_OPTIONS(self):
        self.send_response(204)
        self._cors()
        self.end_headers()

    def do_POST(self):
        try:
            tamanho = int(self.headers.get("Content-Length", 0) or 0)
            corpo_bruto = self.rfile.read(tamanho) if tamanho else b"{}"
            dados = json.loads(corpo_bruto or b"{}")
            usuario = str(dados.get("usuario", "")).strip().lower()
            senha = str(dados.get("senha", ""))

            usuarios = _carregar_usuarios()
            info = usuarios.get(usuario)
            if info and hmac.compare_digest(info["hash"], _sha256(senha)):
                resp = {"token": _gera_token(usuario, info["papel"]), "usuario": usuario, "papel": info["papel"]}
                self.send_response(200)
            else:
                resp = {"error": "Usuário ou senha inválidos"}
                self.send_response(401)
        except Exception as e:
            resp = {"error": str(e)}
            self.send_response(500)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self._cors()
        self.end_headers()
        self.wfile.write(json.dumps(resp, ensure_ascii=False).encode("utf-8"))

    def _cors(self):
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Headers", "Content-Type, Authorization")
        self.send_header("Access-Control-Allow-Methods", "POST, OPTIONS")
