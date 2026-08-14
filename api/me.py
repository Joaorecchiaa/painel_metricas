# -*- coding: utf-8 -*-
"""GET /api/me — confere o token salvo (Authorization: Bearer <token>) -> { logado, usuario, papel }."""
import os
import time
import hmac
import base64
import hashlib
import json
from http.server import BaseHTTPRequestHandler

PAPEIS_VALIDOS = {"admin", "elite", "sniper", "olympus"}
AUTH_SECRET = (os.environ.get("AUTH_SECRET", "") or "troque-este-segredo").encode()


def _valida_token(token):
    if not token:
        return None
    try:
        bruto = base64.urlsafe_b64decode(token.encode()).decode()
        usuario, papel, exp, assinatura = bruto.rsplit("|", 3)
        corpo = f"{usuario}|{papel}|{exp}"
        esperado = hmac.new(AUTH_SECRET, corpo.encode(), hashlib.sha256).hexdigest()
        if not hmac.compare_digest(esperado, assinatura):
            return None
        if int(exp) < int(time.time()):
            return None
        if papel not in PAPEIS_VALIDOS:
            return None
        return {"usuario": usuario, "papel": papel}
    except Exception:
        return None


class handler(BaseHTTPRequestHandler):
    def do_OPTIONS(self):
        self.send_response(204)
        self._cors()
        self.end_headers()

    def do_GET(self):
        auth = self.headers.get("Authorization") or ""
        token = auth[7:] if auth.startswith("Bearer ") else ""
        sessao = _valida_token(token)
        if sessao:
            resp = {"logado": True, "usuario": sessao["usuario"], "papel": sessao["papel"]}
        else:
            resp = {"logado": False}
        self.send_response(200)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self._cors()
        self.end_headers()
        self.wfile.write(json.dumps(resp, ensure_ascii=False).encode("utf-8"))

    def _cors(self):
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Headers", "Content-Type, Authorization")
        self.send_header("Access-Control-Allow-Methods", "GET, OPTIONS")
