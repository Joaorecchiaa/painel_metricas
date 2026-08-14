import os
import re
import csv
import io
import time
import hmac
import base64
import hashlib
import json
import unicodedata
from http.server import BaseHTTPRequestHandler

import requests

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------

PIPEDRIVE_DOMAIN = os.environ.get("PIPEDRIVE_DOMAIN", "boardacademy")
PIPEDRIVE_DOMAIN = PIPEDRIVE_DOMAIN.replace(".pipedrive.com", "").strip()
PIPEDRIVE_API_TOKEN = os.environ.get("PIPEDRIVE_API_TOKEN")

BASE_V2 = f"https://{PIPEDRIVE_DOMAIN}.pipedrive.com/api/v2"
BASE_V1 = f"https://{PIPEDRIVE_DOMAIN}.pipedrive.com/api/v1"

PIPELINES = {
    "SNIPER": 88,
    "OLYMPUS": 46,
    "ELITE": 87,
    "NAVIGATOR": 51,
    "ZENITE": 89,
}

COLLABORATORS_CSV_URL = os.environ.get("COLLABORATORS_CSV_URL") or os.environ.get(
    "COLABORADORES_CSV_URL"
)
if COLLABORATORS_CSV_URL:
    COLLABORATORS_CSV_URL = COLLABORATORS_CSV_URL.strip().strip('"').strip("'")

CARGO_KEYS = ["SDR", "CLOSER", "HEAD", "TEAM_LEADER", "OUTROS", "NAO_MAPEADO"]

STATUS_RULES = {
    "SDR": {"ok_min": 150, "ok_max": 170, "vermelho_acima": 190},
    "CLOSER": {"ok_min": 60, "ok_max": 70, "vermelho_acima": 90},
}


# ---------------------------------------------------------------------------
# Pipedrive
# ---------------------------------------------------------------------------

def _get(url, params=None):
    params = params or {}
    params["api_token"] = PIPEDRIVE_API_TOKEN
    resp = requests.get(url, params=params, timeout=30)
    resp.raise_for_status()
    return resp.json()


def get_stages_by_pipeline():
    stages_by_pipeline = {}
    cursor = None
    use_v1 = False

    while True:
        params = {"limit": 500}
        if cursor:
            params["cursor"] = cursor

        try:
            if use_v1:
                data = _get(f"{BASE_V1}/stages", params)
            else:
                data = _get(f"{BASE_V2}/stages", params)
        except Exception:
            if use_v1:
                raise
            use_v1 = True
            cursor = None
            continue

        for stage in data.get("data", []) or []:
            pid = stage.get("pipeline_id")
            stages_by_pipeline.setdefault(pid, {})[stage["id"]] = {
                "name": stage.get("name"),
                "order_nr": stage.get("order_nr"),
            }

        cursor = (data.get("additional_data", {}) or {}).get("next_cursor")
        if not cursor:
            break

    return stages_by_pipeline


def get_users():
    users = {}
    start = 0
    while True:
        data = _get(f"{BASE_V1}/users", {"start": start, "limit": 500})
        for u in data.get("data", []) or []:
            users[u["id"]] = u.get("name")
        pagination = (data.get("additional_data", {}) or {}).get("pagination", {})
        if not pagination.get("more_items_in_collection"):
            break
        start = pagination.get("next_start", start + 500)
    return users


def get_open_deals():
    all_deals = []

    for pipeline_name, pipeline_id in PIPELINES.items():
        cursor = None
        while True:
            params = {"status": "open", "pipeline_id": pipeline_id, "limit": 500}
            if cursor:
                params["cursor"] = cursor

            try:
                data = _get(f"{BASE_V2}/deals", params)
                deals = data.get("data", []) or []
                for d in deals:
                    d["_pipeline_name"] = pipeline_name
                all_deals.extend(deals)
                cursor = (data.get("additional_data", {}) or {}).get("next_cursor")
                if not cursor:
                    break
            except Exception:
                start = 0
                while True:
                    params_v1 = {
                        "status": "open",
                        "pipeline_id": pipeline_id,
                        "start": start,
                        "limit": 500,
                    }
                    data = _get(f"{BASE_V1}/deals", params_v1)
                    deals = data.get("data", []) or []
                    for d in deals:
                        d["_pipeline_name"] = pipeline_name
                    all_deals.extend(deals)
                    pagination = (data.get("additional_data", {}) or {}).get(
                        "pagination", {}
                    )
                    if not pagination.get("more_items_in_collection"):
                        break
                    start = pagination.get("next_start", start + 500)
                break

    return all_deals


# ---------------------------------------------------------------------------
# Colaboradores (planilha)
# ---------------------------------------------------------------------------

def _normalize(text):
    if not text:
        return ""
    text = text.strip().lower()
    text = "".join(
        c for c in unicodedata.normalize("NFD", text) if unicodedata.category(c) != "Mn"
    )
    return text


def _find_column(fieldnames, *candidates):
    normalized_map = {_normalize(f).replace(" ", ""): f for f in fieldnames}
    for cand in candidates:
        key = _normalize(cand).replace(" ", "")
        if key in normalized_map:
            return normalized_map[key]
    return None


def _classify_cargo(cargo_raw):
    if not cargo_raw:
        return "OUTROS"
    key = _normalize(cargo_raw)
    key = re.sub(r"\d+", "", key).strip()

    if key.startswith("sdr"):
        return "SDR"
    if key.startswith("closer"):
        return "CLOSER"
    if key.startswith("head"):
        return "HEAD"
    if key.startswith("team leader"):
        return "TEAM_LEADER"
    return "OUTROS"


def load_collaborators():
    if not COLLABORATORS_CSV_URL:
        return {}

    resp = requests.get(COLLABORATORS_CSV_URL, timeout=30)
    resp.raise_for_status()
    reader = csv.DictReader(io.StringIO(resp.text))

    col_nome = _find_column(reader.fieldnames, "Nome")
    col_subarea = _find_column(
        reader.fieldnames, "Subarea", "Sub Area", "Sub Área", "Time", "Funil"
    )
    col_cargo = _find_column(reader.fieldnames, "Cargo")

    collaborators = {}
    for row in reader:
        nome = (row.get(col_nome) or "").strip() if col_nome else ""
        if not nome:
            continue
        time_area = (row.get(col_subarea) or "").strip().upper() if col_subarea else ""
        if time_area == "MGM":
            time_area = "OLYMPUS"
        cargo_raw = (row.get(col_cargo) or "").strip() if col_cargo else ""

        collaborators[_normalize(nome)] = {
            "nome": nome,
            "time": time_area,
            "cargo_raw": cargo_raw,
            "categoria": _classify_cargo(cargo_raw),
        }
    return collaborators


def match_owner(owner_name, collaborators):
    key = _normalize(owner_name)
    if key in collaborators:
        return collaborators[key]

    parts = key.split()
    if len(parts) >= 2:
        short_key = " ".join(parts[:2])
        for k, v in collaborators.items():
            if k.startswith(short_key):
                return v

    return None


# ---------------------------------------------------------------------------
# Status (SDR/Closer)
# ---------------------------------------------------------------------------

def get_status(cargo, quantidade):
    regra = STATUS_RULES.get(cargo)
    if not regra:
        return None
    if regra["ok_min"] <= quantidade <= regra["ok_max"]:
        return {"label": "OK", "color": "green"}
    if quantidade > regra["vermelho_acima"]:
        return {"label": "Crítico", "color": "red"}
    return {"label": "Alerta", "color": "yellow"}


# ---------------------------------------------------------------------------
# Autenticação (papéis: ADMIN vê tudo; ELITE/OLYMPUS só o próprio funil)
# ---------------------------------------------------------------------------

AUTH_SECRET = (os.environ.get("AUTH_SECRET", "") or "troque-este-segredo").encode()
FUNIL_POR_PAPEL = {"ELITE": "ELITE", "OLYMPUS": "OLYMPUS"}


def valida_token(token):
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
        return {"usuario": usuario, "papel": papel}
    except Exception:
        return None


# ---------------------------------------------------------------------------
# Montagem da resposta
# ---------------------------------------------------------------------------

def build_response(apenas_funil=None):
    collaborators = load_collaborators()
    users = get_users()
    stages_by_pipeline = get_stages_by_pipeline()
    deals = get_open_deals()
    if apenas_funil:
        deals = [d for d in deals if d.get("_pipeline_name") == apenas_funil]

    result = {}
    for pipeline_name in PIPELINES:
        if apenas_funil and pipeline_name != apenas_funil:
            continue
        result[pipeline_name] = {
            "total": 0,
            "por_cargo": {k: 0 for k in CARGO_KEYS},
            "por_pessoa": {},
        }

    unmapped_owners = set()

    for deal in deals:
        pipeline_name = deal.get("_pipeline_name")

        owner_field = deal.get("owner_id")
        if isinstance(owner_field, dict):
            owner_name = owner_field.get("name")
        else:
            owner_name = users.get(owner_field)

        collaborator = match_owner(owner_name, collaborators) if owner_name else None
        cargo_key = collaborator["categoria"] if collaborator else "NAO_MAPEADO"

        stage_info = stages_by_pipeline.get(deal.get("pipeline_id"), {}).get(
            deal.get("stage_id"), {}
        )

        bucket = result[pipeline_name]
        bucket["total"] += 1
        bucket["por_cargo"][cargo_key] += 1

        pessoa_nome = owner_name or "Sem dono"
        if pessoa_nome not in bucket["por_pessoa"]:
            bucket["por_pessoa"][pessoa_nome] = {
                "cargo": cargo_key,
                "quantidade": 0,
                "negocios": [],
            }
        bucket["por_pessoa"][pessoa_nome]["quantidade"] += 1
        bucket["por_pessoa"][pessoa_nome]["negocios"].append(
            {
                "id": deal.get("id"),
                "nome": deal.get("title"),
                "valor": deal.get("value"),
                "moeda": deal.get("currency"),
                "etapa": stage_info.get("name"),
                "url": f"https://{PIPEDRIVE_DOMAIN}.pipedrive.com/deal/{deal.get('id')}",
            }
        )

        if cargo_key == "NAO_MAPEADO" and owner_name:
            unmapped_owners.add(owner_name)

    for pipeline_name in result:
        pessoas = result[pipeline_name]["por_pessoa"]
        lista = [
            {
                "nome": nome,
                "cargo": info["cargo"],
                "quantidade": info["quantidade"],
                "status": get_status(info["cargo"], info["quantidade"]),
                "negocios": sorted(info["negocios"], key=lambda d: d["nome"] or ""),
            }
            for nome, info in pessoas.items()
        ]
        lista.sort(key=lambda x: x["quantidade"], reverse=True)
        result[pipeline_name]["por_pessoa"] = lista

    resumo_equipes = [
        {"time": pipeline_name, "total": result[pipeline_name]["total"]}
        for pipeline_name in result
    ]

    # total por colaborador, somando todos os funis (guarda também o detalhe por funil)
    colaboradores_geral = {}
    for pipeline_name in result:
        for p in result[pipeline_name]["por_pessoa"]:
            nome = p["nome"]
            if nome not in colaboradores_geral:
                colaboradores_geral[nome] = {"cargo": p["cargo"], "quantidade": 0, "por_funil": {}}
            colaboradores_geral[nome]["quantidade"] += p["quantidade"]
            colaboradores_geral[nome]["por_funil"][pipeline_name] = p["quantidade"]
            if (
                colaboradores_geral[nome]["cargo"] == "NAO_MAPEADO"
                and p["cargo"] != "NAO_MAPEADO"
            ):
                colaboradores_geral[nome]["cargo"] = p["cargo"]

    total_por_colaborador = [
        {
            "nome": nome,
            "cargo": info["cargo"],
            "quantidade": info["quantidade"],
            "status": get_status(info["cargo"], info["quantidade"]),
            "por_funil": [
                {"funil": funil, "quantidade": qtd}
                for funil, qtd in sorted(
                    info["por_funil"].items(), key=lambda x: x[1], reverse=True
                )
            ],
        }
        for nome, info in colaboradores_geral.items()
    ]
    total_por_colaborador.sort(key=lambda x: x["quantidade"], reverse=True)

    return {
        "funis": result,
        "resumo_equipes": resumo_equipes,
        "total_por_colaborador": total_por_colaborador,
        "donos_nao_mapeados": sorted(unmapped_owners),
        "total_negocios": len(deals),
    }


class handler(BaseHTTPRequestHandler):
    def do_GET(self):
        auth = self.headers.get("Authorization", "")
        tok = auth[7:] if auth.startswith("Bearer ") else ""
        sessao = valida_token(tok)
        if not sessao:
            self.send_response(401)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.end_headers()
            self.wfile.write(json.dumps({"error": "não autenticado"}).encode("utf-8"))
            return

        try:
            apenas_funil = FUNIL_POR_PAPEL.get(sessao["papel"])
            data = build_response(apenas_funil)
            self.send_response(200)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Cache-Control", "no-store")
            self.end_headers()
            self.wfile.write(json.dumps(data, ensure_ascii=False).encode("utf-8"))
        except Exception as e:
            self.send_response(500)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.end_headers()
            self.wfile.write(json.dumps({"erro": str(e)}).encode("utf-8"))
